# -*- coding: utf-8 -*-
"""
Firebase Hosting Direct REST API Deployer
- 企業プロキシ・NPM依存なしで直接Firebase Hostingへデプロイ
- サービスアカウントキー (.json) を自動検出して1秒でデプロイ完了
"""
import os
import sys
import glob
import json
import gzip
import hashlib
import requests

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

from google.oauth2 import service_account
from google.auth.transport.requests import Request

PROJECT_ID = "warehouse-work-progress"
SITE_ID = "warehouse-work-progress"
SCOPES = [
    "https://www.googleapis.com/auth/firebase.hosting",
    "https://www.googleapis.com/auth/cloud-platform"
]

def find_service_account_key(base_dir):
    # 1. Look in current directory
    for f in glob.glob(os.path.join(base_dir, "*.json")):
        base = os.path.basename(f)
        if "warehouse-work-progress" in base or "firebase-adminsdk" in base:
            return f
    # 2. Look in user's Downloads
    downloads = os.path.expanduser("~/Downloads")
    for f in glob.glob(os.path.join(downloads, "*warehouse-work-progress*.json")):
        return f
    for f in glob.glob(os.path.join(downloads, "*firebase-adminsdk*.json")):
        return f
    return None

def compute_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def deploy():
    print("========================================================")
    print("  倉庫作業進捗サイネージ - Firebase Hosting 自動デプロイ")
    print(f"  プロジェクト: {PROJECT_ID}")
    print("========================================================")
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    key_file = find_service_account_key(base_dir)
    
    if not key_file:
        print("\n[!] サービスアカウント秘密鍵（.json）が見つかりません。")
        print("    Firebase コンソールの「プロジェクトの設定」>「サービス アカウント」から")
        print("    『新しい秘密鍵の生成』をクリックしてダウンロードしてください。")
        return False
        
    print(f"\n[1/5] 認証キーを検出しました: {os.path.basename(key_file)}")
    
    # 1. Authenticate with Google
    creds = service_account.Credentials.from_service_account_file(key_file, scopes=SCOPES)
    creds.refresh(Request())
    token = creds.token
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print("[2/5] Google Cloud 認証に成功しました。")
    
    # 2. Prepare files to deploy (must be gzipped for Firebase Hosting API)
    files_to_deploy = {}
    for filename in ["index.html", "style.css", "app.js"]:
        filepath = os.path.join(base_dir, filename)
        if os.path.exists(filepath):
            with open(filepath, "rb") as f:
                raw_content = f.read()
            gzipped_content = gzip.compress(raw_content)
            file_hash = compute_sha256(gzipped_content)
            files_to_deploy[f"/{filename}"] = {
                "hash": file_hash,
                "content": gzipped_content,
                "path": filepath
            }
    
    print(f"[3/5] デプロイ対象ファイル: {list(files_to_deploy.keys())}")
    
    # 3. Create a new Hosting version
    version_url = f"https://firebasehosting.googleapis.com/v1beta1/sites/{SITE_ID}/versions"
    version_config = {
        "config": {
            "rewrites": [
                {"glob": "**", "path": "/index.html"}
            ]
        }
    }
    resp = requests.post(version_url, headers=headers, json=version_config)
    if not resp.ok:
        print(f"[!] バージョン作成失敗: {resp.status_code} {resp.text}")
        return False
    version_data = resp.json()
    version_name = version_data["name"]
    print(f"    作成されたバージョン: {version_name}")
    
    # 4. Populate files (hash list)
    file_hashes = {path: info["hash"] for path, info in files_to_deploy.items()}
    populate_url = f"https://firebasehosting.googleapis.com/v1beta1/{version_name}:populateFiles"
    resp = requests.post(populate_url, headers=headers, json={"files": file_hashes})
    if not resp.ok:
        print(f"[!] ハッシュ登録失敗: {resp.status_code} {resp.text}")
        return False
    populate_data = resp.json()
    upload_url_prefix = populate_data.get("uploadUrl", "https://upload-firebasehosting.googleapis.com/upload/v1beta1")
    required_hashes = populate_data.get("uploadRequiredHashes", [])
    
    # 5. Upload required files
    print(f"[4/5] ファイルをアップロード中 ({len(required_hashes)} / {len(files_to_deploy)} ファイル)...")
    for path, info in files_to_deploy.items():
        if info["hash"] in required_hashes:
            upload_url = f"{upload_url_prefix}/{version_name}/files/{info['hash']}"
            resp = requests.post(upload_url, headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/octet-stream"
            }, data=info["content"])
            if not resp.ok:
                print(f"[!] アップロード失敗 ({path}): {resp.status_code} {resp.text}")
                return False
            print(f"    アップロード完了: {path}")
    
    # 6. Finalize version
    finalize_url = f"https://firebasehosting.googleapis.com/v1beta1/{version_name}?update_mask=status"
    resp = requests.patch(finalize_url, headers=headers, json={"status": "FINALIZED"})
    if not resp.ok:
        print(f"[!] バージョン確定失敗: {resp.status_code} {resp.text}")
        return False
        
    # 7. Release version
    release_url = f"https://firebasehosting.googleapis.com/v1beta1/sites/{SITE_ID}/releases?versionName={version_name}"
    resp = requests.post(release_url, headers=headers)
    if not resp.ok:
        print(f"[!] リリース失敗: {resp.status_code} {resp.text}")
        return False
        
    print("\n" + "=" * 60)
    print("  ★ デプロイが完了しました！")
    print(f"  公開URL: https://{SITE_ID}.web.app")
    print(f"           https://{SITE_ID}.firebaseapp.com")
    print("=" * 60)
    return True

if __name__ == "__main__":
    success = deploy()
    if not success:
        sys.exit(1)
