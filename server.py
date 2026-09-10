import os
import sys
import json
import io
import time
import datetime
import urllib.parse
import threading
import shutil
import glob
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

# Import Windows COM & Win32 File API
try:
    import win32com.client
    import win32file
    import win32con
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

import openpyxl
import hashlib

# Set output encoding to UTF-8
try:
    sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
except Exception:
    pass

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, "config.json")
PORT = 8080

DEFAULT_CONFIG = {
    "excel_path": r"%USERPROFILE%\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\平日(水～土)（本番用）\(平日)作業進捗管理データ.xlsm",
    "excel_paths": {
        "平日": r"%USERPROFILE%\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\平日(水～土)（本番用）\(平日)作業進捗管理データ.xlsm",
        "月曜": r"%USERPROFILE%\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\月曜（本番用）\(月)作業進捗管理データ.xlsm",
        "火曜": r"%USERPROFILE%\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\火曜（本番用）\(火)作業進捗管理データ.xlsm",
        "日・祝": r"%USERPROFILE%\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\日祝（本番用）\(日祝)作業進捗管理データ.xlsm"
    },
    "poll_interval_sec": 5,
    "font_size_scale": 1.0,
    "auto_scroll_speed": 40,
    "scroll_speed_px_per_sec": 50,
    "bottom_pause_sec": 4,
    "top_pause_sec": 2,
    "theme": "dark"
}

# Excel常駐監視・スケジューラー設定
EXCEL_SCHEDULE_START_HOUR = 7   # 稼働開始（朝 07:00）
EXCEL_SCHEDULE_END_HOUR = 20    # 稼働停止（夜 20:00）
EXCEL_HEALTHCHECK_INTERVAL = 15 # 死活監視間隔（秒）

DAYS_ORDER = ["平日", "月曜", "火曜", "日・祝"]

DAY_ALIASES = {
    "平日": "平日", "heijitsu": "平日", "weekday": "平日", "水": "平日", "木": "平日", "金": "平日", "土": "平日", "wed": "平日", "thu": "平日", "fri": "平日", "sat": "平日",
    "月曜": "月曜", "月": "月曜", "mon": "月曜", "monday": "月曜",
    "火曜": "火曜", "火": "火曜", "tue": "火曜", "tuesday": "火曜",
    "日・祝": "日・祝", "日祝": "日・祝", "日": "日・祝", "祝": "日・祝", "sun": "日・祝", "sunday": "日・祝", "holiday": "日・祝"
}

DAY_KEYWORD_MAP = {
    "平日": ["(平日)"],
    "月曜": ["(月)", "(月曜)"],
    "火曜": ["(火)", "(火曜)"],
    "日・祝": ["(日祝)", "(日・祝)"]
}

# In-memory thread-safe cache
MEMORY_CACHE = {}
LAST_FILE_MTIME = {}
LAST_LIVE_TIME = {}
CACHE_LOCK = threading.Lock()


def calculate_progress_score(data):
    """進捗度合いをスコアリングして古いデータ（同期遅延・巻き戻り）を検知する"""
    if not data or not isinstance(data, dict):
        return 0
    courses = data.get("courses", [])
    if not courses:
        return 0
    score = 0
    for c in courses:
        for typ in ["furidashi", "sasho"]:
            items = c.get(typ, {}).get("items", [])
            for it in items:
                st = it.get("status")
                if st == 99:
                    score += 10
                elif st == 1:
                    score += 1
    return score



def resolve_canonical_day(day_input):
    try:
        if not day_input or not str(day_input).strip():
            weekday = datetime.datetime.now().weekday()
            # Monday=0, Tuesday=1, Wednesday=2, Thursday=3, Friday=4, Saturday=5, Sunday=6
            if weekday == 6:
                return "日・祝"
            elif weekday == 0:
                return "月曜"
            elif weekday == 1:
                return "火曜"
            else:
                return "平日"

        clean = str(day_input).strip().lower()
        if clean in DAY_ALIASES:
            return DAY_ALIASES[clean]
        for k, v in DAY_ALIASES.items():
            if k.lower() == clean:
                return v
        for k, v in DAY_ALIASES.items():
            if k.lower() in clean:
                return v
        return "平日"
    except Exception:
        return "平日"


def load_config():
    try:
        if not os.path.exists(CONFIG_FILE):
            cfg = dict(DEFAULT_CONFIG)
            save_config(cfg)
            return cfg
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            merged = {**DEFAULT_CONFIG, **cfg}
            if "excel_paths" not in merged or not isinstance(merged["excel_paths"], dict):
                merged["excel_paths"] = dict(DEFAULT_CONFIG["excel_paths"])
            return merged
    except Exception:
        return dict(DEFAULT_CONFIG)


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[CONFIG SAVE ERROR]", e, flush=True)


def expand_path(p):
    """Expands %USERPROFILE%, %USERNAME%, and environment variables in path."""
    if not p or not isinstance(p, str):
        return ""
    expanded = os.path.expandvars(p.strip())
    return os.path.normpath(expanded)


def get_candidate_onedrive_roots():
    candidate_onedrive_roots = []
    user_profile = os.environ.get("USERPROFILE", "")
    onedrive_env = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
    if onedrive_env and os.path.exists(onedrive_env):
        candidate_onedrive_roots.append(onedrive_env)
    if user_profile:
        p1 = os.path.join(user_profile, "OneDrive - トヨタモビリティパーツ株式会社")
        if os.path.exists(p1) and p1 not in candidate_onedrive_roots:
            candidate_onedrive_roots.append(p1)
    for fallback in [
        r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社",
        r"c:\Users\00137184\OneDrive - トヨタモビリティパーツ株式会社"
    ]:
        if os.path.exists(fallback) and fallback not in candidate_onedrive_roots:
            candidate_onedrive_roots.append(fallback)
    return candidate_onedrive_roots


# Cache to log source transitions only when changed
LAST_LOGGED_SOURCE = {}


def find_excel_file_for_day(canonical_day):
    """
    Finds target Excel file with multi-tier automatic fallback:
      Priority 1: Dedicated Server PC (85371-butsuryupc) direct path
      Priority 2: config.json defined candidate paths (supports list & %USERPROFILE%)
      Priority 3: Auto-detected active OneDrive / Shortcuts / ★入力シート on current running PC
    """
    global LAST_LOGGED_SOURCE
    try:
        cfg = load_config()
        current_user = os.environ.get("USERNAME", "ローカルユーザー")
        keywords = DAY_KEYWORD_MAP.get(canonical_day, ["(平日)"])

        # -------------------------------------------------------------
        # Priority 1: Check Dedicated Server PC (85371-butsuryupc) paths
        # -------------------------------------------------------------
        server_pc_candidates = {
            "平日": [
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\平日(水～土)（本番用）\(平日)作業進捗管理データ.xlsm",
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\平日(水～土)（本番用）\(平日)作業進捗管理データ.xlsm"
            ],
            "月曜": [
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\月曜（本番用）\(月)作業進捗管理データ.xlsm",
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\月曜（本番用）\(月)作業進捗管理データ.xlsm"
            ],
            "火曜": [
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\火曜（本番用）\(火)作業進捗管理データ.xlsm",
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\火曜（本番用）\(火)作業進捗管理データ.xlsm"
            ],
            "日・祝": [
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\日祝（本番用）\(日祝)作業進捗管理データ.xlsm",
                r"C:\Users\85371-butsuryupc\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\日祝（本番用）\(日祝)作業進捗管理データ.xlsm"
            ]
        }

        for candidate in server_pc_candidates.get(canonical_day, []):
            p = expand_path(candidate)
            if p and os.path.exists(p):
                if LAST_LOGGED_SOURCE.get(canonical_day) != p:
                    LAST_LOGGED_SOURCE[canonical_day] = p
                    print(f"[DATA SOURCE: 優先①] [{canonical_day}] 専用サーバーPC(85371-butsuryupc)のOneDrive検出: {p}", flush=True)
                return p

        # -------------------------------------------------------------
        # Priority 1.5: Current User PC (00137184 etc.) direct path check
        # -------------------------------------------------------------
        user_pc_candidates = {
            "平日": [
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\平日(水～土)（本番用）\(平日)作業進捗管理データ.xlsm",
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\平日(水～土)（本番用）\(平日)作業進捗管理データ.xlsm"
            ],
            "月曜": [
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\月曜（本番用）\(月)作業進捗管理データ.xlsm",
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\月曜（本番用）\(月)作業進捗管理データ.xlsm"
            ],
            "火曜": [
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\火曜（本番用）\(火)作業進捗管理データ.xlsm",
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\火曜（本番用）\(火)作業進捗管理データ.xlsm"
            ],
            "日・祝": [
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\Shortcuts\新体制移行の情報共有 - ★入力シート\日祝（本番用）\(日祝)作業進捗管理データ.xlsm",
                rf"C:\Users\{current_user}\OneDrive - トヨタモビリティパーツ株式会社\新体制移行の情報共有 - ★入力シート\日祝（本番用）\(日祝)作業進捗管理データ.xlsm"
            ]
        }
        for candidate in user_pc_candidates.get(canonical_day, []):
            p = expand_path(candidate)
            if p and os.path.exists(p):
                if LAST_LOGGED_SOURCE.get(canonical_day) != p:
                    LAST_LOGGED_SOURCE[canonical_day] = p
                    print(f"[DATA SOURCE: 優先①-2] [{canonical_day}] 実行中ユーザー({current_user})のOneDrive検出: {p}", flush=True)
                return p

        # -------------------------------------------------------------
        # Priority 2: Configured paths in config.json
        # -------------------------------------------------------------
        configured_entry = cfg.get("excel_paths", {}).get(canonical_day)
        candidates_from_config = []
        if isinstance(configured_entry, list):
            candidates_from_config.extend(configured_entry)
        elif isinstance(configured_entry, str) and configured_entry.strip():
            candidates_from_config.append(configured_entry.strip())

        if canonical_day == "平日":
            legacy_p = cfg.get("excel_path", "")
            if legacy_p and legacy_p not in candidates_from_config:
                candidates_from_config.append(legacy_p)

        for raw_p in candidates_from_config:
            p = expand_path(raw_p)
            if p and os.path.exists(p):
                if LAST_LOGGED_SOURCE.get(canonical_day) != p:
                    LAST_LOGGED_SOURCE[canonical_day] = p
                    print(f"[DATA SOURCE: 優先②] [{canonical_day}] config.json指定パス使用: {p}", flush=True)
                return p

        # -------------------------------------------------------------
        # Priority 3: Automatic fallback on current PC's OneDrive / Shortcuts
        # -------------------------------------------------------------
        search_dirs = []
        candidate_onedrive_roots = get_candidate_onedrive_roots()

        for onedrive_root in candidate_onedrive_roots:
            if not os.path.exists(onedrive_root):
                continue
            try:
                for root_dir, subdirs, _ in os.walk(onedrive_root):
                    rel = os.path.relpath(root_dir, onedrive_root)
                    depth = len(rel.split(os.sep)) if rel != "." else 0
                    if depth <= 3:
                        if root_dir not in search_dirs:
                            if "入力シート" in root_dir or "新体制" in root_dir:
                                search_dirs.insert(0, root_dir)
                            else:
                                search_dirs.append(root_dir)
                    else:
                        subdirs.clear()
            except Exception:
                pass

        matched_candidates = []

        for s_dir in search_dirs:
            if not s_dir or not os.path.exists(s_dir):
                continue
            try:
                for fname in os.listdir(s_dir):
                    if fname.endswith((".xlsx", ".xlsm")) and not fname.startswith("~$"):
                        if any(ng in fname for ng in ["トライアル", "テスト", "パワポ表示用", "コピー"]):
                            continue
                        if canonical_day == "月曜" and "(平日)" in fname:
                            continue
                        if any(kw in fname for kw in keywords):
                            full_file = os.path.join(s_dir, fname)
                            mtime = os.path.getmtime(full_file)
                            is_official = ("作業進捗管理データ" in fname)
                            matched_candidates.append((is_official, mtime, full_file))
            except Exception:
                pass

        if matched_candidates:
            matched_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
            chosen = matched_candidates[0][2]
            if LAST_LOGGED_SOURCE.get(canonical_day) != chosen:
                LAST_LOGGED_SOURCE[canonical_day] = chosen
                print(f"[PATH FALLBACK: 優先③] [{canonical_day}] サーバーPC未検出のため、{current_user}のOneDriveを自動検知して使用: {chosen}", flush=True)
            return chosen

        return ""
    except Exception as e:
        print(f"[PATH ERROR] [{canonical_day}]: {e}", flush=True)
        return ""


def read_locked_file_bytes(filepath):
    """Reads bytes from a file even if locked by Excel for writing."""
    if HAS_WIN32:
        try:
            handle = win32file.CreateFile(
                filepath,
                win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None
            )
            file_size = win32file.GetFileSize(handle)
            _, data = win32file.ReadFile(handle, file_size)
            win32file.CloseHandle(handle)
            return data
        except Exception:
            pass
    try:
        with open(filepath, "rb") as f:
            return f.read()
    except Exception:
        return None


# =========================================================
# バックグラウンド非表示Excel自動管理エンジン (Headless Engine)
# =========================================================
class HeadlessExcelEngine:
    """バックグラウンドで非表示Excelを常駐させ、高速に最新データを取得するエンジン"""
    def __init__(self):
        self.lock = threading.Lock()
        self.xl = None
        self.workbooks = {}
        self.last_opened_paths = {}

    def ensure_excel_app(self):
        if not HAS_WIN32:
            return None
        if self.xl is not None:
            try:
                _ = self.xl.Visible
                return self.xl
            except Exception:
                self.cleanup()
        try:
            pythoncom.CoInitialize()
            self.xl = win32com.client.DispatchEx("Excel.Application")
            self.xl.Visible = False
            self.xl.DisplayAlerts = False
            self.xl.EnableEvents = False
            self.xl.ScreenUpdating = False
            self.xl.AskToUpdateLinks = False
            return self.xl
        except Exception as e:
            self.xl = None
            return None

    def get_data(self, target_excel_path, canonical_day):
        if not HAS_WIN32:
            return None
        with self.lock:
            try:
                pythoncom.CoInitialize()
                xl = self.ensure_excel_app()
                if not xl:
                    return None

                wb = self.workbooks.get(canonical_day)
                last_p = self.last_opened_paths.get(canonical_day)

                needs_open = False
                if wb is None or last_p != target_excel_path:
                    needs_open = True
                else:
                    try:
                        _ = wb.Name
                    except Exception:
                        needs_open = True

                if needs_open:
                    if wb is not None:
                        try:
                            wb.Close(False)
                        except Exception:
                            pass
                    if not os.path.exists(target_excel_path):
                        return None
                    try:
                        wb = xl.Workbooks.Open(target_excel_path, ReadOnly=True, UpdateLinks=0)
                        self.workbooks[canonical_day] = wb
                        self.last_opened_paths[canonical_day] = target_excel_path
                    except Exception:
                        return None

                # 最新のクラウド再読込＆再計算
                try:
                    wb.UpdateFromFile()
                except Exception:
                    pass

                try:
                    wb.Calculate()
                except Exception:
                    pass

                ws_disp = None
                ws_data = None
                for s in wb.Worksheets:
                    if "表示" in s.Name:
                        ws_disp = s
                    elif "データ" in s.Name:
                        ws_data = s
                if ws_disp is None:
                    ws_disp = wb.Worksheets(1)
                if ws_data is None:
                    ws_data = ws_disp

                disp_range = ws_disp.Range("A1:CZ145").Value
                data_range = ws_data.Range("A1:CZ145").Value

                disp_rows = [list(r) for r in disp_range]
                data_rows = [list(r) for r in data_range]

                return {
                    "title": ws_disp.Name,
                    "disp_rows": disp_rows,
                    "data_rows": data_rows,
                    "last_modified": datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S") + " (Auto Headless)"
                }
            except Exception as e:
                self.cleanup()
                return None

    def cleanup(self):
        for wb in self.workbooks.values():
            try:
                wb.Close(False)
            except Exception:
                pass
        self.workbooks.clear()
        self.last_opened_paths.clear()
        if self.xl is not None:
            try:
                self.xl.Quit()
            except Exception:
                pass
            self.xl = None


HEADLESS_ENGINE = HeadlessExcelEngine()
import atexit
atexit.register(HEADLESS_ENGINE.cleanup)


def get_user_open_excel_rows(target_excel_path, canonical_day):
    """
    ユーザーが画面上で実際に開いているExcel（Visible == True）から
    リアルタイムの生データを直接取得するエンジン（全76コース・保存前でも即座に反映）。
    ※非表示の常駐Excel（ゾンビ）は絶対に起動しません。
    """
    if not HAS_WIN32:
        return None

    try:
        pythoncom.CoInitialize()
        xl_app = win32com.client.GetActiveObject("Excel.Application")
        if not xl_app:
            return None

        # 画面上に表示されているExcelのみを対象にする
        try:
            if not xl_app.Visible:
                return None
        except Exception:
            return None

        wb = None
        target_name = os.path.basename(target_excel_path).lower()
        keywords = DAY_KEYWORD_MAP.get(canonical_day, [])

        for w in xl_app.Workbooks:
            try:
                if w.FullName.lower() == target_excel_path.lower() or w.Name.lower() == target_name:
                    wb = w
                    break
            except Exception:
                pass

        if wb is None:
            for w in xl_app.Workbooks:
                try:
                    w_name = w.Name.lower()
                    if any(kw.lower() in w_name for kw in keywords):
                        wb = w
                        break
                except Exception:
                    pass

        if wb is not None:
            ws_disp = None
            ws_data = None
            for s in wb.Worksheets:
                if "表示" in s.Name:
                    ws_disp = s
                elif "データ" in s.Name:
                    ws_data = s
            if ws_disp is None:
                ws_disp = wb.Worksheets(1)
            if ws_data is None:
                ws_data = ws_disp

            # 全行（76コース以上、300行まで）を取得
            disp_range = ws_disp.Range("A1:CZ300").Value
            data_range = ws_data.Range("A1:CZ300").Value

            disp_rows = [list(r) for r in disp_range]
            data_rows = [list(r) for r in data_range]

            now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            return {
                "title": ws_disp.Name,
                "disp_rows": disp_rows,
                "data_rows": data_rows,
                "last_modified": f"{now_str} (画面Excel)"
            }
    except Exception:
        pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return None


def get_live_com_rows(target_excel_path, canonical_day):
    return get_user_open_excel_rows(target_excel_path, canonical_day)


def format_cell_name(val):
    if val is None:
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    if isinstance(val, int):
        return str(val)
    s = str(val).strip()
    if s.endswith(".0"):
        try:
            float(s)
            s = s[:-2]
        except ValueError:
            pass
    return s


def format_cell_time(time_val):
    if time_val is None:
        return ""
    if isinstance(time_val, (datetime.time, datetime.datetime)):
        return time_val.strftime("%H:%M")
    if isinstance(time_val, (int, float)):
        total_minutes = int(round(time_val * 24 * 60))
        hours = (total_minutes // 60) % 24
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"
    s = str(time_val).strip()
    if " " in s:
        s = s.split(" ")[-1]
    if len(s) >= 4 and ":" in s:
        parts = s.split(":")
        try:
            return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
        except ValueError:
            return s
    return s


def parse_rows_into_signage_data(canonical_day, excel_path, disp_rows, data_rows, last_mod):
    courses = []
    last_vehicle = ""

    for r_idx in range(2, len(disp_rows), 2):
        r1 = disp_rows[r_idx]
        r2 = disp_rows[r_idx + 1] if r_idx + 1 < len(disp_rows) else [None] * 30

        v_val = r1[0] if len(r1) > 0 and r1[0] is not None else None
        c_val = r1[1] if len(r1) > 1 and r1[1] is not None else None
        time_val = r1[2] if len(r1) > 2 and r1[2] is not None else None

        if v_val is None and c_val is None and time_val is None:
            continue

        v_str = format_cell_name(v_val)
        if v_str != "":
            last_vehicle = v_str

        course_name = format_cell_name(c_val)
        time_str = format_cell_time(time_val)

        data_r1 = data_rows[r_idx] if r_idx < len(data_rows) else [None] * 105
        data_r2 = data_rows[r_idx + 1] if r_idx + 1 < len(data_rows) else [None] * 105

        # 1 & 2. 振出・査照のタクト判定:
        # Excelの G..AE列（1〜25タクト）の「パワポ表示用変換」値を完全評価
        # （データシート上でエンドカード99が打たれた場合、手前のタクトはすべて99完了となり青色表示される）
        def eval_row_tacts(data_row):
            # AG..BE列 (インデックス 32..56) の現場入力値を取得
            ag_be = []
            for i in range(25):
                c = 32 + i
                v = data_row[c] if len(data_row) > c and data_row[c] is not None else 0
                try:
                    v_num = int(float(v))
                except (ValueError, TypeError):
                    v_num = 0
                ag_be.append(v_num if v_num in (99, 1) else 0)

            # エンドカード(99)が存在する場合、そのタクトおよび手前はすべて99(完了/青)
            if 99 in ag_be:
                match_99 = ag_be.index(99)
                result = []
                for idx in range(25):
                    if idx <= match_99:
                        result.append(99)
                    else:
                        result.append(ag_be[idx])
                return [{"num": i + 1, "status": result[i]} for i in range(25)]
            else:
                return [{"num": i + 1, "status": ag_be[i]} for i in range(25)]

        furidashi_items = eval_row_tacts(data_r1)
        sagyo_items = eval_row_tacts(data_r2)

        # 3. 伝票: データシート CH列 (インデックス 85) が 1
        slip_val = data_r1[85] if len(data_r1) > 85 and data_r1[85] is not None else 0
        is_slip_done = (slip_val in (1, "1", 1.0) or slip_val is True)

        # 4. 配送なし (No Delivery) の判定
        # 条件: データシートの (G～AE列が0 かつ AF列が99) または 表示シートの (AF列がTrue)
        all_furidashi_zero = all(item["status"] == 0 for item in furidashi_items)
        data_af_val = data_r1[31] if len(data_r1) > 31 else None
        disp_af_val = r1[31] if len(r1) > 31 else None

        is_data_af_99 = (data_af_val in (99, "99", 99.0) or str(data_af_val).strip() == "99")
        is_disp_af_true = (disp_af_val is True or str(disp_af_val).strip().upper() in ("TRUE", "1"))

        is_no_delivery = (all_furidashi_zero and is_data_af_99) or is_disp_af_true

        # 5. コース完了判定 (データシート CK列=1 または 配送なし)
        is_course_completed = False
        if len(data_r1) > 88 and data_r1[88] is not None:
            ck_val = str(data_r1[88]).strip()
            if ck_val in ("1", "1.0") or data_r1[88] == 1 or data_r1[88] is True:
                is_course_completed = True
        if is_no_delivery:
            is_course_completed = True

        # 6. 集約コース完了時間 (CM列: インデックス 90) & 時間差 (CN列: インデックス 91)
        cm_val = data_r1[90] if len(data_r1) > 90 else None
        cn_val = data_r1[91] if len(data_r1) > 91 else None
        group_completed_time = format_cell_time(cm_val)

        group_diff_min = None
        if cn_val is not None and str(cn_val).strip() != "":
            try:
                group_diff_min = int(round(float(cn_val)))
            except (ValueError, TypeError):
                group_diff_min = None

        course_id = f"course_{r_idx//2 + 1}"

        courses.append({
            "id": course_id,
            "row_index": r_idx + 1,
            "vehicle": last_vehicle,
            "course": course_name,
            "time": time_str,
            "is_completed": is_course_completed,
            "is_no_delivery": is_no_delivery,
            "group_completed_time": group_completed_time,
            "group_diff_minutes": group_diff_min,
            "furidashi": {
                "label": "振出",
                "items": furidashi_items
            },
            "sagyo": {
                "label": "査照",
                "items": sagyo_items
            },
            "slip": {
                "label": "伝票",
                "is_done": is_slip_done
            }
        })

    return {
        "success": True,
        "day": canonical_day,
        "title": canonical_day,
        "excel_file": os.path.basename(excel_path),
        "count": len(courses),
        "last_modified": last_mod,
        "courses": courses
    }


def refresh_data_for_day(canonical_day):
    """Background update for a single day. Returns data object with stale protection."""
    global LAST_LIVE_TIME
    try:
        excel_path = find_excel_file_for_day(canonical_day)
        if not excel_path or not os.path.exists(excel_path):
            with CACHE_LOCK:
                cached = MEMORY_CACHE.get(canonical_day)
                if cached:
                    return cached
            return {
                "success": False,
                "day": canonical_day,
                "error": f"「{canonical_day}」のExcelファイルが見つかりません"
            }

        # 1. 【最優先】ユーザーが画面上で開いているExcel(Visible==True)からLive RAMデータを取得
        # 編集中の変更や未保存データ、Teamsリアルタイム共同編集内容を0秒で即座に反映
        live_result = get_user_open_excel_rows(excel_path, canonical_day)
        if live_result and live_result.get("disp_rows"):
            try:
                parsed_data = parse_rows_into_signage_data(
                    canonical_day,
                    excel_path,
                    live_result["disp_rows"],
                    live_result["data_rows"],
                    live_result["last_modified"]
                )
                with CACHE_LOCK:
                    cached = MEMORY_CACHE.get(canonical_day)
                    cached_score = calculate_progress_score(cached)
                    new_score = calculate_progress_score(parsed_data)

                    # 一括リセット判定（QRデータ削除等により全タクトがクリアされ、スコアが0または極小になった状態）
                    is_full_reset = (new_score == 0) or (new_score <= 10 and cached_score >= 50)

                    # Liveデータが進捗後退していないか（または初回、または一括リセット）
                    if not cached or new_score >= cached_score or is_full_reset:
                        if is_full_reset and cached_score > 0:
                            print(f"[RESET DETECTED: 画面Excel] [{canonical_day}] 一括リセット（QRデータ削除）を検知しました。画面を初期化します (旧スコア:{cached_score} -> 新スコア:{new_score})", flush=True)
                        LAST_FILE_MTIME[canonical_day] = -1.0
                        LAST_LIVE_TIME[canonical_day] = time.time()
                        MEMORY_CACHE[canonical_day] = parsed_data
                        return parsed_data
                    else:
                        # 万一Liveデータでも一時的なゴミデータが入った場合は最新キャッシュを維持
                        return cached
            except Exception as e:
                print(f"[LIVE COM PARSE ERROR] [{canonical_day}]: {e}", flush=True)

        # 2. 画面上のExcelからLiveデータが取れなかった場合
        # 直近（60秒以内）にLiveデータを取得しており、かつディスクファイルがそれ以降に保存更新されていない場合、
        # ディスクファイルは確実にLiveデータより古い（未保存・同期前）ため、キャッシュを維持してディスクを読まない！
        try:
            mtime = os.path.getmtime(excel_path)
        except Exception:
            mtime = 0

        with CACHE_LOCK:
            cached = MEMORY_CACHE.get(canonical_day)
            last_live = LAST_LIVE_TIME.get(canonical_day, 0)
            if cached and cached.get("success") and last_live > 0 and (time.time() - last_live) < 60 and mtime <= last_live:
                return cached

            last_mtime = LAST_FILE_MTIME.get(canonical_day)
            # ファイルに変更がなく、すでにキャッシュがある場合は即時返却
            if cached and cached.get("success") and last_mtime == mtime and mtime > 0:
                return cached

        # 3. ディスクファイルから直接読み込み (openpyxl)
        file_bytes = read_locked_file_bytes(excel_path)
        if file_bytes is not None:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            sheetnames = wb.sheetnames
            ws_disp = None
            ws_data = None
            for s in sheetnames:
                if "表示" in s:
                    ws_disp = wb[s]
                elif "データ" in s:
                    ws_data = wb[s]
            if ws_disp is None:
                ws_disp = wb[sheetnames[0]]
            if ws_data is None:
                ws_data = wb[sheetnames[1]] if len(sheetnames) > 1 else ws_disp

            disp_rows = list(ws_disp.iter_rows(values_only=True))
            data_rows = list(ws_data.iter_rows(values_only=True))
            last_mod = datetime.datetime.fromtimestamp(mtime).strftime("%Y/%m/%d %H:%M:%S")

            parsed_data = parse_rows_into_signage_data(canonical_day, excel_path, disp_rows, data_rows, last_mod)
            with CACHE_LOCK:
                cached = MEMORY_CACHE.get(canonical_day)
                cached_score = calculate_progress_score(cached)
                new_score = calculate_progress_score(parsed_data)

                # 一括リセット判定（QRデータ削除等により全タクトがクリアされ、スコアが0または極小になった状態）
                is_full_reset = (new_score == 0) or (new_score <= 10 and cached_score >= 50)

                # 【古いファイル拾い防止ガード】
                # 現在のキャッシュが進捗を持っており、読み込んだディスクデータの進捗スコアが後退している場合、
                # （ただし一括リセット時、またはファイル自体が新しく保存・更新された場合は正当な修正として受入）
                cached_mtime = LAST_FILE_MTIME.get(canonical_day, 0)
                is_newer_file = (mtime > cached_mtime and mtime > 0 and cached_mtime > 0)

                if cached and cached.get("success") and cached_score > 0 and new_score < cached_score:
                    if is_full_reset:
                        print(f"[RESET DETECTED: ディスク] [{canonical_day}] 一括リセット（QRデータ削除）を検知しました。画面を初期化します (旧スコア:{cached_score} -> 新スコア:{new_score})", flush=True)
                    elif is_newer_file:
                        print(f"[STALE GUARD: 更新受入] [{canonical_day}] ファイル更新を検知したため進捗修正（スコア減少）を適用しました (現スコア:{cached_score} -> 新スコア:{new_score})", flush=True)
                    else:
                        cached_mod = cached.get("last_modified", "")
                        today_str = datetime.datetime.now().strftime("%Y/%m/%d")
                        if today_str in cached_mod or "画面Excel" in cached_mod:
                            print(f"[STALE GUARD: 抑止] [{canonical_day}] 古いファイルの読み込みをブロックしました (現スコア:{cached_score} > 読込スコア:{new_score})", flush=True)
                            return cached

                LAST_FILE_MTIME[canonical_day] = mtime
                MEMORY_CACHE[canonical_day] = parsed_data
            return parsed_data

        # Fallback to cached if available
        with CACHE_LOCK:
            cached = MEMORY_CACHE.get(canonical_day)
            if cached:
                return cached
        return {
            "success": False,
            "day": canonical_day,
            "error": f"ファイル読込待機中: {os.path.basename(excel_path)}"
        }
    except Exception as e:
        with CACHE_LOCK:
            cached = MEMORY_CACHE.get(canonical_day)
            if cached:
                return cached
        return {"success": False, "day": canonical_day, "error": f"エクセル解析エラー: {str(e)}"}


# =========================================================
# Firebase Firestore Cloud Sync Engine
# =========================================================
FIREBASE_PROJECT_ID = "warehouse-work-progress"
FIREBASE_CREDS = None
FIREBASE_TOKEN = None
FIREBASE_TOKEN_EXPIRY = 0
LAST_CLOUD_SYNC_TIME = 0
LAST_CLOUD_SYNC_HASH = ""

def find_firebase_credentials():
    candidates = [
        os.path.join(APP_DIR, "firebase_credentials.json"),
        os.path.join(APP_DIR, "warehouse-work-progress-firebase-adminsdk-fbsvc-8bf0b53b80.json"),
        os.path.expanduser("~/Downloads/warehouse-work-progress-firebase-adminsdk-fbsvc-8bf0b53b80.json")
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    for f in glob.glob(os.path.join(APP_DIR, "*adminsdk*.json")):
        return f
    for f in glob.glob(os.path.expanduser("~/Downloads/*warehouse-work-progress*.json")):
        return f
    return None

def sync_to_firestore_cloud(all_days_data, cfg=None):
    """Pushes current memory cache to Firestore live document so web users stay updated."""
    global FIREBASE_CREDS, FIREBASE_TOKEN, FIREBASE_TOKEN_EXPIRY, LAST_CLOUD_SYNC_TIME, LAST_CLOUD_SYNC_HASH
    try:
        from google.oauth2 import service_account
        from google.auth.transport.requests import Request
        import requests

        key_path = find_firebase_credentials()
        if not key_path:
            return

        now = time.time()
        days_hash = hashlib.md5(json.dumps(all_days_data, sort_keys=True, default=str).encode("utf-8")).hexdigest()
        
        # 変化がない場合は30秒に1回の定時ハートビート
        if days_hash == LAST_CLOUD_SYNC_HASH and (now - LAST_CLOUD_SYNC_TIME) < 30:
            return
            return

        if not FIREBASE_TOKEN or now >= FIREBASE_TOKEN_EXPIRY:
            FIREBASE_CREDS = service_account.Credentials.from_service_account_file(
                key_path, scopes=["https://www.googleapis.com/auth/datastore"]
            )
            FIREBASE_CREDS.refresh(Request())
            FIREBASE_TOKEN = FIREBASE_CREDS.token
            FIREBASE_TOKEN_EXPIRY = now + 3000

        now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        full_payload = {
            "days": all_days_data,
            "timestamp": now_str,
            "config": cfg or load_config()
        }

        doc_url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT_ID}/databases/(default)/documents/signage_data/live"
        body = {
            "fields": {
                "payload": {
                    "stringValue": json.dumps(full_payload, ensure_ascii=False)
                },
                "updated_at": {
                    "stringValue": now_str
                }
            }
        }
        headers = {
            "Authorization": f"Bearer {FIREBASE_TOKEN}",
            "Content-Type": "application/json"
        }
        resp = requests.patch(doc_url, headers=headers, json=body, timeout=5)
        if resp.ok:
            LAST_CLOUD_SYNC_TIME = now
            LAST_CLOUD_SYNC_HASH = days_hash
            print(f"[CLOUD SYNC] 最新データをFirestoreへ送信しました ({now_str})", flush=True)
    except Exception as e:
        pass


def cleanup_zombie_excel():
    """ウィンドウを持たない残留・ゾンビ化した非表示EXCELプロセスを安全に検知して終了する"""
    try:
        # MainWindowTitle が空の孤立EXCEL.EXEを強制終了
        cmd = 'powershell -NoProfile -Command "Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } | Stop-Process -Force -ErrorAction SilentlyContinue"'
        os.system(cmd)
    except Exception:
        pass


def launch_minimized_excel(target_excel_path):
    """ExcelをVisible=Trueかつ最小化（タスクバー格納）で安全に起動・開く"""
    if not target_excel_path or not os.path.exists(target_excel_path):
        return False
    if not HAS_WIN32:
        return False

    try:
        pythoncom.CoInitialize()
        xl_app = None
        try:
            xl_app = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            try:
                xl_app = win32com.client.Dispatch("Excel.Application")
            except Exception:
                xl_app = None

        if not xl_app:
            norm_p = os.path.normpath(target_excel_path)
            os.system(f'start /min "" excel.exe "{norm_p}"')
            time.sleep(3)
            try:
                xl_app = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                pass

        if xl_app:
            xl_app.Visible = True
            xl_app.DisplayAlerts = False
            try:
                xl_app.WindowState = -4140  # win32con.xlMinimized (-4140: 最小化)
            except Exception:
                pass

            target_name = os.path.basename(target_excel_path).lower()
            is_already_open = False
            for wb in xl_app.Workbooks:
                try:
                    if wb.FullName.lower() == target_excel_path.lower() or wb.Name.lower() == target_name:
                        is_already_open = True
                        break
                except Exception:
                    pass

            if not is_already_open:
                print(f"[SUPERVISOR] 対象Excelを最小化で自動起動します: {target_excel_path}", flush=True)
                try:
                    xl_app.Workbooks.Open(target_excel_path)
                    try:
                        xl_app.WindowState = -4140  # 確実に最小化を維持
                    except Exception:
                        pass
                except Exception as open_err:
                    err_str = str(open_err)
                    # 同名ブックが別プロセスでロックされている場合、ゾンビを掃除してリトライ
                    if "使用されています" in err_str or "同じ名前のブック" in err_str or "-2146827284" in err_str:
                        print(f"[SUPERVISOR WARN] 同名ブックの排他ロックを検知しました。孤立ゾンビExcelをクリーンアップします...", flush=True)
                        cleanup_zombie_excel()
                        time.sleep(2)
                    raise open_err

            xl_app.DisplayAlerts = True
            return True
    except Exception as e:
        print(f"[SUPERVISOR LAUNCH ERROR] {e}", flush=True)
        return False
    return False


def close_excel_safely():
    """停止時間帯（夜20:00）にExcelを自動保存して安全にクローズする"""
    if not HAS_WIN32:
        return
    try:
        pythoncom.CoInitialize()
        try:
            xl_app = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return

        if xl_app:
            xl_app.DisplayAlerts = False
            for wb in list(xl_app.Workbooks):
                try:
                    wb.Save()
                except Exception:
                    pass
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            try:
                xl_app.Quit()
            except Exception:
                pass
            print("[SUPERVISOR] 夜間停止時刻(20:00)のため、Excelを自動保存して正常終了しました。", flush=True)
    except Exception as e:
        print(f"[SUPERVISOR CLOSE ERROR] {e}", flush=True)


def is_excel_running(target_excel_path, canonical_day):
    """対象Excelが現在Windows上で起動・生存しているか確認する"""
    if not HAS_WIN32:
        return False
    try:
        pythoncom.CoInitialize()
        xl_app = win32com.client.GetActiveObject("Excel.Application")
        if not xl_app:
            return False

        target_name = os.path.basename(target_excel_path).lower() if target_excel_path else ""
        keywords = DAY_KEYWORD_MAP.get(canonical_day, [])

        for wb in xl_app.Workbooks:
            try:
                matched = False
                if target_excel_path and wb.FullName.lower() == target_excel_path.lower():
                    matched = True
                elif target_name and wb.Name.lower() == target_name:
                    matched = True
                elif any(kw.lower() in wb.Name.lower() for kw in keywords):
                    matched = True

                if matched:
                    # 非表示になっていても生存していれば表示・最小化状態に復帰
                    try:
                        if not xl_app.Visible:
                            xl_app.Visible = True
                            xl_app.WindowState = -4140
                    except Exception:
                        pass
                    return True
            except Exception:
                pass
        return False
    except Exception:
        return False


def excel_supervisor_worker():
    """
    Excel常駐監視・スケジューラーデーモンスレッド:
      1. 朝 07:00 〜 夜 20:00:
         - 本日の曜日Excelを特定し、最小化(WindowState=-4140)で常駐
         - 15秒おきに死活監視を行い、手動終了やクラッシュを検知したら直ちに自動再起動
      2. 夜 20:00 〜 翌朝 07:00:
         - Excelが開かれていれば自動保存して安全にクローズ
         - 停止時間帯は再起動を行わずに待機
    """
    if not HAS_WIN32:
        print("[SUPERVISOR] Win32モジュールが無いため常駐監視はスキップされます。", flush=True)
        return

    # 起動前に残留している非表示ゾンビExcelを安全に一掃
    cleanup_zombie_excel()

    print("[SUPERVISOR] Excel自動常駐・スケジューラー（朝7:00起動/夜20:00停止/死活監視/ゾンビ保護）が稼働開始しました。", flush=True)
    last_state = None
    consecutive_errors = 0

    while True:
        try:
            now = datetime.datetime.now()
            current_hour = now.hour
            is_active_hours = (EXCEL_SCHEDULE_START_HOUR <= current_hour < EXCEL_SCHEDULE_END_HOUR)

            if is_active_hours:
                # --- 稼働時間帯 (07:00 〜 20:00) ---
                canonical_day = resolve_canonical_day("")
                target_file = find_excel_file_for_day(canonical_day)

                if target_file and os.path.exists(target_file):
                    running = is_excel_running(target_file, canonical_day)
                    if not running:
                        if last_state != "running":
                            print(f"[SUPERVISOR] 朝の起動時刻(07:00)または初回起動を検知しました: [{canonical_day}]", flush=True)
                        else:
                            print(f"[SUPERVISOR WARN] Excelの停止・クラッシュを検知しました。自動再起動します: [{canonical_day}]", flush=True)

                        ok = launch_minimized_excel(target_file)
                        if ok:
                            last_state = "running"
                            consecutive_errors = 0
                        else:
                            consecutive_errors += 1
                            if consecutive_errors >= 3:
                                # 連続失敗時はゾンビExcelを排除し、ログ溢れ防止のため長めに待機
                                cleanup_zombie_excel()
                                time.sleep(30)
                                consecutive_errors = 0
                    else:
                        last_state = "running"
                        consecutive_errors = 0
            else:
                # --- 停止時間帯 (20:00 〜 翌朝 07:00) ---
                if last_state != "stopped":
                    close_excel_safely()
                    last_state = "stopped"
                    consecutive_errors = 0

        except Exception as e:
            print(f"[SUPERVISOR EXCEPTION] {e}", flush=True)

        time.sleep(EXCEL_HEALTHCHECK_INTERVAL)


def background_cache_worker():
    """Background polling daemon thread that keeps in-memory data fresh continuously."""
    print("[CACHE ENGINE] 高速インメモリ・キャッシュエンジンが起動しました。", flush=True)

    # Initial load
    for day in DAYS_ORDER:
        try:
            data = refresh_data_for_day(day)
            with CACHE_LOCK:
                MEMORY_CACHE[day] = data
            print(f"[CACHE READY] [{day}] -> {data.get('count', 0)} コース ({data.get('excel_file', '')})", flush=True)
        except Exception as e:
            print(f"[CACHE ERROR] [{day}]: {e}", flush=True)
        time.sleep(0.1)

    # Initial Firestore Cloud sync
    with CACHE_LOCK:
        init_cache_copy = dict(MEMORY_CACHE)
    sync_to_firestore_cloud(init_cache_copy)

    while True:
        try:
            time.sleep(2)
            for day in DAYS_ORDER:
                try:
                    data = refresh_data_for_day(day)
                    with CACHE_LOCK:
                        MEMORY_CACHE[day] = data
                except Exception:
                    pass
                time.sleep(0.1)

            # Sync to Firestore Cloud
            with CACHE_LOCK:
                cache_copy = dict(MEMORY_CACHE)
            sync_to_firestore_cloud(cache_copy)
        except Exception:
            time.sleep(2)


def get_static_version():
    """app.js, style.css, index.html の最終更新日時(mtime)からバージョン文字列を生成"""
    try:
        targets = ["app.js", "style.css", "index.html"]
        latest_mtime = 0
        for name in targets:
            p = os.path.join(APP_DIR, name)
            if os.path.exists(p):
                mt = os.path.getmtime(p)
                if mt > latest_mtime:
                    latest_mtime = mt
        if latest_mtime > 0:
            dt = datetime.datetime.fromtimestamp(latest_mtime)
            return dt.strftime("%Y%m%d_%H%M%S")
    except Exception:
        pass
    return "default_v1"


class SignageRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=APP_DIR, **kwargs)

    def log_message(self, format, *args):
        pass

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        try:
            parsed = urllib.parse.urlparse(self.path)
            if parsed.path == "/api/status":
                static_ver = get_static_version()
                resp_json = json.dumps({
                    "success": True,
                    "static_version": static_ver,
                    "server_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }, ensure_ascii=False)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                self.send_header("Pragma", "no-cache")
                self.send_header("Expires", "0")
                self.end_headers()
                self.wfile.write(resp_json.encode("utf-8"))
                return

            if parsed.path == "/api/data":
                query = urllib.parse.parse_qs(parsed.query)
                day_param = query.get("day", [""])[0]
                canonical_day = resolve_canonical_day(day_param)

                # Instant in-memory cache lookup (< 1ms)
                with CACHE_LOCK:
                    data = MEMORY_CACHE.get(canonical_day)

                if not data:
                    data = {
                        "success": True,
                        "day": canonical_day,
                        "title": canonical_day,
                        "count": 0,
                        "courses": [],
                        "last_modified": "初期化中..."
                    }

                cfg = load_config()
                resp_data = dict(data)
                resp_data["config"] = cfg
                resp_data["static_version"] = get_static_version()

                resp_json = json.dumps(resp_data, ensure_ascii=False)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                self.send_header("Pragma", "no-cache")
                self.send_header("Expires", "0")
                self.end_headers()
                self.wfile.write(resp_json.encode("utf-8"))
                return

            if parsed.path == "/api/config":
                cfg = load_config()
                resp_json = json.dumps(cfg, ensure_ascii=False)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                self.send_header("Pragma", "no-cache")
                self.send_header("Expires", "0")
                self.end_headers()
                self.wfile.write(resp_json.encode("utf-8"))
                return

            super().do_GET()
        except Exception as e:
            try:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode("utf-8"))
            except Exception:
                pass

    def do_POST(self):
        try:
            parsed = urllib.parse.urlparse(self.path)
            if parsed.path == "/api/config":
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                try:
                    new_cfg = json.loads(body.decode("utf-8"))
                    current_cfg = load_config()
                    updated_cfg = {**current_cfg, **new_cfg}
                    save_config(updated_cfg)

                    self.send_response(200)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": True, "config": updated_cfg}).encode("utf-8"))
                except Exception as e:
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode("utf-8"))
                return

            self.send_response(404)
            self.end_headers()
        except Exception:
            pass


def main():
    cfg = load_config()
    port = int(cfg.get("port", PORT))
    print("=" * 60, flush=True)
    print("  倉庫作業進捗サイネージシステム (高速インメモリ・複数曜日対応)", flush=True)
    print(f"  起動URL: http://localhost:{port}", flush=True)
    print("  ※この画面を閉じるとサイネージが停止します", flush=True)
    print("=" * 60, flush=True)

    try:
        ThreadingHTTPServer.allow_reuse_address = True
        server = ThreadingHTTPServer(("", port), SignageRequestHandler)
        server.daemon_threads = True
        print(f"[SERVER] サーバーがポート {port} で正常に起動しました (即時受付可能)。", flush=True)
    except Exception as e:
        print(f"[SERVER FATAL ERROR] {e}", flush=True)
        return

    # Start background polling cache thread
    cache_thread = threading.Thread(target=background_cache_worker, daemon=True)
    cache_thread.start()

    # Start Excel supervisor daemon thread (scheduler & health-check)
    supervisor_thread = threading.Thread(target=excel_supervisor_worker, daemon=True)
    supervisor_thread.start()

    while True:
        try:
            server.serve_forever()
        except Exception as e:
            print(f"[SERVER EXCEPTION] {e}", flush=True)
            time.sleep(1)
        except KeyboardInterrupt:
            print("\nサーバーを停止しました。", flush=True)
            break


if __name__ == "__main__":
    main()
