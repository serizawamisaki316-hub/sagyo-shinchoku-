# Firebase Hosting 一括自動デプロイスクリプト
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Firebase Hosting デプロイ (warehouse-work-progress)"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  倉庫作業進捗サイネージ - Firebase Hosting デプロイ" -ForegroundColor Cyan
Write-Host "  対象プロジェクト: warehouse-work-progress" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

$workDir = $PSScriptRoot
Set-Location $workDir

# 1. node_portable パス設定
$nodeDir = Join-Path $workDir "node_portable\node-v20.18.0-win-x64"
if (Test-Path "$nodeDir\node.exe") {
    $env:PATH = "$nodeDir;$env:PATH"
    Write-Host "[OK] ポータブルNode.js環境を検出しました。" -ForegroundColor Green
} else {
    Write-Host "[INFO] システムのNode.js環境を使用します。" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "[1/2] Firebase ログイン認証の確認..." -ForegroundColor Yellow
Write-Host "※ブラウザが起動した場合は" -ForegroundColor White
Write-Host "  kanagawa.toyota.parts@gmail.com で「許可」をクリックしてください。" -ForegroundColor White
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

& npx.cmd -y firebase-tools login

Write-Host ""
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "[2/2] Firebase Hosting へファイルをアップロード（デプロイ）中..." -ForegroundColor Yellow
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

& npx.cmd -y firebase-tools deploy --only hosting --project warehouse-work-progress

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "  [成功] デプロイが完了しました！" -ForegroundColor Green
    Write-Host "  サイネージ公開URL:" -ForegroundColor White
    Write-Host "  https://warehouse-work-progress.web.app" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Green
    
    Start-Process "https://warehouse-work-progress.web.app"
} else {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "  [エラー] デプロイに失敗しました。画面のエラー内容をご確認ください。" -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
}

Write-Host ""
Write-Host "何かキーを押すとこの画面を閉じます..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
