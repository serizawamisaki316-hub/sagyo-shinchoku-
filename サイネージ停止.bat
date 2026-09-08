@echo off
setlocal
cd /d "%~dp0"
title 倉庫作業進捗サイネージ 停止

echo ========================================================
echo   倉庫作業進捗サイネージシステムを停止しています...
echo ========================================================

:: 1. ポート8080を使用しているプロセスを探して終了
powershell -NoProfile -Command "$ports = Get-NetTCPConnection -LocalPort 8080 -ErrorAction SilentlyContinue; if ($ports) { $ports | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue } }"

:: 2. server.py を実行している python / pythonw プロセスがあれば終了
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*server.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"

:: 3. start_signage.exe が動いていれば終了
taskkill /f /im start_signage.exe >nul 2>&1

echo.
echo [完了] サイネージサーバーを停止しました。
echo.
timeout /t 2 >nul
exit
