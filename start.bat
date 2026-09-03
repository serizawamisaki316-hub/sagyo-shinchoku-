@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
title 倉庫作業進捗サイネージ 起動

echo ========================================================
echo   倉庫作業進捗サイネージシステムを起動しています...
echo ========================================================

:: 1. スタンドアロン実行ファイル（start_signage.exe）がある場合は最優先で起動
if exist "%~dp0start_signage.exe" (
    echo [OK] スタンドアロン実行ファイル（start_signage.exe）を使用します。
    start "SignageServer" "%~dp0start_signage.exe"
    goto wait_and_open
)

:: 2. 同梱のpython_runtimeがあるか確認
if exist "%~dp0python_runtime\python.exe" (
    set "PYTHONHOME=%~dp0python_runtime"
    set "PYTHONPATH=%~dp0python_runtime\Lib;%~dp0python_runtime\Lib\site-packages"
    echo [OK] 同梱のPythonランタイムを使用します。
    start "SignageServer" cmd /k "set "PYTHONHOME=%~dp0python_runtime" && set "PYTHONPATH=%~dp0python_runtime\Lib;%~dp0python_runtime\Lib\site-packages" && "%~dp0python_runtime\python.exe" server.py"
    goto wait_and_open
)

:: 3. システムのPythonがあるか確認
where python >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] システムのPythonを使用します。
    start "SignageServer" cmd /k "python server.py"
    goto wait_and_open
)

:: 実行環境が見つからない場合
echo.
echo ========================================================
echo 【エラー】起動に必要なプログラムが見つかりません。
echo ========================================================
echo start_signage.exe または python_runtime がフォルダ内に存在しません。
echo.
pause
exit /b 1

:wait_and_open
:: 4. サーバーの起動待機（最大10秒）
echo [待機中] サーバーの応答を待っています...
set /a count=0
:wait_loop
set /a count+=1
ping 127.0.0.1 -n 2 >nul
powershell -NoProfile -Command "(New-Object System.Net.Sockets.TcpClient).Connect('127.0.0.1', 8080)" >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] サーバーが正常に起動しました！
    goto launch_browser
)
if %count% geq 5 (
    goto launch_browser
)
goto wait_loop

:launch_browser
echo [開く] ブラウザでサイネージ画面を表示します...
start http://localhost:8080
exit
