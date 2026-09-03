@echo off
chcp 65001 >nul
title サイネージ自動起動 解除

echo ========================================================
echo   Windows起動時の自動起動を解除します
echo ========================================================

set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "SHORTCUT_PATH=%STARTUP_FOLDER%\倉庫作業進捗サイネージ.lnk"

if exist "%SHORTCUT_PATH%" (
    del "%SHORTCUT_PATH%"
    echo.
    echo 【完了】自動起動の登録を解除しました。
) else (
    echo.
    echo 自動起動は登録されていませんでした。
)

echo.
pause
