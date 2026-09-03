@echo off
chcp 65001 >nul
title サイネージ自動起動 登録

echo ========================================================
echo   Windows起動時にサイネージを自動起動するよう登録します
echo ========================================================

set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_SCRIPT=%TEMP%\create_shortcut.vbs"
set "TARGET_BAT=%~dp0start.bat"
set "SHORTCUT_PATH=%STARTUP_FOLDER%\倉庫作業進捗サイネージ.lnk"

echo Set oWS = WScript.CreateObject("WScript.Shell") > "%VBS_SCRIPT%"
echo sLinkFile = "%SHORTCUT_PATH%" >> "%VBS_SCRIPT%"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%VBS_SCRIPT%"
echo oLink.TargetPath = "%TARGET_BAT%" >> "%VBS_SCRIPT%"
echo oLink.WorkingDirectory = "%~dp0" >> "%VBS_SCRIPT%"
echo oLink.Save >> "%VBS_SCRIPT%"

cscript //nologo "%VBS_SCRIPT%"
del "%VBS_SCRIPT%"

echo.
echo 【完了】スタートアップフォルダにショートカットを登録しました！
echo 次回からPCを起動すると、自動的にサイネージ画面が立ち上がります。
echo.
pause
