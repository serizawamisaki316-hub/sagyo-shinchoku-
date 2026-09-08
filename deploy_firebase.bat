@echo off
chcp 65001 > nul
cd /d "%~dp0"
title Firebase Hosting 自動デプロイ

set "PYTHON_EXE=C:\Users\00137184\AppData\Local\Programs\Python\Python314\python.exe"
if not exist "%PYTHON_EXE%" (
    set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" "%~dp0deploy_firebase.py"

pause
