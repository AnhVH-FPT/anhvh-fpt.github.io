@echo off
chcp 65001 >nul
title Database manager - http://localhost:8765/
cd /d "%~dp0"

rem 1. Python must be installed (python.org, tick "Add python.exe to PATH")
where python >nul 2>nul
if errorlevel 1 (
  echo Python chua duoc cai. Tai tai https://www.python.org/downloads/ va tick "Add python.exe to PATH" khi cai.
  pause
  exit /b 1
)

rem 2. Install openpyxl the first time on a new computer
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Dang cai thu vien openpyxl lan dau...
  python -m pip install --user -r "%~dp0requirements.txt"
)

echo Database manager: http://localhost:8765/
python "%~dp0server.py" --port 8765 %*
if errorlevel 1 pause
