@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo 파이썬이 설치되어 있지 않습니다. https://www.python.org/downloads 에서 설치하세요.
  echo 설치 시 "Add Python to PATH" 체크 필수!
  pause
  exit /b
)
python -c "import openpyxl" 2>nul || pip install openpyxl
python dart_gui.py
if errorlevel 1 pause
