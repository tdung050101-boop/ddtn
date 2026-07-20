@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Khong tim thay Python. Hay cai Python 3.11 hoac moi hon va tick Add Python to PATH.
  pause
  exit /b 1
)

py -3 -m venv .venv
if errorlevel 1 goto :error

.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto :error

if not exist .env copy .env.example .env >nul
notepad .env

echo.
echo Cai dat xong. Dan Bot Token vao file .env, luu lai, roi chay 2_KIEM_TRA.bat.
pause
exit /b 0

:error
echo.
echo Cai dat that bai. Hay chup lai man hinh loi.
pause
exit /b 1
