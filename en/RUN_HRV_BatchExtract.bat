@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY if exist "%USERPROFILE%\anaconda3\python.exe" set PY="%USERPROFILE%\anaconda3\python.exe"
if not defined PY if exist "%USERPROFILE%\miniconda3\python.exe" set PY="%USERPROFILE%\miniconda3\python.exe"
if not defined PY if exist "C:\ProgramData\anaconda3\python.exe" set PY="C:\ProgramData\anaconda3\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\anaconda3\python.exe" set PY="%LOCALAPPDATA%\anaconda3\python.exe"
if not defined PY (
  echo Python not found. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
  pause
  exit /b 1
)
if not exist ".deps_ok" (
  echo Installing libraries on the first run...
  %PY% -m pip install --user pyautogui pyperclip pygetwindow pandas numpy openpyxl pillow
  if errorlevel 1 (
    echo Failed to install libraries. Check the internet connection and try again.
    pause
    exit /b 1
  )
  %PY% -m pip install --user winsdk
  echo ok> ".deps_ok"
)
%PY% HRV_BatchExtract.py
if errorlevel 1 pause
