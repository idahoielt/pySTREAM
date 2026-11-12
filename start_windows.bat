@echo off
rem Starts pySTREAM on Windows. Double-click this file to run it.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo pySTREAM is not installed yet. Double-click install_windows.bat first.
    pause
    exit /b 1
)
echo Starting pySTREAM. The window can take up to a minute to appear...
".venv\Scripts\python.exe" program\pySTREAM.py
rem Keep this window open if pySTREAM stopped with an error, so the message can be read
if errorlevel 1 pause
