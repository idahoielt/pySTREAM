@echo off
rem Installs pySTREAM on Windows: creates a Python environment in the .venv folder
rem and installs the packages pySTREAM needs. Double-click this file to run it.
cd /d "%~dp0"
echo Installing pySTREAM. This can take a few minutes...
echo.

rem Find Python: the 'py' launcher comes with the installer from python.org
set PYTHON=
py -3 --version >nul 2>nul && set PYTHON=py -3
if not defined PYTHON (
    python --version >nul 2>nul && set PYTHON=python
)
if not defined PYTHON (
    echo Python was not found.
    echo Install Python 3.12 or newer from https://www.python.org/downloads/
    echo and then double-click this file again.
    pause
    exit /b 1
)

%PYTHON% -c "import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)"
if errorlevel 1 (
    echo pySTREAM needs Python 3.12 or newer, but this is:
    %PYTHON% --version
    echo Install a newer Python from https://www.python.org/downloads/
    echo and then double-click this file again.
    pause
    exit /b 1
)

%PYTHON% -c "import tkinter" >nul 2>nul
if errorlevel 1 (
    echo Python was installed without tkinter, which pySTREAM uses for its window.
    echo Run the Python installer again, choose "Modify" and tick "tcl/tk and IDLE".
    pause
    exit /b 1
)

%PYTHON% -m venv .venv
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed

echo.
echo pySTREAM is installed. Start it by double-clicking start_windows.bat
pause
exit /b 0

:failed
echo.
echo The installation failed; see the messages above.
echo Check your internet connection and try again.
pause
exit /b 1
