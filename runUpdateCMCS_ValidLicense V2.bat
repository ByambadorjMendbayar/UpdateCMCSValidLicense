@echo off
REM CMCS Valid License Updater - Batch Launcher
REM This batch file runs the Python script and keeps the window open

cd /d "%~dp0"

echo Starting CMCS Valid License Updater...
echo.

python "UpdateCMCS_ValidLicense v2.py"

REM Check if Python executed successfully
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Python script failed to run
    echo.
    echo Possible causes:
    echo - Python is not installed
    echo - Python is not in your system PATH
    echo.
    echo Please install Python 3.8 or higher from https://www.python.org
    echo.
    pause
)