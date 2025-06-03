@echo off
cd /d "%~dp0"
echo Starting AI Excel Tool...
python main.py
if errorlevel 1 (
    echo.
    echo Failed to start. Possible Python environment issue.
    echo Please ensure Python is installed and properly configured.
    pause
) 