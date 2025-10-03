@echo off
REM Build eksel executable for Windows
echo Building eksel executable...

REM Check if virtual environment exists
if not exist "..\\.venv\\Scripts\\activate.bat" (
    echo Error: Virtual environment not found at ..\.venv
    echo Please create a virtual environment first: python -m venv .venv
    pause
    exit /b 1
)

REM Activate virtual environment
echo Activating virtual environment...
call "..\\.venv\\Scripts\\activate.bat"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found in virtual environment
    pause
    exit /b 1
)

REM Run the build script
python build_exe.py

REM Deactivate virtual environment
deactivate

echo.
echo Press any key to exit...
pause >nul
