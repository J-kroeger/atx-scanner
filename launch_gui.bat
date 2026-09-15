@echo off
echo Starting ATX-Scanner...

REM Check if running from dist (exe mode) or source (python mode)
if exist "%~dp0ATX-Scanner.exe" (
    start "" "%~dp0ATX-Scanner.exe"
    exit /b
)

REM Fallback: run from Python source
if exist "%~dp0atx_scanner.py" (
    python "%~dp0atx_scanner.py"
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to start. Make sure Python 3.8+ is installed.
        echo Install dependencies: pip install -r requirements.txt
        pause
    )
) else (
    echo ERROR: Neither ATX-Scanner.exe nor atx_scanner.py found.
    echo Please run this from the ATX-Scanner directory.
    pause
)
