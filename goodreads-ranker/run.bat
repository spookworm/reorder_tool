@echo off
setlocal
cd /d "%~dp0"

echo.
echo ==========================================
echo   Goodreads To-Read Ranker
echo ==========================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python environment...
    py -m venv .venv

    if errorlevel 1 (
        echo.
        echo Could not create the Python environment.
        echo Make sure Python is installed and available through "py".
        echo.
        pause
        exit /b 1
    )
)

echo Checking required packages...
.venv\Scripts\python.exe -m pip --isolated show openpyxl >nul 2>&1

if errorlevel 1 (
    echo Installing openpyxl...
    .venv\Scripts\python.exe -m pip --isolated install --index-url https://pypi.org/simple openpyxl

    if errorlevel 1 (
        echo.
        echo Could not install openpyxl.
        echo.
        echo If your internet connection is unavailable,
        echo install openpyxl manually and run ranker.py.
        echo.
        pause
        exit /b 1
    )
)

echo.
echo Starting Goodreads To-Read Ranker...
echo.

.venv\Scripts\python.exe ranker.py

set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo ==========================================
    echo   The application exited with an error.
    echo   Exit code: %EXIT_CODE%
    echo ==========================================
    echo.
    pause
)

exit /b %EXIT_CODE%