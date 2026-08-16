@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python environment...
    py -m venv .venv

    if errorlevel 1 (
        echo.
        echo Could not create the Python environment.
        echo Make sure Python is installed.
        pause
        exit /b 1
    )
)

if not exist ".venv\Lib\site-packages\openpyxl" (
    echo Installing openpyxl...
    .venv\Scripts\python.exe -m pip --isolated install --index-url https://pypi.org/simple openpyxl

    if errorlevel 1 (
        echo.
        echo Could not install openpyxl.
        echo.
        echo If your internet connection is unavailable,
        echo install openpyxl manually and run ranker.py.
        pause
        exit /b 1
    )
)

echo.
echo Starting Goodreads To-Read Ranker...
echo.

.venv\Scripts\python.exe ranker.py

if errorlevel 1 (
    echo.
    echo The application exited with an error.
    pause
)