@echo off
chcp 936 >nul
echo ============================================================
echo   SharePoint Excel Tool - Build EXE
echo ============================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.8+
    pause
    exit /b
)

REM Check PyInstaller
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo [START] Building EXE...
echo.

pyinstaller --onefile --name "SharePoint_Excel_Tool" --console --collect-all playwright --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywinauto sharepoint_excel_automation.py

echo.
if exist "dist\SharePoint_Excel_Tool.exe" (
    echo ============================================================
    echo   SUCCESS! Output: dist\SharePoint_Excel_Tool.exe
    echo ============================================================
    echo.
    echo   Files to distribute:
    echo     1. dist\SharePoint_Excel_Tool.exe
    echo     2. launch_edge_debug.bat
    echo.
) else (
    echo [ERROR] Build failed. Check error messages above.
)

pause
