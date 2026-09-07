@echo off
chcp 936 >nul
echo ============================================================
echo   Launch Microsoft Edge in Debug Mode
echo   (Keep this window open!)
echo ============================================================
echo.

set DEBUG_PORT=9222
set USER_DATA=%TEMP%\edge_sp_automation

REM Detect Edge install path
set EDGE_PATH=
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_PATH=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
)
if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_PATH=C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)
if exist "%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_PATH=%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"
)
if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_PATH=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
)

if "%EDGE_PATH%"=="" (
    echo [ERROR] Microsoft Edge not found!
    pause
    exit /b
)

REM Check if port is already in use
netstat -ano | findstr ":%DEBUG_PORT% " >nul 2>&1
if not errorlevel 1 (
    echo [WARN] Port %DEBUG_PORT% is already in use.
    echo        Edge may already be running in debug mode.
    echo        If tool cannot connect, close all Edge windows and retry.
    echo.
    pause
    exit /b
)

echo   Edge Path : %EDGE_PATH%
echo   Debug Port: %DEBUG_PORT%
echo   User Data : %USER_DATA%
echo.
echo   Please login to SharePoint in the Edge window.
echo   Then run SharePoint_Excel_Tool.exe
echo.
echo -----------------------------------------------------------
echo.

"%EDGE_PATH%" --remote-debugging-port=%DEBUG_PORT% --remote-allow-origins=* --user-data-dir="%USER_DATA%" --no-first-run --no-default-browser-check

echo.
echo   Edge closed.
pause
