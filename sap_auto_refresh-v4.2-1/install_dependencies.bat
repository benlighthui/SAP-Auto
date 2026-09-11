@echo off
echo ================================================
echo  Installing Python dependencies...
echo ================================================
echo.
pip install playwright pywin32 pywinauto
echo.
echo Installing Playwright browser driver...
python -m playwright install chromium
echo.
echo Done! All dependencies installed.
pause
