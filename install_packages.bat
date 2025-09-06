@echo off
echo Installing BIMailer dependencies...
echo.

REM Check if we're in a virtual environment
if defined VIRTUAL_ENV (
    echo Using virtual environment: %VIRTUAL_ENV%
) else (
    echo No virtual environment detected. Consider using one for better isolation.
)

echo.
echo Installing required packages...
pip install Pillow>=9.0.0
pip install reportlab>=3.6.0
pip install pandas>=1.3.0

echo.
echo Installation complete!
echo.
echo To test the system, run:
echo   python test_basic_functionality.py
echo.
echo To run the full system, run:
echo   cd Scripts
echo   python main.py diagnostics
echo.
pause
