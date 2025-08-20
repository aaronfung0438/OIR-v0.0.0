@echo off
echo Starting OIR Report System v6...
echo.
echo Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH!
    pause
    exit /b 1
)

echo.
echo Installing/Updating dependencies...
pip install -r requirements.txt

echo.
echo Starting Flask server...
echo Open your browser and go to: http://127.0.0.1:5000
echo Press Ctrl+C to stop the server
echo.

python app.py

pause
