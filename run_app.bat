@echo off
cd /d "%~dp0"
echo Installing / updating dependencies...
pip install -r requirements.txt -q
echo.
echo Starting BrokerageAI...
echo.
echo  Open your browser at:  http://localhost:8502
echo.
echo  Close this window to stop the app.
echo.
start "" "http://localhost:8502"
python server.py
pause
