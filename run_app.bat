@echo off
cd /d "%~dp0"
echo Installing / updating dependencies...
pip install -r requirements.txt
echo.
echo Starting Brokerage Filler app...
echo The browser will open automatically. Keep this window open while using the app.
echo Close this window to stop the app.
echo.
streamlit run app.py
pause
