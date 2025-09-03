@echo off
REM --- Flask app run ---
echo Starting Flask server...
start cmd /k "python app.py"

REM --- Wait for 5 seconds to let Flask start ---
timeout /t 5

REM --- Ngrok run ---
echo Starting Ngrok tunnel...
start cmd /k "ngrok http 5000"

echo Done! Flask is running locally and public link is available via Ngrok.
pause
