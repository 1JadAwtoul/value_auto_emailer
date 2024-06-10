@echo off
REM Navigate to the script directory
cd /d C:\Users\aawtoul\Desktop\auto_emailer

REM Activate the virtual environment
CALL C:\Users\aawtoul\Desktop\auto_emailer\venv\Scripts\activate.bat

REM Run the Python script
python main.py

REM Optional: Add a pause to see the output if running manually for testing
REM pause
