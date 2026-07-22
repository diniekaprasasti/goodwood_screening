@echo off
call venv\Scripts\activate.bat
start http://localhost:5001
python app.py
