@echo off
echo === Goodwood Screening System - Installer ===
python -m venv venv
call venv\Scripts\activate.bat
pip install -r requirements.txt
echo.
echo Instalasi selesai. Jalankan start_windows.bat untuk memulai.
pause
