@echo off
setlocal

REM Build Windows EXE for the app.
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --name topstories app.py

echo.
echo Build complete. EXE is in the dist folder.
pause
