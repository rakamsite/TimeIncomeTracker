@echo off
setlocal
where python >nul 2>nul
if %errorlevel% neq 0 (
  echo Python not found. Please install Python 3.12+ for building.
  exit /b 1
)
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --name TimeIncomeTracker main.py
if %errorlevel% neq 0 (
  echo Build failed.
  exit /b 1
)
echo Build successful: dist\TimeIncomeTracker.exe
endlocal
