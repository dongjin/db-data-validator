@echo off
setlocal

python -m pip install --upgrade pip
python -m pip install -r requirements-dev.txt
python -m pip install -e .

pyinstaller --noconfirm --clean --windowed --name db_data_validator --paths src run_app.py

echo.
echo Build complete. Find the application under dist\db_data_validator\
pause
