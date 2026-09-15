@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python -m pip install -q pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "GeneratorEzhednevnyhMenyu" ^
  --add-data "templates;templates" ^
  --collect-all customtkinter ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.cell._writer ^
  --hidden-import gui ^
  main.py
if exist "dist\GeneratorEzhednevnyhMenyu.exe" (
  copy /Y "dist\GeneratorEzhednevnyhMenyu.exe" "..\GeneratorEzhednevnyhMenyu.exe" >nul
  echo.
  echo Готово:
  echo   menu_generator\dist\GeneratorEzhednevnyhMenyu.exe
  echo   GeneratorEzhednevnyhMenyu.exe  ^(рядом с tm/kp^)
) else (
  echo Ошибка сборки
)
pause
