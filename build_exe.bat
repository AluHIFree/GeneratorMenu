@echo off
setlocal EnableExtensions

rem Always run from this script's folder
cd /d "%~dp0"

rem --- Find Python (no hardcoded paths, no .venv) ---
set "PY_EXE="
set "USE_PY=0"

where python >nul 2>&1
if not errorlevel 1 (
  for /f "delims=" %%I in ('where python 2^>nul') do (
    set "PY_EXE=%%I"
    goto :have_python
  )
)

where py >nul 2>&1
if not errorlevel 1 (
  py -3 -c "import sys" >nul 2>&1
  if not errorlevel 1 (
    set "USE_PY=1"
    goto :have_python
  )
)

where python >nul 2>&1
if not errorlevel 1 (
  set "PY_EXE=python"
  goto :have_python
)

echo.
echo Python not found.
echo Install Python 3 from https://www.python.org/downloads/
echo and enable "Add python.exe to PATH", then open a new terminal and run this script again.
echo.
pause
exit /b 1

:have_python
if "%USE_PY%"=="1" (
  echo Using Python: py -3
  py -3 --version
) else (
  echo Using Python: "%PY_EXE%"
  "%PY_EXE%" --version
)
if errorlevel 1 (
  echo.
  echo Python was detected but failed to start.
  echo Install Python 3 and add it to PATH, then try again.
  echo.
  pause
  exit /b 1
)

if not exist "generator.py" (
  echo.
  echo generator.py not found in:
  echo   %CD%
  echo Put build_exe.bat next to generator.py and run it again.
  echo.
  pause
  exit /b 1
)

rem Ensure PyInstaller is available
if "%USE_PY%"=="1" (
  py -3 -c "import PyInstaller" >nul 2>&1
) else (
  "%PY_EXE%" -c "import PyInstaller" >nul 2>&1
)
if errorlevel 1 (
  echo PyInstaller not found. Installing via pip...
  if "%USE_PY%"=="1" (
    py -3 -m pip install -U pip pyinstaller
  ) else (
    "%PY_EXE%" -m pip install -U pip pyinstaller
  )
  if errorlevel 1 (
    echo.
    echo Failed to install PyInstaller. Check your internet connection and pip.
    echo.
    pause
    exit /b 1
  )
)

echo.
echo Building one-file executable from generator.py ...
if "%USE_PY%"=="1" (
  py -3 -m PyInstaller --noconfirm --clean --onefile ^
    --name "GeneratorEzhednevnyhMenyu" ^
    --add-data "templates;templates" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell._writer ^
    generator.py
) else (
  "%PY_EXE%" -m PyInstaller --noconfirm --clean --onefile ^
    --name "GeneratorEzhednevnyhMenyu" ^
    --add-data "templates;templates" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell._writer ^
    generator.py
)

if exist "dist\GeneratorEzhednevnyhMenyu.exe" (
  copy /Y "dist\GeneratorEzhednevnyhMenyu.exe" "GeneratorEzhednevnyhMenyu.exe" >nul
  echo.
  echo Done:
  echo   dist\GeneratorEzhednevnyhMenyu.exe
  echo   GeneratorEzhednevnyhMenyu.exe
) else (
  echo.
  echo Build failed: dist\GeneratorEzhednevnyhMenyu.exe was not created.
  pause
  exit /b 1
)

pause
endlocal
