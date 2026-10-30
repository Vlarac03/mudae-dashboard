@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
rem First time: no settings yet
if not exist "dades\config.json" (
  call "%~dp0Configura.bat"
  exit /b
)
rem Find Python 3.12+: python on PATH, else the py launcher (installed even without "Add python.exe to PATH"), else the usual folder
set "PYEXE="
for /f "delims=" %%p in ('python -c "import sys; sys.version_info >= (3, 12) and print(sys.executable)" 2^>nul') do set "PYEXE=%%p"
if not defined PYEXE for /f "delims=" %%p in ('py -3 -c "import sys; sys.version_info >= (3, 12) and print(sys.executable)" 2^>nul') do set "PYEXE=%%p"
if not defined PYEXE for /d %%d in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%d\python.exe" set "PYEXE=%%d\python.exe"
if not defined PYEXE goto setup
"%PYEXE%" -c "import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)" >nul 2>nul
if errorlevel 1 goto setup
set "PYWEXE=%PYEXE:python.exe=pythonw.exe%"
if not exist "%PYWEXE%" set "PYWEXE=%PYEXE%"
rem 1) Update the Excel file and the dashboard from the dades folder
"%PYEXE%" mudae.py
if errorlevel 1 goto failed
goto open
:setup
call "%~dp0Configura.bat"
exit /b
:failed
echo.
echo Could not update - is the Excel file open? The dashboard will open with the previous data.
pause
:open
rem 2) Open the dashboard. The button server keeps running in the background, with no window.
start "" "%PYWEXE%" mudae.py tauler
timeout /t 3 >nul
