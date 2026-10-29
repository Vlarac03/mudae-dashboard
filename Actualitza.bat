@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
rem First time: no settings yet
if not exist "dades\config.json" (
  call "%~dp0Configura.bat"
  exit /b
)
set "PYDIR=%USERPROFILE%\AppData\Local\Programs\Python\Python313"
set "PYEXE=python"
set "PYWEXE=pythonw"
if exist "%PYDIR%\python.exe" set "PYEXE=%PYDIR%\python.exe"
if exist "%PYDIR%\pythonw.exe" set "PYWEXE=%PYDIR%\pythonw.exe"
rem 1) Update the Excel file and the dashboard from the dades folder
"%PYEXE%" mudae.py
if errorlevel 1 goto failed
goto open
:failed
echo.
echo Could not update - is the Excel file open? The dashboard will open with the previous data.
pause
:open
rem 2) Open the dashboard. The button server keeps running in the background, with no window.
start "" "%PYWEXE%" mudae.py tauler
timeout /t 3 >nul
