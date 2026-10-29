@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set "PYDIR=%USERPROFILE%\AppData\Local\Programs\Python\Python313"
set "PYEXE=python"
set "PYWEXE=pythonw"
if exist "%PYDIR%\python.exe" set "PYEXE=%PYDIR%\python.exe"
if exist "%PYDIR%\pythonw.exe" set "PYWEXE=%PYDIR%\pythonw.exe"
"%PYEXE%" -c "import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)" >nul 2>nul
if errorlevel 1 goto nopython
echo.
echo === Mudae dashboard setup ===
echo.
echo Language / Idioma:
echo   1 = English - default
echo   2 = Espanol
echo   3 = Catala
set "IDIOMA="
set /p "IDIOMA=[1/2/3]: "
if "%IDIOMA%"=="2" goto es
if "%IDIOMA%"=="3" goto ca
set "IDIOMA=1"
set "Q_USER=Your Mudae username - the one shown in 'Belongs to ...': "
set "Q_SERVER=Your Discord server name - optional, press Enter to skip: "
set "M_DONE=Done! The dashboard will open now. From now on, open it with the Mudae shortcut on your desktop."
goto ask
:es
set "Q_USER=Tu nombre en Mudae - el que sale en 'Pertenece a ...': "
set "Q_SERVER=Nombre de vuestro servidor de Discord - opcional, Enter para saltar: "
set "M_DONE=Hecho! Ahora se abrira el panel. A partir de ahora, abrelo con el acceso directo Mudae del escritorio."
goto ask
:ca
set "Q_USER=El teu nom a Mudae - el que surt a 'Pertenece a ...': "
set "Q_SERVER=Nom del vostre servidor de Discord - opcional, Enter per saltar: "
set "M_DONE=Fet! Ara s'obrira el tauler. A partir d'ara obre'l amb l'acces directe Mudae de l'escriptori."
:ask
echo.
set "USUARI="
set "SERVIDOR="
set /p "USUARI=%Q_USER%"
set /p "SERVIDOR=%Q_SERVER%"
"%PYEXE%" mudae.py configura "%USUARI%" "%SERVIDOR%" "%IDIOMA%"
if errorlevel 1 goto failed
echo.
echo %M_DONE%
pause
call "%~dp0Actualitza.bat"
exit /b
:failed
pause
exit /b 1
:nopython
echo Python 3.12 or newer is needed and was not found.
echo Download it from https://www.python.org/downloads/
echo When installing it, tick "Add python.exe to PATH". Then open this file again.
echo.
echo Cal Python 3.12 o mes nou / Hace falta Python 3.12 o mas nuevo: https://www.python.org/downloads/
pause
exit /b 1
