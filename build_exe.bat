@echo off
rem ==========================================================================
rem  Shine - construit dist\Shine.exe
rem  Si Python 3.10+ est absent, il est telecharge depuis python.org puis
rem  installe pour l'utilisateur courant (sans droits administrateur).
rem  Necessite un acces internet. A lancer une seule fois.
rem ==========================================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "PYVER=3.12.8"
set "PYDIR=Python312"
set "ARCH=amd64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "ARCH=arm64"
set "PY="

rem --- 1) Python 3.10+ deja present ?
py -3 -c "import sys; sys.exit(0 if sys.version_info>=(3,10) else 1)" >nul 2>&1 && set "PY=py -3"
if defined PY goto :have_python
python -c "import sys; sys.exit(0 if sys.version_info>=(3,10) else 1)" >nul 2>&1 && set "PY=python"
if defined PY goto :have_python

rem --- 2) installation utilisateur precedente ?
if exist "%LOCALAPPDATA%\Programs\Python\%PYDIR%\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\%PYDIR%\python.exe"
if defined PY goto :have_python

rem --- 3) telechargement + installation de Python
echo.
echo Python 3.10+ introuvable : telechargement de Python %PYVER% depuis python.org ...
set "INST=%TEMP%\python-%PYVER%-%ARCH%.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/%PYVER%/python-%PYVER%-%ARCH%.exe' -OutFile '%INST%'"
if errorlevel 1 goto :err_download
if not exist "%INST%" goto :err_download

echo Verification de la signature de l'installateur ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=Get-AuthenticodeSignature '%INST%'; if($s.Status -eq 'Valid' -and $s.SignerCertificate.Subject -like '*Python Software Foundation*'){exit 0}else{exit 1}"
if errorlevel 1 goto :err_signature

echo Installation de Python (quelques minutes, sans droits administrateur) ...
"%INST%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_test=0 Include_doc=0 Include_pip=1
if errorlevel 1 goto :err_install
del "%INST%" >nul 2>&1
set PY="%LOCALAPPDATA%\Programs\Python\%PYDIR%\python.exe"
if not exist %PY% goto :err_install

:have_python
echo.
echo Python utilise :
%PY% --version || goto :err
echo.
echo Installation des dependances (XlsxWriter, PyYAML, PyInstaller) ...
%PY% -m pip install --disable-pip-version-check -r requirements.txt pyinstaller || goto :err
echo.
echo Construction de Shine.exe ...
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name Shine --add-data "config.yaml;." shine.py || goto :err
if not exist "dist\Shine.exe" goto :err

echo.
echo ==========================================================================
echo  Termine : %~dp0dist\Shine.exe
echo  Copiez Shine.exe ou vous voulez. Un shine.yaml place a cote est charge
echo  automatiquement (date limite, contact...).
echo ==========================================================================
pause
exit /b 0

:err_download
echo.
echo ECHEC : telechargement de Python impossible (reseau, proxy ou pare-feu ?).
echo Installez Python 3.10+ a la main depuis https://www.python.org/downloads/
echo (cochez "Add python.exe to PATH") puis relancez ce fichier.
goto :err_end

:err_signature
echo.
echo ECHEC : la signature de l'installateur Python n'est pas valide. Installation annulee.
del "%INST%" >nul 2>&1
goto :err_end

:err_install
echo.
echo ECHEC : l'installation de Python a echoue. Installez-le a la main depuis
echo https://www.python.org/downloads/ puis relancez ce fichier.
goto :err_end

:err
echo.
echo ECHEC de la construction. Verifiez le message d'erreur ci-dessus.

:err_end
pause
exit /b 1
