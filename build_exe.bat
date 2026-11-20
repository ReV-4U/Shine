@echo off
rem ==========================================================================
rem  Shine - construit dist\Shine.exe
rem  Le script verifie d'abord ce qui est DEJA installe et ne telecharge que
rem  ce qui manque :
rem    1) Python 3.10+      (sinon : telechargement depuis python.org, installation
rem                          pour l'utilisateur courant, sans droits administrateur)
rem    2) XlsxWriter, PyYAML, openpyxl, PyInstaller (sinon : pip install)
rem    3) Shine.exe         (deja a jour : rien a refaire ; /force pour reconstruire)
rem  Si tout est deja present, aucune connexion internet n'est necessaire.
rem ==========================================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "PYVER=3.12.8"
set "PYDIR=Python312"
set "ARCH=amd64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "ARCH=arm64"
set "PY="

echo.
echo [1/3] Verification de Python 3.10+ ...
py -3 -c "import sys; sys.exit(0 if sys.version_info>=(3,10) else 1)" >nul 2>&1 && set "PY=py -3"
if defined PY goto :have_python
python -c "import sys; sys.exit(0 if sys.version_info>=(3,10) else 1)" >nul 2>&1 && set "PY=python"
if defined PY goto :have_python
if exist "%LOCALAPPDATA%\Programs\Python\%PYDIR%\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\%PYDIR%\python.exe"
if defined PY goto :have_python

rem --- Python absent : telechargement + installation
echo      Python 3.10+ introuvable : telechargement de Python %PYVER% depuis python.org ...
set "INST=%TEMP%\python-%PYVER%-%ARCH%.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/%PYVER%/python-%PYVER%-%ARCH%.exe' -OutFile '%INST%'"
if errorlevel 1 goto :err_download
if not exist "%INST%" goto :err_download

echo      Verification de la signature de l'installateur ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=Get-AuthenticodeSignature '%INST%'; if($s.Status -eq 'Valid' -and $s.SignerCertificate.Subject -like '*Python Software Foundation*'){exit 0}else{exit 1}"
if errorlevel 1 goto :err_signature

echo      Installation de Python (quelques minutes, sans droits administrateur) ...
"%INST%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_test=0 Include_doc=0 Include_pip=1
if errorlevel 1 goto :err_install
del "%INST%" >nul 2>&1
set PY="%LOCALAPPDATA%\Programs\Python\%PYDIR%\python.exe"
if not exist %PY% goto :err_install

:have_python
%PY% --version
echo      OK : Python est present.

echo.
echo [2/3] Verification des composants (XlsxWriter, PyYAML, openpyxl, PyInstaller) ...
set "NEED_RT="
set "NEED_PI="
%PY% -c "import sys, xlsxwriter; sys.exit(0 if tuple(int(x) for x in xlsxwriter.__version__.split('.')[:2])>=(3,2) else 1)" >nul 2>&1 || set "NEED_RT=1"
%PY% -c "import sys, yaml; sys.exit(0 if int(yaml.__version__.split('.')[0])>=6 else 1)" >nul 2>&1 || set "NEED_RT=1"
%PY% -c "import sys, openpyxl; sys.exit(0 if tuple(int(x) for x in openpyxl.__version__.split('.')[:2])>=(3,1) else 1)" >nul 2>&1 || set "NEED_RT=1"
%PY% -c "import PyInstaller" >nul 2>&1 || set "NEED_PI=1"
if not defined NEED_RT if not defined NEED_PI goto :deps_ok

%PY% -m pip --version >nul 2>&1 || %PY% -m ensurepip --upgrade >nul 2>&1
if not defined NEED_RT goto :skip_rt
echo      Installation de XlsxWriter / PyYAML / openpyxl ...
%PY% -m pip install --disable-pip-version-check -r requirements.txt || goto :err_pip
:skip_rt
if not defined NEED_PI goto :deps_done
echo      Installation de PyInstaller ...
%PY% -m pip install --disable-pip-version-check pyinstaller || goto :err_pip
goto :deps_done
:deps_ok
echo      OK : tous les composants sont deja installes.
:deps_done

echo.
echo [3/3] Construction de Shine.exe ...
if /i "%~1"=="/force" goto :build
if not exist "dist\Shine.exe" goto :build
powershell -NoProfile -ExecutionPolicy Bypass -Command "$exe=(Get-Item 'dist\Shine.exe').LastWriteTime; $src=Get-ChildItem -Path . -Include *.py,config.yaml,config.spec -Recurse -File | Where-Object { $_.FullName -notmatch '\\(build|dist|tests)\\' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if($src -and $src.LastWriteTime -gt $exe){exit 1}else{exit 0}"
if errorlevel 1 goto :build
echo      OK : dist\Shine.exe est deja a jour (relancer avec /force pour le reconstruire).
goto :done

:build
%PY% -m PyInstaller --noconfirm --clean config.spec || goto :err
if not exist "dist\Shine.exe" goto :err

:done
echo.
echo ==========================================================================
echo  Termine : %~dp0dist\Shine.exe
echo  Copiez Shine.exe ou vous voulez. Un config.yaml place a cote est charge
echo  automatiquement (contact, organisation...) : pas besoin de reconstruire.
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

:err_pip
echo.
echo ECHEC : installation des composants impossible (reseau ou proxy bloquant pypi.org ?).
goto :err_end

:err
echo.
echo ECHEC de la construction. Verifiez le message d'erreur ci-dessus.

:err_end
pause
exit /b 1
