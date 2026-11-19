@echo off
rem Construit Shine.exe (à lancer une fois, sur un PC Windows avec Python 3.10+ et un accès internet).
cd /d "%~dp0"
py -m pip install -r requirements.txt pyinstaller || goto :err
py -m PyInstaller --noconfirm --clean --onefile --windowed --name Shine --add-data "config.yaml;." shine.py || goto :err
echo.
echo Termine : %~dp0dist\Shine.exe
echo Copiez Shine.exe ou vous voulez (un shine.yaml place a cote est charge automatiquement).
pause
exit /b 0
:err
echo.
echo Echec de la construction.
pause
exit /b 1
