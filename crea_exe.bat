@echo off
rem Crea eFootballMonitor.exe (serve Python installato SOLO su questo PC)
cd /d "%~dp0"
echo Installo gli strumenti necessari...
python -m pip install --upgrade pyinstaller openpyxl scapy || goto errore
echo Creo il programma...
python -m PyInstaller --onefile --uac-admin --name eFootballMonitor --collect-submodules scapy --clean efootball_monitor.py || goto errore
echo.
echo Fatto! Il programma e' in: %~dp0dist\eFootballMonitor.exe
pause
exit /b 0
:errore
echo.
echo Qualcosa e' andato storto: copia il messaggio qui sopra e mandalo.
pause
exit /b 1
