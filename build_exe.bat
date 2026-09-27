@echo off
rem Builds eFootballMonitor.exe (Python is needed ONLY on this PC)
cd /d "%~dp0"
echo Installing build tools...
python -m pip install --upgrade pyinstaller openpyxl scapy || goto error
echo Building the program...
python -m PyInstaller --onefile --uac-admin --name eFootballMonitor --collect-submodules scapy --clean efootball_monitor.py || goto error
echo.
echo Done! The program is in: %~dp0dist\eFootballMonitor.exe
pause
exit /b 0
:error
echo.
echo Something went wrong: copy the message above and send it.
pause
exit /b 1
