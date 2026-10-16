@echo off
rem Ручной разовый запуск сервера «Журнал приказов» на Windows.
rem Порт можно изменить: start_server.bat 8090  (по умолчанию 8080)
cd /d "%~dp0..\.."
set PORT=%1
if "%PORT%"=="" set PORT=8080
python app.py --port %PORT%
pause
