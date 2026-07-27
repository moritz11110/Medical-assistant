@echo off
setlocal

set "BASIS=%~dp0"
set "PY=%BASIS%.venv\Scripts\python.exe"

if not exist "%PY%" set "PY=python"

echo [DEBUG] Basis: %BASIS%
echo [DEBUG] Python: %PY%
echo [DEBUG] Starte updater.py ...
"%PY%" "%BASIS%updater.py"
echo [DEBUG] updater.py beendet mit ExitCode=%ERRORLEVEL%

echo [DEBUG] Starte main.py sichtbar ...
"%PY%" "%BASIS%main.py"
echo [DEBUG] main.py beendet mit ExitCode=%ERRORLEVEL%

echo.
echo [DEBUG] Zum Schliessen Taste druecken ...
pause >nul
exit /b 0
