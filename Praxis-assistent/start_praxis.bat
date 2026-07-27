@echo off
setlocal

set "BASIS=%~dp0"
set "PY=%BASIS%.venv\Scripts\python.exe"
set "PYW=%BASIS%.venv\Scripts\pythonw.exe"

if not exist "%PY%" set "PY=python"
if not exist "%PYW%" set "PYW=pythonw"

"%PYW%" "%BASIS%main.py"
if errorlevel 1 "%PY%" "%BASIS%main.py"
exit /b 0
