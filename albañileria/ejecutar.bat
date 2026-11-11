@echo off
REM ALBANILERIA E.070 - doble clic para abrir el programa
cd /d "%~dp0"
python main.py
if errorlevel 1 pause
