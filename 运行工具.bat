@echo off
cd /d "%~dp0"
python src\main.py
if %errorlevel% neq 0 pause
