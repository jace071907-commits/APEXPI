@echo off
cd /d "%~dp0"
python build-dashboard.py
if errorlevel 1 pause
