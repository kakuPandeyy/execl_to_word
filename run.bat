@echo off
cd /d "%~dp0"
python excel_to_word.py
if errorlevel 1 pause
