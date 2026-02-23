@echo off
chcp 65001 >nul
cd /d "%~dp0"
python analyze_switch_structure.py
type structure_report.txt
pause
