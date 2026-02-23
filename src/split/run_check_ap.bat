@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1단계] AP 원본 구조 점검 (DNI/DNI_AP_LIST.XLSX)
python check_structure_ap.py
pause
