@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [2단계] AP 원본 분석 - ap_structure_report.txt (DNI/DNI_AP_LIST.XLSX)
python analyze_ap_structure.py
echo.
echo === 분석 결과 ===
type ap_structure_report.txt
pause
