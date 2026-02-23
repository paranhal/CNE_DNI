@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [2단계] POE 원본 분석 - poe_structure_report_v1.txt 생성
python analyze_poe_structure_v1.py
echo.
echo === 분석 결과 ===
type poe_structure_report_v1.txt
pause
