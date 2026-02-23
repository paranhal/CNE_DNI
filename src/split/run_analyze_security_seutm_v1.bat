@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [2단계] 보안장비 원본 분석 - security_structure_report_seutm_v1.txt 생성
python analyze_security_structure_seutm_v1.py
echo.
echo === 분석 결과 ===
type security_structure_report_seutm_v1.txt
pause
