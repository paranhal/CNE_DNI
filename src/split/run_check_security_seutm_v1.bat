@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1단계] 보안장비 원본 구조 점검
python check_structure_security_seutm_v1.py
pause
