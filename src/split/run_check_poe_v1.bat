@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1단계] POE 원본 구조 점검 (POE_V1)
python check_structure_poe_v1.py
pause
