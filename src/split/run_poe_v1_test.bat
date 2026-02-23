@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [3단계] POE 학교별 분리 (테스트 - OUTPUT 폴더, POE_V1)
python split_school_poe_v1.py --test
pause
