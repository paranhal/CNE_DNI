@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo [학교별 측정 값 현황] 유선망 전처리 - 대전
python wired_preprocess_v1.py -r DNI
pause
