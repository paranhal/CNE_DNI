@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo [학교별 측정 값 현황] 유선망 전처리
echo - 상단 RUN_REGION 값에 따라 충남/대전/둘 다 실행
python wired_preprocess_v1.py
pause
