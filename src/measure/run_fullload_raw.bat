@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo [전부하] FULLLOAD_RAWA_1 -^> 통계용 원본 -^> 학교별 평균
echo - 1차/2차 선택 로직, 학교코드 수정, 로그 저장
python fullload_raw_to_source.py
pause
