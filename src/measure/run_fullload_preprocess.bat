@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo [학교별 측정 값 현황] 전부하 측정 전처리
echo - CNE_FULLLOAD_MEASURE.xlsx 파일의 전부하측정 시트에서 학교별 평균 계산
python fullload_preprocess_v1.py -r CNE
pause
