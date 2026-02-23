@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo [학교별 측정 값 현황] 무선망 ISP 전처리
echo - CNE_ISP_MEASURE.XLSX 파일의 ISP측정 시트에서 학교별 평균 계산
python wireless_preprocess_v1.py -r CNE
pause
