@echo off
chcp 65001 >nul
echo ============================================================
echo [학교별 측정 리포트] 생성
echo 1. TOTAL_MEASURE_LIST 통합 (없으면 build_total_measure_list.py 실행)
echo 2. 템플릿 기반 학교별 파일 생성
echo ============================================================
cd /d "%~dp0"
if not exist "CNE\TOTAL_MEASURE_LIST_V1.XLSX" (
    echo [1단계] TOTAL_MEASURE_LIST 생성...
    python build_total_measure_list.py
    if errorlevel 1 (
        echo [오류] 통합 실패. TOTAL_MEASURE_LIST_V1.XLSX를 수동으로 준비하세요.
        pause
        exit /b 1
    )
)
echo [2단계] 학교별 리포트 생성...
python school_report_generator.py
echo [3단계] 학교별 자료 유무 확인...
python school_data_check.py
echo.
pause
