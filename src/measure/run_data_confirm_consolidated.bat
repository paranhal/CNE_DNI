@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [데이터확인 통합] 개별 리포트 -^> 통합 파일
python build_data_confirm_consolidated.py
pause
