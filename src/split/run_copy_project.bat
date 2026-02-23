@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [복사] D:\CNE_DNI\src\split 로 프로젝트 파일 복사
python copy_project.py
echo.
pause
