@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [테스트] OUTPUT 폴더에 AP 파일 생성
python split_school_ap_v1.py --DNI --test
pause
