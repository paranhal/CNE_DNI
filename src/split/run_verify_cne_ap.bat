@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [검증] 충남 AP - 빠진 학교
python verify_missing_by_code.py --CNE --AP
pause
