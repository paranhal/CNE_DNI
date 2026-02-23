@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [3단계] 보안장비 학교별 분리 (테스트 - OUTPUT 폴더)
python split_school_security_seutm_v1.py --test
pause
