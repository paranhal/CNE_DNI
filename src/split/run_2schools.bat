@echo off
chcp 65001 >nul
cd /d "%~dp0"
python split_school_ap_v1.py --only-schools N108151062ES,N108151079ES
pause
