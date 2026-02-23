@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM 통합 스크립트: 장비 + 지역 + 기타 옵션
REM 예: run_split_all.bat AP DNI --test
REM 예: run_split_all.bat switch CNE --missed-only
REM 예: run_split_all.bat security DNI
set DEVICE=%1
set REGION=%2
shift
shift
if "%DEVICE%"=="" set DEVICE=AP
if "%REGION%"=="" set REGION=DNI
echo [통합] 장비=%DEVICE% 지역=%REGION% %*
python split_school_all_v1.py --%DEVICE% --%REGION% %*
if errorlevel 1 pause
