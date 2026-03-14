@echo off
chcp 65001 >nul
REM Windows용 ISP 원데이터 생성기 실행 파일 빌드
REM 사용: 프로젝트 루트(CNE_DNI)에서 scripts\build_isp_orig_generator_win.bat 실행
REM 필요: Python + pip install pyinstaller openpyxl

cd /d "%~dp0\.."
if not exist "src\measure\isp_orig_data_generator.py" (
    echo [오류] src\measure\isp_orig_data_generator.py 를 찾을 수 없습니다. 프로젝트 루트에서 실행하세요.
    pause
    exit /b 1
)

echo [빌드] Windows용 ISP 원데이터 생성기 ...
py -3 -m PyInstaller scripts\isp_orig_data_generator.spec 2>nul
if errorlevel 1 (
    python -m PyInstaller scripts\isp_orig_data_generator.spec
)
if errorlevel 1 (
    echo PyInstaller를 찾을 수 없으면: pip install pyinstaller openpyxl
    pause
    exit /b 1
)

if exist "dist\ISP_원데이터_생성기.exe" (
    echo.
    echo [완료] dist\ISP_원데이터_생성기.exe
) else (
    echo [확인] dist 폴더에 생성된 실행 파일 이름을 확인하세요.
)
pause
