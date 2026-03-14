@echo off
chcp 65001 >nul
REM Windows용 BSSID 매칭 실행 파일 빌드
REM 사용: 프로젝트 루트(CNE_DNI)에서 scripts\build_bssid_csv_win.bat 실행
REM 필요: Python (PyInstaller는 표준 라이브러리만 사용)

cd /d "%~dp0\.."
if not exist "src\measure\build_bssid_csv.py" (
    echo [오류] src\measure\build_bssid_csv.py 를 찾을 수 없습니다. 프로젝트 루트에서 실행하세요.
    pause
    exit /b 1
)

echo [빌드] Windows용 BSSID_매칭 ...
py -3 -m PyInstaller scripts\build_bssid_csv.spec 2>nul
if errorlevel 1 (
    python -m PyInstaller scripts\build_bssid_csv.spec
)
if errorlevel 1 (
    echo PyInstaller를 찾을 수 없으면: pip install pyinstaller
    pause
    exit /b 1
)

if exist "dist\BSSID_매칭.exe" (
    echo.
    echo [완료] dist\BSSID_매칭.exe
) else (
    echo [확인] dist 폴더에 생성된 실행 파일 이름을 확인하세요.
)
pause
