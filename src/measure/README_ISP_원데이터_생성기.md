# ISP 1차/2차 원데이터 생성기

기획: [docs/06_원데이터_생성_프로그램_기획.md](../../docs/06_원데이터_생성_프로그램_기획.md)

## 실행 방법

### Python으로 실행

```bash
# 프로젝트 루트에서
python -m src.measure.isp_orig_data_generator

# 또는 src/measure에서
cd src/measure
python isp_orig_data_generator.py
```

실행 폴더(또는 `row_data`/`raw_data`)에 있는 엑셀 파일 목록이 나오면 번호로 파일 선택 후, 시트·열 매핑을 선택하면 됩니다.

### 실행 파일(.exe) 빌드

1. PyInstaller 설치: `pip install pyinstaller`
2. 프로젝트 루트에서:
   ```bash
   pyinstaller scripts/isp_orig_data_generator.spec
   ```
3. 생성 파일: `dist/ISP_원데이터_생성기` (macOS/Linux), Windows에서는 `dist/ISP_원데이터_생성기.exe`

**주의**: 빌드 후에도 `src/measure/` 소스 파일은 삭제하지 마세요. 수정·재빌드 시 필요합니다.

### Windows용 .exe 빌드 (Windows PC에서만 가능)

PyInstaller는 **크로스 컴파일을 지원하지 않아**, Windows용 `.exe`는 **Windows에서** 빌드해야 합니다.

1. **Windows PC**에서 프로젝트 폴더(CNE_DNI)를 연다.
2. Python 3 설치 후 터미널(명령 프롬프트 또는 PowerShell)에서:
   ```bat
   cd C:\경로\CNE_DNI
   py -3 -m pip install pyinstaller openpyxl
   py -3 -m PyInstaller scripts\isp_orig_data_generator.spec
   ```
3. 또는 **배치 파일** 실행:
   ```bat
   cd C:\경로\CNE_DNI
   scripts\build_isp_orig_generator_win.bat
   ```
4. 생성 파일: `dist\ISP_원데이터_생성기.exe`

### 실행 파일로 배포했을 때 (다른 사람에게 전달)

- **데이터 위치**: 실행 파일(**ISP_원데이터_생성기.exe** 또는 Mac 실행 파일)이 있는 **같은 폴더**에 기초 데이터 엑셀(.xlsx/.xlsm)을 넣어 두세요.
- **Windows·Mac 공통**: 프로그램은 “실행 파일이 있는 폴더”만 읽어서 파일 목록을 보여 줍니다. 생성된 1차/2차 원데이터 엑셀도 기본적으로 이 폴더에 저장됩니다.
- 예: `C:\ISP도구\` 폴더에 `ISP_원데이터_생성기.exe`와 `충남_무선속도_통합_머지파일_....xlsm`을 함께 두고 실행하면 됩니다.

## 입력

- **기초 데이터**: 1차/2차 **평균** 데이터가 있는 엑셀 파일 (예: 충남_무선속도_통합_머지파일_보고서용_ver38_20260304_전부하포함_Y1.xlsm)
- 사용자가 **파일 → 시트 → 열 매핑**(학교코드, 학교명, 장비, 다운로드, 업로드, RTT, RSSI, CH) 선택

## 출력

- `ISP_원데이터_1차_YYYYMMDD_HHMMSS.xlsx`
- `ISP_원데이터_2차_YYYYMMDD_HHMMSS.xlsx`
- 형식: ISP_샘플.xlsx와 동일 (Date, StartTime, 학교, 메모1, 측정순번, DL, UL, RTT, RSSI, CH, 학교코드 등)

## 규칙 요약

- 장비당 6회 측정(DL 3회, UL 3회) → 평균이 입력 시트의 해당 행과 일치하도록 생성
- DL/UL: 10~60 Mbps, RTT 0.7, RSSI -1, CH 동일
- 시간: 14:00 시작, 학교 간 30분, 장비 간 3~4분, 회차 간 29초, DL-UL 15초, **17:00 전** 종료
