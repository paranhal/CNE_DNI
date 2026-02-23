# 학교별 장비 현황 엑셀 분리 스크립트

대전(DNI) / 충남(CNE) 지역 학교별 AP, 스위치, 보안, POE 장비 현황을 원본 엑셀에서 분리하는 통합 스크립트입니다.

---

## 1. 환경 셋팅

### 1.1 필요 환경
- **Python 3.x** (openpyxl 패키지 필요)
- Windows / macOS / Linux

### 1.2 패키지 설치
```bash
pip install openpyxl
```

### 1.3 폴더 구조
```
기호준_측정값/
├── split_school_all_v1.py   # 통합 실행 스크립트
├── verify_missing_by_code.py # 빠진 학교 검증
├── split_config.py           # 원본 경로·시트 규칙 (수정용)
├── school_utils.py           # 학교 리스트·유틸
├── DNI/                      # 대전 원본 (AP, 스위치, 보안, POE)
├── CNE/                      # 충남 원본 (선택)
├── OUTPUT/                   # 테스트 출력 (--test 시)
│   ├── DNI/                  # 대전 테스트
│   │   └── {시군구}/{학교명_코드}/
│   └── CNE/                  # 충남 테스트
│       └── {시군구}/{학교명_코드}/
├── SCHOOL_REG_LIST_DNI.xlsx  # 대전 학교 리스트 (또는 .csv)
├── SCHOOL_REG_LIST_CNE.xlsx  # 충남 학교 리스트 (또는 .csv)
└── missed_schools.csv        # 빠진 학교 목록 (선택)
```

### 1.4 원본 파일 배치 (split_config.py 규칙)

| 지역 | 장비 | 폴더 | 파일명 |
|------|------|------|--------|
| DNI | AP | DNI/ | DNI_AP_LIST.XLSX |
| DNI | 스위치 | DNI/ | DNI_SWITCH_LIST.xlsx |
| DNI | 보안 | DNI/ | DNI_SEUTM_LIST.xlsx |
| DNI | POE | DNI/ | DNI_POE_LIST.xlsx |
| CNE | AP | CNE/ | CNE_AP_LIST.xlsx |
| CNE | 스위치 | CNE/ | CNE_SWITCH_LIST.xlsx |
| CNE | 보안 | CNE/ | CNE_SEUTM_LIST.xlsx |
| CNE | POE | CNE/ | CNE_POE_LIST.xlsx |

### 1.5 학교 리스트 파일
- **대전**: `SCHOOL_REG_LIST_DNI.xlsx` 또는 `school_reg_list_DNI.csv` (작업폴더)
- **충남**: `SCHOOL_REG_LIST_CNE.xlsx` 또는 `school_reg_list_CNE.csv` (작업폴더)
- **열 구성**: 학교코드 | 지역(구/시군) | 학교명 (1행은 헤더 가능)

### 1.6 출력 경로 (split_config.py)
- **실제 저장**: `OUTPUT_BASE_BY_REGION` (Y:\...\DJE, Y:\...\CNE 등)
- **테스트 저장**: `OUTPUT_BASE_TEST` → `OUTPUT/DNI/`, `OUTPUT/CNE/`

---

## 2. 옵션 사용법

### 2.1 기본 실행
```bash
python split_school_all_v1.py --{장비} --{지역} [옵션]
```

### 2.2 장비 선택 (필수, 하나만)
| 옵션 | 설명 | 대소문자 |
|------|------|----------|
| `--AP` | AP 장비 | 무관 |
| `--switch` | 스위치 | 무관 |
| `--security` | 보안(SEUTM) | 무관 |
| `--poe` | POE | 무관 |
| `-e AP` | `-e`로 장비 지정 | 무관 |

### 2.3 지역 선택
| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `--DNI` | 대전 | DNI |
| `--CNE` | 충남 | |

### 2.4 기타 옵션
| 옵션 | 설명 |
|------|------|
| `--test`, `-t` | 테스트: OUTPUT/DNI 또는 OUTPUT/CNE에 저장 |
| `--missed-only` | 로그에 없는(빠진) 학교만 처리 |
| `--only-schools 108140237,108140238` | 지정 학교코드만 처리 (쉼표 구분) |
| `--schools-file 파일경로` | 학교코드 목록 파일 (한 줄에 하나 또는 CSV) |
| `--today-from-missed` | missed_schools.csv에 있는 학교만 처리 |
| `--from-log split_log_AP_DNI_20260222.csv` | 지정 로그 파일의 학교만 처리 |
| `--new-log` | 이번 실행만 별도 로그 파일 생성 (기존 로그에 추가 안 함) |
| `--source`, `-s 경로` | 원본 엑셀 파일 직접 지정 |

### 2.5 실행 예시
```bash
# 대전 AP 테스트
python split_school_all_v1.py --AP --DNI --test

# 충남 스위치 실제 저장
python split_school_all_v1.py --switch --CNE

# 대전 보안, 빠진 학교만
python split_school_all_v1.py --security --DNI --missed-only

# -e 옵션 사용 (대소문자 무관)
python split_school_all_v1.py -e poe --cne --test
python split_school_all_v1.py -e POE --CNE --missed-only
```

### 2.6 배치 파일 (run_split_all.bat)
```batch
run_split_all.bat {장비} {지역} [추가옵션]

REM 예시
run_split_all.bat AP DNI --test
run_split_all.bat switch CNE --missed-only
run_split_all.bat security DNI
```

---

## 3. 로그 파일
- **형식**: `split_log_{장비}_{지역}_{날짜}.csv`
- **예**: `split_log_AP_DNI_20260222.csv`, `split_log_switch_CNE_20260222.csv`
- **내용**: 학교명, 학교코드, 지역, 시작행, 끝행, 복사행수, 저장경로

---

## 4. 빠진 학교 검증 (verify_missing_by_code.py)

학교코드 기준으로 split 로그에 없는(빠진) 학교를 검증합니다. **지역·장비 옵션 필수** (디폴트 없음).

### 4.1 옵션
| 구분 | 옵션 | 설명 |
|------|------|------|
| 지역 (필수) | `--DNI`, `--dni` | 대전 (DJE) |
| | `--CNE`, `--cne` | 충남 |
| 장비 (필수) | `--AP`, `--ap` | AP |
| | `--switch`, `--Switch`, `--SWITCH` | 스위치 |
| | `--security`, `--Security`, `--SECURITY` | 보안(SEUTM) |
| | `--poe`, `--Poe`, `--POE` | POE |
| | `-e AP` | `-e`로 장비 지정 (대소문자 무관) |

### 4.2 실행 예시
```bash
python verify_missing_by_code.py --DNI --AP
python verify_missing_by_code.py --CNE --switch
python verify_missing_by_code.py --dni -e security
python verify_missing_by_code.py --cne --poe
```

---

## 5. 설정 수정 (split_config.py)
- **원본 경로/파일명**: `SOURCE_FILENAME_RULES` 수정
- **시트 우선순위**: `SHEET_PRIORITY_BY_EQUIPMENT` 수정
- **출력 경로**: `OUTPUT_BASE_BY_REGION`, `OUTPUT_BASE_TEST` 수정
