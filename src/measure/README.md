# 학교별 측정 값 현황

문제점 분석 내용을 학교별로 출력하는 프로젝트입니다.

## 프로젝트 목적

- **1단계**: 출력 데이터 정리 및 계산
  - 유선망 측정 데이터 전처리
  - 무선망 데이터 정리
  - 유선망/무선망 데이터 계산
- **2단계**: 데이터 출력
  - 템플릿 파일에 학교별 데이터 입력 및 파일 생성

---

## 유선망 전처리 (1단계-1)

### 작업 순서

1. **유선망 1차 측정 결과 통합**
   - 충남: `CNE/유선망품질측정결과1차_충남/` 폴더의 학교별 엑셀 파일
   - 대전: `DNI/유선망품질측정결과1차_대전/` 폴더의 학교별 엑셀 파일

2. **학교별 평균 계산**
   - 통합 데이터에서 학교별 평균값 계산
   - `CNE_WIRED_MEANSURE_V1.XLSX` (충남) / `DNI_WIRED_MEANSURE_V1.XLSX` (대전) 출력

### 출력 파일 구조 (충남)

| 시트 | 설명 |
|------|------|
| CNE_TOTAL | 통합 원본 데이터 (모든 학교) |
| CNE_WIRED_MEANSURE_AVG | 학교별 평균값 + 학교명, 장비개수, 진단결과 |

**AVG 시트 열 구성**:  
학교코드 | 학교명 | 장비개수 | K열 | L열 | M열 | N열 | 진단결과  
- 학교코드/학교명: CNE_TOTAL 시트 B열, C열과 동일 (소스 1열, 2열)
- 진단결과: K열(Avg Throughput) 700 Mbps 기준 → 양호/미흡

### 폴더 구조

```
src/measure/
├── CNE/
│   ├── 유선망품질측정결과1차_충남/   # 충남 원본 (학교별 파일)
│   └── CNE_WIRED_MEANSURE_V1.XLSX   # 출력
├── DNI/
│   ├── 유선망품질측정결과1차_대전/   # 대전 원본 (학교별 파일)
│   └── DNI_WIRED_MEANSURE_V1.XLSX   # 출력
├── wired_preprocess_config.py
├── wired_preprocess_v1.py           # 유선망 전처리 (충남/대전 공통)
├── run_wired_preprocess.bat       # 상단 RUN_REGION에 따라 실행
├── run_wired_preprocess_충남.bat  # 충남 실행
├── run_wired_preprocess_dni.bat   # 대전 실행
└── analyze_wired_structure.py    # 구조 분석
```

### 의존성

- openpyxl, tqdm (`pip install tqdm`)

### 실행 방법

```bash
# 유선망 전처리 (상단 RUN_REGION에 따라 충남/대전/둘 다)
cd src/measure
python wired_preprocess_v1.py

# 배치 파일 (CMD)
run_wired_preprocess.bat          # RUN_REGION 사용
run_wired_preprocess_충남.bat     # 충남만
run_wired_preprocess_dni.bat      # 대전만

# PowerShell (한글 깨짐 시)
.\run_wired_preprocess.ps1
# 또는: $env:PYTHONIOENCODING="utf-8"; python wired_preprocess_v1.py
```

### 학교코드 추출

- 파일명에서 `G107441266MS` 형식 또는 12자리 숫자 패턴으로 학교코드 추출
- 추출 실패 시 파일명 앞 20자 사용
