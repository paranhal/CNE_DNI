# CNE_DNI

충남/대전 지역 학교 유무선 데이터 진단 및 개선 사업 관련
크롤링/정제/분석 작업 프로젝트

## Environment
- Python 3.13
- Virtual Env: .venv
- Main libs: pandas, openpyxl, jupyter, tqdm

## 프로젝트

| 프로젝트 | 설명 | 경로 |
|----------|------|------|
| **학교별 측정 값 현황** | 대전/충남 학교별 측정 데이터 분리·관리 | `src/measure/` |
| **학교별 장비 분리** | AP, 스위치, 보안, POE 장비 현황 분리 | `src/split/` |

## 구조
- src: 파이썬 스크립트
  - measure: 학교별 측정 값 현황
  - split: 학교별 장비 분리
  - common: 공통 모듈
- config: 경로/설정 파일
- output: 결과물 (Git 제외)
- logs: 로그 (Git 제외)

## Start
source .venv/bin/activate
jupyter lab

## 장비별·가상/구성별 파일 생성 기반
- **대상 리스트**: `output/CNE_LIST.xlsx` (충남 작업 기준)
- **장비**: PoE, AP, 스위치, 보안장비 (정의: `src/sheet_defs.py` → `src/export_config.py`)
- **출력**: `df_{장비}_가상자산.csv`, `df_{장비}_구성정보.csv` (필터 후 `output/`에 저장)
- **설정**: `config/paths.local.json`에 `va_file`, `cfg_file` 지정 시 해당 경로 사용 (비면 기본 경로)
- **실행**: `python -m src.integrate_export` → 통합 VA/CFG 로드 → 대상 학교 필터 → 장비별 CSV 8개 + 로그
