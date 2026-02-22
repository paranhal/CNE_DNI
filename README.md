# CNE_DNI

충남/대전 지역 학교 유무선 데이터 진단 및 개선 사업 관련
크롤링/정제/분석 작업 프로젝트

## Environment
- Python 3.13
- Virtual Env: .venv
- Main libs: pandas, openpyxl, jupyter, tqdm

## Start
source .venv/bin/activate
jupyter lab

## 장비별·가상/구성별 파일 생성 기반
- **대상 리스트**: `output/CNE_LIST.xlsx` (충남 작업 기준)
- **장비**: PoE, AP, 스위치, 보안장비 (정의: `src/sheet_defs.py` → `src/export_config.py`)
- **출력**: `df_{장비}_가상자산.csv`, `df_{장비}_구성정보.csv` (필터 후 `output/`에 저장)
- **설정**: `config/paths.local.json`에 `va_file`, `cfg_file` 지정 시 해당 경로 사용 (비면 기본 경로)
- **실행**: `python -m src.integrate_export` → 통합 VA/CFG 로드 → 대상 학교 필터 → 장비별 CSV 8개 + 로그

## Structure
- src/ : Python scripts
- notebooks/ : Jupyter notebooks
- config/ : local/shared path configs
- output/ : outputs (git ignored)
- logs/ : logs (git ignored)
