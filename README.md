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
