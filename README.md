<<<<<<< HEAD
﻿# CNE_DNI

충남/대전 데이터 통합 및 분석 작업용 프로젝트

## 구조
- src: 파이썬 스크립트
- notebooks: 주피터 노트북
- config: 경로/설정 파일
- output: 결과물 (Git 제외)
- logs: 로그 (Git 제외)
=======
# CNE_DNI

충남/대전 지역 학교 유무선 데이터 진단 및 개선 사업 관련
크롤링/정제/분석 작업 프로젝트

## Environment
- Python 3.13
- Virtual Env: .venv
- Main libs: pandas, openpyxl, jupyter

## Start
```bash
source .venv/bin/activate
jupyter lab




### 2) `WORKLOG.md` 기본 템플릿
```bash
cat <<'EOF' > WORKLOG.md
# WORKLOG

## 2026-02-21 (MacBook Air)
- 맥북에어 표준 개발환경 초기 세팅
- Homebrew + Python 3.13 구성
- .venv 생성 및 패키지 설치 (openpyxl, pandas, jupyter)
- Git 초기화 및 .gitignore 설정
- config/output/logs 구조 생성
- paths.example.json / paths.local.json 설정

### Next
- src/, notebooks/ 폴더 생성
- 윈도우 노트북 기존 작업 스크립트 이관
- Google Drive 내 원본 데이터 폴더(CNE_DNI_RAW) 구조 정리
>>>>>>> 7bac18f (초기 프로젝트 구조 및 맥북에어 표준 환경 세팅)
