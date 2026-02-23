# 학교별 AP 장비 현황 엑셀 분리 스크립트

## 필요 패키지 설치
```bash
pip install openpyxl pandas
```

## 사용 방법

### 1. 구조 확인 (선택)
먼저 학교 리스트와 AP 시트 구조를 확인하려면:
```bash
python check_structure.py
```

### 2. 학교별 분리 실행
```bash
python split_school_ap_v1.py
```

**다른 원본 파일 사용 시** (데이터/서식이 다른 파일이면):
```bash
python split_school_ap_v1.py --source "경로\파일명.xlsx"
```

### 3. 빠진 학교만 처리
```bash
# 1) 빠진 학교 리스트 확인/생성
python extract_missed_schools.py

# 2) 빠진 학교만 분리 실행 (기존 로그 유지, 신규 처리분만 추가)
python split_school_ap_v1.py --missed-only
```

## 입력 파일
- **00.충남_AP_자산_첨부(충남전체)_.xlsx** : 원본 (AP(2) 시트)
- **SCHOOL_REG_LIST.XLSX** : 학교 리스트 (A:학교코드, B:지역, C:학교명)

## 출력
- **OUTPUT/지역/학교명_학교코드/학교명_AP 장비 현황 상세.XLSX**
- **split_log_YYYYMMDD.csv** : 처리 로그 (오늘 날짜, 이전 로그 건드리지 않음)

## 데이터 구조 가정
- 1행: 제목
- 2행: 제목행(헤더)
- 3행~: 데이터 (A열에 학교코드, B~L열에 데이터)
- A열의 학교코드가 SCHOOL_REG_LIST와 일치하는 행만 해당 학교 파일로 복사
