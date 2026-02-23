# -*- coding: utf-8 -*-
"""
학교별 측정 값 현황 - 설정

[폴더 규칙]
- BASE_DIR: 이 파일 기준 작업 폴더
- DNI_DIR: BASE_DIR/DNI (대전 원본)
- CNE_DIR: BASE_DIR/CNE (충남 원본)

[측정 데이터 규칙]
- 기존 장비 현황(AP, 스위치, 보안, POE)과 동일한 구조 사용
- split 패키지의 school_utils, split_config 규칙 재사용 가능
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DNI_DIR = os.path.join(BASE_DIR, "DNI")
CNE_DIR = os.path.join(BASE_DIR, "CNE")

# ========== 측정 데이터 원본 경로 ==========
SOURCE_DIR_BY_REGION = {
    "DNI": DNI_DIR,
    "CNE": CNE_DIR,
}

# ========== 출력 경로 (config/paths 활용 시) ==========
# 프로젝트 루트 config/paths.local.json 의 OUT_ROOT, LOG_ROOT 사용 가능
