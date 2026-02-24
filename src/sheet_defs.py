"""
통합 DB 및 학교별 DB에서 '실데이터' 시트만 구분하여 로딩하기 위한 정의.

- 가상자산: 장비 시트(PoE, AP, 스위치, 보안장비) + 학교정보
- 구성정보: 장비 시트(AP, PoE, 스위치, 보안장비)
- 가상자산 제목행: 관리코드는 원칙상 유일(1개)이어야 함. 원본 오기로 '관리코드'가 2개인 파일이 있는데, 대부분 학교코드를 적어야 할 열에 '관리코드'라고 잘못 적힌 경우임. 그런 자료도 처리할 수 있도록 2개일 때 학교코드-장비 형식이 있는 쪽을 관리번호로 사용.
- 구성정보: 주 연결 장비 식별 컬럼은 '장비관리번호' 또는 '장비관리코드'.
"""

from __future__ import annotations

# 가상자산 DB에서 실데이터로 사용할 시트 (장비 4종 + 학교정보)
VA_DATA_SHEETS_EQUIPMENT = ("PoE", "AP", "스위치", "보안장비")
VA_DATA_SHEETS_ALL = (*VA_DATA_SHEETS_EQUIPMENT, "학교정보")

# 구성정보 DB에서 실데이터로 사용할 시트
CFG_DATA_SHEETS = ("AP", "PoE", "스위치", "보안장비")

# 장비 시트 헤더: 관리번호/관리코드 중 1개는 있어야 함. 구성정보는 장비관리번호 또는 장비관리코드.
VA_HEADER_MUST_HAVE = ("관리번호", "장비명")
CFG_HEADER_MUST_HAVE = ("관리번호", "장비관리번호")  # 장비관리코드도 동일 의미로 사용됨
SCHOOL_INFO_HEADER_MUST_HAVE = ("학교코드", "학교명")

# 학교코드-장비 형식 데이터가 들어 있는 컬럼 후보. VA: 원본 오기로 '관리코드' 2개인 경우, 학교코드 열을 관리코드라고 잘못 쓴 쪽이 있음.
VA_MGMT_COLUMN_CANDIDATES = ("관리번호", "관리코드")
CFG_MGMT_COLUMN_CANDIDATES = ("장비관리번호", "장비관리코드", "관리번호", "관리코드")


def get_va_header_row(sheet_name: str) -> int:
    """가상자산 파일에서 해당 시트의 헤더 행(1-based)."""
    if sheet_name == "학교정보":
        return 3
    return 1


def get_cfg_header_row(sheet_name: str) -> int:
    """구성정보 파일에서 해당 시트의 헤더 행(1-based)."""
    return 1


def get_data_start_row(sheet_name: str, file_kind: str) -> int:
    """헤더 다음 행 = 데이터 시작 행(1-based). file_kind: 'va' | 'cfg'."""
    if file_kind == "va":
        return get_va_header_row(sheet_name) + 1
    return get_cfg_header_row(sheet_name) + 1


def is_va_data_sheet(sheet_name: str) -> bool:
    """가상자산에서 실데이터로 쓸 시트인지 여부."""
    return sheet_name in VA_DATA_SHEETS_ALL


def is_cfg_data_sheet(sheet_name: str) -> bool:
    """구성정보에서 실데이터로 쓸 시트인지 여부."""
    return sheet_name in CFG_DATA_SHEETS


def va_sheets_to_load(sheet_names: list[str]) -> list[str]:
    """가상자산 파일의 시트 목록 중 실데이터 시트만 순서 유지하여 반환."""
    return [s for s in VA_DATA_SHEETS_ALL if s in sheet_names]


def cfg_sheets_to_load(sheet_names: list[str]) -> list[str]:
    """구성정보 파일의 시트 목록 중 실데이터 시트만 순서 유지하여 반환."""
    return [s for s in CFG_DATA_SHEETS if s in sheet_names]


def _series_has_school_code_format(ser: "pd.Series") -> int:
    """시리즈 값 중 'XXX-YYY' 형태(학교코드-장비)가 몇 개인지. 관리코드 컬럼 후보 선택용."""
    import pandas as pd
    if ser is None or ser.empty:
        return 0
    s = ser.dropna().astype(str).str.strip()
    return int((s.str.contains("-", na=False) & (s.str.len() > 2)).sum())


def resolve_mgmt_column_va(df: "pd.DataFrame") -> str | None:
    """
    가상자산 장비 시트에서 학교코드-장비 형식이 들어 있는 컬럼 이름 반환.
    관리코드는 원칙상 유일. 원본 오기로 '관리코드' 2개인 경우, 학교코드 열을 관리코드라고 잘못 적은 쪽이 있으므로
    '관리번호' 우선, 없으면 '관리코드' 중 데이터가 학교코드-장비 형식인 쪽(2개면 '-' 많은 쪽) 선택.
    """
    import pandas as pd
    if df is None or df.empty:
        return None
    if "관리번호" in df.columns:
        col = df["관리번호"]
        if isinstance(col, pd.DataFrame):
            col = col.iloc[:, 0]
        if _series_has_school_code_format(col) > 0:
            return "관리번호"
    for cand in VA_MGMT_COLUMN_CANDIDATES:
        if cand == "관리번호":
            continue
        if cand not in df.columns:
            continue
        col = df[cand]
        if isinstance(col, pd.DataFrame):
            # 제목행에 '관리코드' 2개(원본 오기): 학교코드 열을 관리코드라고 잘못 적은 쪽이 있음. '-' 많은 쪽이 학교코드-장비 형식.
            best_idx = 0
            best_count = 0
            for i in range(col.shape[1]):
                c = col.iloc[:, i]
                n = _series_has_school_code_format(c)
                if n > best_count:
                    best_count = n
                    best_idx = i
            if best_count > 0:
                return cand  # 동일 이름이면 iloc으로 구분해야 하므로 호출측에서 df[cand].iloc[:, best_idx] 사용
        else:
            if _series_has_school_code_format(col) > 0:
                return cand
    return None


def resolve_mgmt_column_cfg(df: "pd.DataFrame") -> str | None:
    """
    구성정보 장비 시트에서 주 연결 장비 식별 컬럼 이름 반환.
    '장비관리번호' 또는 '장비관리코드' 우선, 없으면 '관리번호'·'관리코드' 중 데이터 있는 것.
    """
    import pandas as pd
    if df is None or df.empty:
        return None
    for cand in CFG_MGMT_COLUMN_CANDIDATES:
        if cand not in df.columns:
            continue
        col = df[cand]
        if isinstance(col, pd.DataFrame):
            col = col.iloc[:, 0]
        if _series_has_school_code_format(col) > 0:
            return cand
    return None


def normalize_mgmt_column(df: "pd.DataFrame", file_kind: str) -> "pd.DataFrame":
    """
    VA/CFG 장비 시트에서 실제 사용할 컬럼을 찾아 '관리번호'로 통일.
    VA: 관리코드는 원칙상 유일. 원본 오기로 관리코드 2개인 경우(학교코드 열을 관리코드라고 잘못 적은 경우)도 처리.
    CFG: 장비관리번호/장비관리코드 등.
    """
    import pandas as pd
    if file_kind == "va":
        resolved = resolve_mgmt_column_va(df)
    else:
        resolved = resolve_mgmt_column_cfg(df)
    if resolved is None:
        return df
    if resolved == "관리번호" and "관리번호" in df.columns:
        col = df["관리번호"]
        if isinstance(col, pd.DataFrame):
            df = df.copy()
            df["관리번호"] = col.iloc[:, 0]
        return df
    col = df[resolved]
    if isinstance(col, pd.DataFrame):
        # 관리코드 2개: '-' 많은 쪽
        best_idx = 0
        best_count = 0
        for i in range(col.shape[1]):
            n = _series_has_school_code_format(col.iloc[:, i])
            if n > best_count:
                best_count = n
                best_idx = i
        out = df.copy()
        out["관리번호"] = col.iloc[:, best_idx].values
        return out
    out = df.copy()
    out["관리번호"] = col.values
    return out
