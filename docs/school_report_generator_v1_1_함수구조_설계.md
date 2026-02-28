# school_report_generator_v1_1.py 함수 구조 재설계 (설계 전용)

- **대상 파일**: `src/measure/school_report_generator_v1_1.py`
- **실행 방식 유지**: `python -m src.measure.school_report_generator_v1_1` 또는 `cd src/measure && python school_report_generator_v1_1.py` 와 충돌 없음
- **이 문서**: 코드 작성 없이 함수 역할 재분류, 문제점 분석, 새 구조 제안, 호출 흐름만 정리

---

## 1. 현재 파일의 역할을 함수 단위로 재분류

| 구분 | 함수명 | 현재 역할 (한 줄) | 비고 |
|------|--------|-------------------|------|
| **시트/셀** | `_norm_sheet_name` | 시트명 정규화(비교용) | 순수 유틸 |
| | `_resolve_sheet_name` | 설정 시트명 → 실제 시트명 매핑 | alias/키워드 보정 포함 |
| | `_pick_best_row` | 다중행 중 “값이 많은” 행 선택 | row_def 기반 |
| | `_find_header_col` | 헤더 키워드로 열 번호 탐색 | |
| | `_resolve_isp_cols_by_header` | ISP 시트 DL/UL/RSSI 열 보정 | _find_header_col 사용 |
| **파일/경로** | `sanitize_filename` | 파일명 불가 문자 제거 | |
| | `find_template` | 템플릿 xlsx 경로 탐색 | 여러 후보 순회 |
| | `resolve_total_measure_path` | 통계 파일 기본 경로 | |
| | `resolve_output_dir` | 출력 디렉터리 기본값 | |
| | `resolve_log_dir` | 로그 디렉터리 기본값 | |
| **사용자 입력** | `_input` | 안전한 input, EOF/중단 시 exit | |
| | `_discover_total_measure_candidates` | 통계 파일 후보 목록 수집 | resolve 후보와 중복 |
| | `select_total_measure_path` | 사용자 선택/직접입력으로 통계 경로 결정 | |
| | `select_output_dir` | 출력 폴더 사용자 지정 | |
| | `select_output_layout` | region / flat 선택 | |
| **학교 목록** | `load_full_school_list` | 외부 파일에서 719개 학교 목록 로드 | CSV/Excel 분기, 헤더 탐색 중복 |
| | `load_school_meta_from_sheet1` | 통계 wb의 Sheet1에서 코드/이름/지역 로드 | |
| | `load_stats_by_school` | 통계 wb에서 학교코드별 (시트→행) 인덱스 | |
| | `get_school_name_from_stats` | 학교명 없을 때 통계 시트에서 이름 보강 | |
| **측정값** | `get_school_values` | 한 학교·한 row_def에 대한 원시값 (v1,v2) | 시트/행/열/ISP보정 모두 포함 |
| | `get_numeric` | 값 → 숫자 변환 | |
| | `format_value` | 단일 값 → 출력용 문자열/숫자 | |
| | `format_output_value` | v1,v2+row_def → J열용 문자열 | fmt 타입별 분기 다수 |
| **판정** | `judge` | (val, op, threshold, val2) → 정상/개선필요 | op 종류 많음 |
| **리포트** | `_fill_measurement_and_judgment` | J열·L열 채우기 | J_OUTPUT_MAP/L_JUDGMENT_MAP 순회 |
| | `_fill_summary_cells` | G21, G36/L36, G37/L37 | |
| | `generate_school_report` | 템플릿 로드 + 위 두 함수 호출 후 wb 반환 | |
| **오케스트** | `_setup_paths` | 템플릿/통계/출력/로그/레이아웃 결정 | 실패 시 sys.exit |
| | `_load_workbook_and_school_lists` | wb 로드 + 학교 목록·메타·by_school | |
| | `_write_missing_schools_log` | 통계 없는 학교 로그 기록 | |
| | `_generate_and_save_all` | 학교별 생성·저장 루프, 건수/제외 목록 반환 | |
| | `_append_skipped_log` | 제외 학교 로그 추가 | |
| | `main` | 진입점, 위 단계 순서 호출 | |

---

## 2. 문제점: 과대 함수, 중복, 책임 혼재

### 2.1 너무 큰 함수

| 함수 | 예상 라인 수 | 문제 |
|------|--------------|------|
| `load_full_school_list` | ~60 | CSV/Excel 두 경로에서 **헤더 찾기 + 행 순회**가 거의 동일하게 반복됨. 한쪽만 수정 시 다른 쪽 누락 위험. |
| `format_output_value` | ~50 | `fmt` 타입별 분기가 한 함수에 모두 있음. 타입 추가 시 계속 길어짐. |
| `judge` | ~50 | `op` 종류가 많아 한 함수에 모든 분기. 단위 테스트는 쉬우나 가독성·확장 시 부담. |
| `_resolve_sheet_name` | ~50 | 정규 매칭 + alias_map + 키워드 폴백이 한 덩어리. |

### 2.2 중복 로직

| 내용 | 위치 | 제안 |
|------|------|------|
| **통계 파일 후보 목록** | `resolve_total_measure_path`의 candidates, `_discover_total_measure_candidates`의 candidates | 동일 리스트가 두 군데 하드코딩. 한 곳(상수 또는 `_get_total_measure_candidate_list()`)에서만 정의하고, resolve는 “첫 존재 경로”, discover는 “존재하는 것만 수집”으로 역할 분리. |
| **헤더에서 열 인덱스 찾기** (학교코드/학교명/지역) | `load_full_school_list`(CSV·Excel 각각), `load_school_meta_from_sheet1`, `get_school_name_from_stats`, `load_stats_by_school` | “헤더 행에서 키워드로 열 번호 찾기”가 여러 형태로 반복. `_find_header_col`과 유사한 **공통 헤더→열 매핑** 유틸 하나로 묶을 수 있음. |
| **지역 정규화** (빈 값 → "미분류") | `_generate_and_save_all` 내 `region_totals` 계산과 `out_path` 계산 시 | `region = code_to_region.get(sc,"").strip() or "미분류"` 반복. `_normalize_region(region)` 같은 한 줄 함수로 통일. |

### 2.3 책임이 섞인 부분

| 함수 | 혼재 내용 | 제안 |
|------|-----------|------|
| `get_school_values` | 시트 해석 + 행 선택 + 열 해석(일반/ISP) + fixed·h_only 예외 | “시트/행 결정”과 “해당 행에서 열 읽기”를 나누면 테스트·재사용 쉬움. |
| `_fill_measurement_and_judgment` | J열 쓰기 + L열 판정 + 폰트 설정 | “한 row_def에 대한 J값·L값 계산”과 “ws에 쓰기+스타일” 분리 가능. |
| `_setup_paths` | 경로 해석 + 사용자 선택 + **실패 시 sys.exit** | “경로 후보/기본값 계산”과 “대화형 선택 + 실패 처리”를 나누면 비대화형/테스트 시 유리. |
| `_load_workbook_and_school_lists` | 파일 검증 + wb 로드 + **Sheet1 vs 외부 목록** 분기 + by_school 로드 + **메시지/exit** | “메타 소스 결정(Sheet1 우선)”과 “wb 로드 + by_school 로드”를 단계로 나누고, exit/print는 오케스트레이션 쪽에만 두는 편이 명확. |

---

## 3. 제안: 새 함수 구조

- **진입점**: `main()` 유지. `python -m src.measure.school_report_generator_v1_1` 동작 변경 없음.
- **모듈 레벨**: 상수·경로 후보는 상단 또는 전용 함수로 한 곳만 정의.

### 3.1 역할별 블록 (제안)

| 블록 | 역할 | 함수 (신규/이름 변경 포함) |
|------|------|----------------------------|
| **A. 시트/셀 유틸** | 시트명·행·열 해석 | `_norm_sheet_name`, `_resolve_sheet_name`, `_pick_best_row`, `_find_header_col`, `_resolve_isp_cols_by_header` (유지) |
| **B. 경로/파일** | 파일명·경로 해석만 (사용자 입력 없음) | `sanitize_filename`, `find_template`, `resolve_total_measure_path`, `resolve_output_dir`, `resolve_log_dir`, **`_get_total_measure_candidate_list()`** (후보 리스트 단일 정의, resolve/discover에서 사용) |
| **C. 사용자 입력** | 프롬프트·선택 | `_input`, `select_total_measure_path`, `select_output_dir`, `select_output_layout` / **`_discover_total_measure_candidates`** 는 B의 후보 리스트 사용 |
| **D. 학교 목록** | 목록·메타·인덱스 로드 | **`_parse_school_table(rows, header_row, path_is_csv)`** 같은 공통 “헤더 찾기 + (코드,이름,지역) 행 반환” 도입 후, `load_full_school_list`는 “파일 찾기 + _parse_school_table 호출”만 담당. `load_school_meta_from_sheet1`, `load_stats_by_school`, `get_school_name_from_stats` 유지 또는 D 전용 헤더 유틸 사용. |
| **E. 측정값** | 값 추출·포맷 | **값 추출**: `get_school_values` 유지하되, 내부에서 **`_resolve_data_row`**(시트+행 결정), **`_read_cell_values(ws, data_row, row_def)`**(열 읽기)로 분리 권장. **포맷**: `get_numeric`, `format_value`, `format_output_value` 유지. `format_output_value`는 **`_format_output_*`** (location5, n_ac_ax, cable, fullload, 기본)으로 fmt별 작은 함수로 쪼개면 확장·테스트 용이. |
| **F. 판정** | 정상/개선필요 | `judge` 유지. (선택) op별로 `_judge_ge`, `_judge_le` 등으로 나누고 `judge`는 디스패처만 두는 방식 가능. |
| **G. 리포트 시트 채우기** | 워크시트에 쓰기 | **`_compute_row_output(wb_stats, school_data, row_def)`** → (j_val, v1, v2) 반환. **`_compute_judgment_for_row(row, v1, v2, l_map)`** → L열 값. **`_fill_measurement_and_judgment`** 는 위 두 개 호출하고 ws에 쓰기+폰트만. `_fill_summary_cells`, `generate_school_report` 유지. |
| **H. 오케스트레이션** | 실행 순서·exit·로그 | **`_ensure_template_path()`**: find_template + 없으면 exit (경로 메시지만). **`_setup_paths()`**: 위 + select_* 호출해 5-tuple 반환. **`_load_workbook_and_school_lists`**: 파일 검증 + wb + 메타/목록 + by_school, exit/print는 호출자(main) 또는 최소한으로. **`_write_missing_schools_log`**, **`_generate_and_save_all`**, **`_append_skipped_log`** 유지. **`_normalize_region(region)`** 로 지역 문자열 통일. **`main`**: 단계만 순서대로 호출. |

### 3.2 새로 도입 권장 함수 요약

| 함수 (제안) | 역할 | 입력 | 반환 |
|-------------|------|------|------|
| `_get_total_measure_candidate_list()` | 통계 파일 후보 경로 리스트 (한 곳 정의) | 없음 | `list[str]` |
| `_normalize_region(region)` | 지역 빈 값 → "미분류" | `str` | `str` |
| `_parse_school_table(rows, header_row, path_is_csv)` 또는 유사 | 헤더에서 코드/이름/지역 열 찾고, 행 리스트 → (codes, code_to_name, code_to_region) | 행/헤더, csv 여부 | `(list, dict, dict)` |
| `_resolve_data_row(wb_stats, school_data, row_def)` | 시트+행만 결정 (fixed/h_only 포함) | wb_stats, school_data, row_def | `(ws, data_row)` 또는 None |
| `_read_cell_values(ws, data_row, row_def)` | 한 행에서 row_def에 따른 (v1,v2) 읽기 | ws, data_row, row_def | `(v1, v2)` |
| `_compute_row_output(wb_stats, school_data, row_def)` | J열 문자열 + 판정용 (v1,v2) | wb_stats, school_data, row_def | `(j_str, v1, v2)` |
| `_compute_judgment_for_row(row, v1, v2, l_map)` | L열 판정 문자열 | row, v1, v2, l_map | `str` |
| `_ensure_template_path()` | find_template 호출, 없으면 print+sys.exit(1) | 없음 | `str` (template_path) |
| (선택) `_format_output_location5`, `_format_output_cable` 등 | fmt별 J열 문자열 | v1, v2 등 | `str` |

---

## 4. 함수별 역할·입력·반환 정리 (제안 구조 반영)

### 4.1 시트/셀 유틸 (A)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| `_norm_sheet_name(name)` | 시트명 비교용 정규화 | `str` | `str` |
| `_resolve_sheet_name(wb_stats, target_name)` | 설정 시트명 → 실제 시트명 | wb, 시트명 | `str \| None` |
| `_pick_best_row(ws, rows, row_def)` | 다중행 중 매핑 열 기준 최적 행 | ws, 행 리스트, row_def | `int \| None` |
| `_find_header_col(ws, include, exclude)` | 헤더 키워드로 열 번호 | ws, 포함/제외 키워드 | `int \| None` |
| `_resolve_isp_cols_by_header(ws)` | ISP 시트 DL/UL/RSSI 열 | ws | `dict` (행→열) |

### 4.2 경로/파일 (B)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| `sanitize_filename(s)` | 파일명 불가 문자 제거 | `str` | `str` |
| `find_template()` | 템플릿 xlsx 경로 탐색 | 없음 | `str \| None` |
| `_get_total_measure_candidate_list()` | 통계 파일 후보 리스트 (단일 정의) | 없음 | `list[str]` |
| `resolve_total_measure_path()` | 후보 중 첫 존재 경로 | 없음 | `str` |
| `resolve_output_dir()` | 출력 디렉터리 기본값 | 없음 | `str` |
| `resolve_log_dir()` | 로그 디렉터리 기본값 | 없음 | `str` |

### 4.3 사용자 입력 (C)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| `_input(prompt)` | 안전 input, EOF/중단 시 exit | `str` | `str` |
| `_discover_total_measure_candidates()` | 존재하는 통계 파일 목록 (B 후보 사용) | 없음 | `list[str]` |
| `select_total_measure_path()` | 사용자 선택/직접입력으로 통계 경로 | 없음 | `str` |
| `select_output_dir()` | 출력 폴더 사용자 지정 | 없음 | `str` |
| `select_output_layout()` | region / flat | 없음 | `"region" \| "flat"` |

### 4.4 학교 목록 (D)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| `_parse_school_table(...)` (제안) | 헤더에서 코드/이름/지역 열 찾아 행 파싱 | rows, header, csv 여부 | `(codes, code_to_name, code_to_region)` |
| `load_full_school_list()` | 외부 파일에서 학교 목록 로드 | 없음 | `(list, dict, dict)` |
| `load_school_meta_from_sheet1(wb_stats)` | Sheet1에서 코드/이름/지역 | wb | `(list, dict, dict)` |
| `load_stats_by_school(wb_stats)` | 학교코드별 시트→행 인덱스 | wb | `dict[code -> {sheet -> [rows]}]` |
| `get_school_name_from_stats(wb_stats, school_data)` | 통계 시트에서 학교명 보강 | wb, school_data | `str` |

### 4.5 측정값 (E)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| `get_school_values(wb_stats, school_code, school_data, row_def)` | 한 학교·한 row_def 원시값 | wb, code, data, row_def | `(v1, v2)` |
| (제안) `_resolve_data_row(...)` | 시트+데이터 행 결정 | wb_stats, school_data, row_def | `(ws, row) \| None` |
| (제안) `_read_cell_values(ws, row, row_def)` | 한 행에서 열 값 읽기 | ws, row, row_def | `(v1, v2)` |
| `get_numeric(val)` | 값 → 숫자 | any | `float \| None` |
| `format_value(v)` | 단일 값 → 출력용 | any | str/숫자 |
| `format_output_value(v1, v2, row_def)` | J열 출력 문자열 | v1, v2, row_def | `str` |

### 4.6 판정 (F)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| `judge(val, op, threshold, val2)` | 정상/개선필요 판정 | val, op, threshold, val2? | `str` |

### 4.7 리포트 시트 채우기 (G)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| (제안) `_compute_row_output(wb_stats, school_data, row_def)` | J열 문자열 + 판정용 v1,v2 | wb, school_data, row_def | `(j_str, v1, v2)` |
| (제안) `_compute_judgment_for_row(row, v1, v2, l_map)` | L열 판정 문자열 | row, v1, v2, l_map | `str` |
| `_fill_measurement_and_judgment(ws, wb_stats, school_data)` | J열·L열 쓰기+스타일 | ws, wb, school_data | 없음 |
| `_fill_summary_cells(ws)` | G21, G36/L36, G37/L37 | ws | 없음 |
| `generate_school_report(template_path, wb_stats, school_code, school_data)` | 템플릿 로드 후 채우기 | 경로, wb, code, data | wb |

### 4.8 오케스트레이션 (H)

| 함수 | 역할 | 입력 | 반환 |
|------|------|------|------|
| (제안) `_ensure_template_path()` | 템플릿 경로 확인, 없으면 exit | 없음 | `str` |
| `_setup_paths()` | 경로·레이아웃 사용자와 결정 | 없음 | `(template_path, total_measure_path, output_dir, log_dir, output_layout)` |
| `_load_workbook_and_school_lists(total_measure_path)` | wb + 학교 목록·메타·by_school | 경로 | `(wb_stats, all_schools, code_to_name, code_to_region, by_school)` |
| `_write_missing_schools_log(log_path, ...)` | 통계 없는 학교 로그 | log_path, all_schools, schools_with_data, code_to_name | `list` (missing) |
| (제안) `_normalize_region(region)` | 지역 빈 값 → "미분류" | `str` | `str` |
| `_generate_and_save_all(...)` | 학교별 생성·저장 루프 | template_path, wb_stats, output_dir, output_layout, by_school, code_to_name, code_to_region | `(generated_count, skipped_codes)` |
| `_append_skipped_log(log_path, skipped_codes)` | 제외 학교 로그 추가 | log_path, list | 없음 |
| `main()` | 진입점, 단계 순서 호출 | 없음 | 없음 |

---

## 5. 호출 흐름 (텍스트 순서도)

```
main()
├── _setup_paths()
│   ├── find_template()                    → template_path (없으면 _ensure_template_path()에서 exit)
│   ├── select_total_measure_path()
│   │   ├── _discover_total_measure_candidates()  [→ _get_total_measure_candidate_list() 사용]
│   │   └── _input(...)
│   ├── select_output_dir()  → _input(...)
│   ├── select_output_layout()  → _input(...)
│   └── resolve_log_dir()
│   → (template_path, total_measure_path, output_dir, log_dir, output_layout)
│
├── _load_workbook_and_school_lists(total_measure_path)
│   ├── load_workbook(total_measure_path)
│   ├── load_school_meta_from_sheet1(wb_stats)   → sheet1 있으면 메타 사용
│   ├── load_full_school_list()                  → 없으면 외부 목록 [→ _parse_school_table 제안]
│   └── load_stats_by_school(wb_stats)
│   → (wb_stats, all_schools, code_to_name, code_to_region, by_school)
│
├── os.makedirs(log_dir); log_path = ...
├── _write_missing_schools_log(log_path, all_schools, schools_with_data, code_to_name)  → missing
│
├── _generate_and_save_all(template_path, wb_stats, output_dir, output_layout, by_school, code_to_name, code_to_region)
│   ├── region_totals 계산 [→ _normalize_region(region) 사용]
│   └── for school_code in save_codes:
│         ├── code_to_name.get / get_school_name_from_stats(wb_stats, by_school[code])
│         ├── sanitize_filename(school_name)
│         ├── generate_school_report(template_path, wb_stats, school_code, school_data)
│         │   ├── load_workbook(template_path); ws = ...
│         │   ├── _fill_measurement_and_judgment(ws, wb_stats, school_data)
│         │   │   └── for row_def in J_OUTPUT_MAP:
│         │   │         ├── get_school_values(wb_stats, ..., school_data, row_def)  [→ _resolve_data_row, _read_cell_values 분리 제안]
│         │   │         ├── format_output_value(v1, v2, row_def)
│         │   │         ├── ws.cell(J_COL) = out_val; 폰트
│         │   │         ├── judge(...)  → L열 값
│         │   │         └── ws.cell(L_COL) = result; 폰트
│         │   └── _fill_summary_cells(ws)
│         │   → wb
│         ├── out_path = output_dir / [region_dir] / f"{code}_{safe_name}.xlsx"  [→ _normalize_region]
│         ├── wb.save(out_path); wb.close()
│         └── (generated_count, skipped_codes)
│
├── wb_stats.close()
├── _append_skipped_log(log_path, skipped_codes)
└── print 완료 메시지
```

- **공통 의존**: `school_report_config_v1_1` (J_OUTPUT_MAP, L_JUDGMENT_MAP, 열 상수 등) — 변경 없음.
- **실행**: `if __name__ == "__main__": main()` 유지하므로 `python -m src.measure.school_report_generator_v1_1` 그대로 사용 가능.

---

## 6. 정리

- **재분류**: 현재 8개 블록(시트/경로/입력/학교목록/측정값/판정/리포트/오케스트) 유지, 내부에서 **경로 후보 단일화**, **헤더 파싱 공통화**, **지역 정규화** 도입.
- **과대·중복·혼재**: `load_full_school_list` 파싱 분리, `format_output_value` fmt별 분리, `get_school_values` 시트/행 vs 열 읽기 분리, 오케스트레이션의 exit/print 최소화로 책임을 나눔.
- **새 함수**: `_get_total_measure_candidate_list`, `_normalize_region`, `_parse_school_table`, `_resolve_data_row`, `_read_cell_values`, `_compute_row_output`, `_compute_judgment_for_row`, `_ensure_template_path` (및 선택적 `_format_output_*`, `_judge_*`) 제안.
- **호출 흐름**: `main` → 경로 설정 → 데이터 로드 → 누락 로그 → 학교별 생성·저장 → 제외 로그 → 완료 출력.

이 설계대로 구현 시 기존 실행 방식 및 config 의존성과 충돌 없이, 테스트·확장이 쉬운 구조를 만들 수 있습니다.
