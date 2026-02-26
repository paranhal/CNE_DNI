# -*- coding: utf-8 -*-
"""
DNI 측정 자료 빌더(대체본)

기존 interactive builder 소스가 없는 상태에서,
현재 남아 있는 DNI 처리 스크립트들을 순서대로 실행하고
최종 TOTAL 파일명을 사용자가 선택해 저장할 수 있도록 제공한다.
"""

from __future__ import print_function

import os
import sys
import shutil

if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

from dni_isp_meansure_to_measure import process_dni_isp_meansure
from dni_isp_merge_with_school_list import merge_to_new_sheet
from dni_fullload_copy_second_to_center_v2 import main as run_fullload_copy
from dni_fullload_select_final import main as run_fullload_final


_RUN_DIR = os.getcwd()
_DNI_DIR = os.path.join(_BASE_DIR, "DNI")


def _log(msg):
    print(msg, flush=True)


def _input(prompt):
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        _log("\n[중단] 사용자 입력으로 종료합니다.")
        sys.exit(1)


def _run_task(task_name, fn):
    _log(f"\n[실행] {task_name}")
    try:
        result = fn()
        if result is False:
            _log(f"[오류] 실패: {task_name}")
            return False
        _log(f"[완료] {task_name}")
        return True
    except Exception as e:
        _log(f"[오류] {task_name}: {e}")
        return False


def _pick_file(prompt, candidates):
    existing = [p for p in candidates if os.path.isfile(p)]
    _log(f"\n{prompt}")
    if existing:
        for i, p in enumerate(existing, 1):
            _log(f"  {i}. {p}")
        _log("  0. 직접 경로 입력")
        s = _input("번호 선택 (Enter: 1번): ")
        if not s:
            return existing[0]
        if s == "0":
            return _input("전체 경로 입력: ")
        try:
            idx = int(s)
            if 1 <= idx <= len(existing):
                return existing[idx - 1]
        except ValueError:
            pass
        _log("[경고] 잘못된 입력. 1번 사용")
        return existing[0]
    _log("  후보 파일 없음. 직접 경로를 입력하세요.")
    return _input("전체 경로 입력: ")


def _collect_files(search_dirs, exts=(".xlsx", ".xlsm", ".xls"), name_keywords=None):
    items = []
    for d in search_dirs:
        if not d or not os.path.isdir(d):
            continue
        try:
            for f in os.listdir(d):
                p = os.path.join(d, f)
                if not os.path.isfile(p):
                    continue
                if exts and not f.lower().endswith(exts):
                    continue
                if name_keywords:
                    name_l = f.lower()
                    if not any(k in name_l for k in name_keywords):
                        continue
                items.append(p)
        except Exception:
            pass
    # 중복 제거 + 정렬
    uniq = []
    seen = set()
    for p in items:
        np = os.path.normcase(os.path.abspath(p))
        if np in seen:
            continue
        seen.add(np)
        uniq.append(p)
    return sorted(uniq)


def _pick_isp_inputs():
    isp_scan_dirs = [
        _RUN_DIR,
        os.path.join(_RUN_DIR, "DNI"),
        os.path.join(_RUN_DIR, "SOURCE"),
        os.path.join(_RUN_DIR, "SOURCE", "DNI"),
        _DNI_DIR,
    ]
    isp_dynamic = _collect_files(
        isp_scan_dirs,
        exts=(".xlsx", ".xlsm", ".xls"),
        name_keywords=("isp", "무선", "대전", "measure", "meansure"),
    )
    isp_meansure = _pick_file(
        "ISP 원본 파일(DNI_ISP_MEANSURE.XLSX)을 선택하세요.",
        isp_dynamic + [
            os.path.join(_RUN_DIR, "DNI_ISP_MEANSURE.XLSX"),
            os.path.join(_RUN_DIR, "DNI", "DNI_ISP_MEANSURE.XLSX"),
            os.path.join(_DNI_DIR, "DNI_ISP_MEANSURE.XLSX"),
            os.path.join(_RUN_DIR, "DNI_ISP_MEANSURE.xlsx"),
            os.path.join(_RUN_DIR, "DNI", "DNI_ISP_MEANSURE.xlsx"),
        ],
    )
    csv_scan_dirs = [
        _RUN_DIR,
        os.path.join(_RUN_DIR, "split"),
        os.path.join(_BASE_DIR, "split"),
        os.path.join(os.path.dirname(_BASE_DIR), "split"),
    ]
    csv_dynamic = _collect_files(csv_scan_dirs, exts=(".csv",), name_keywords=("school_reg_list_dni",))
    school_csv = _pick_file(
        "학교 리스트 CSV(school_reg_list_DNI.csv)를 선택하세요.",
        csv_dynamic + [
            os.path.join(_RUN_DIR, "school_reg_list_DNI.csv"),
            os.path.join(_RUN_DIR, "split", "school_reg_list_DNI.csv"),
            os.path.join(_BASE_DIR, "split", "school_reg_list_DNI.csv"),
            os.path.join(os.path.dirname(_BASE_DIR), "split", "school_reg_list_DNI.csv"),
        ],
    )
    isp_measure = os.path.join(os.path.dirname(isp_meansure), "DNI_ISP_MEASURE.XLSX")
    return isp_meansure, isp_measure, school_csv


def _run_isp_flow():
    isp_meansure, isp_measure, school_csv = _pick_isp_inputs()
    ok1 = _run_task(
        "ISP 변환",
        lambda: process_dni_isp_meansure(input_path=isp_meansure, output_path=isp_measure),
    )
    ok2 = _run_task(
        "ISP 315학교 병합",
        lambda: merge_to_new_sheet(school_list_path=school_csv, isp_measure_path=isp_measure),
    ) if ok1 else False
    return ok1 and ok2


def _pick_fullload_inputs():
    fl_scan_dirs = [
        _RUN_DIR,
        os.path.join(_RUN_DIR, "DNI"),
        os.path.join(_RUN_DIR, "SOURCE"),
        os.path.join(_RUN_DIR, "SOURCE", "DNI"),
        _DNI_DIR,
    ]
    fl_dynamic = _collect_files(
        fl_scan_dirs,
        exts=(".xlsx", ".xlsm", ".xls"),
        name_keywords=("full", "fullload", "전부하", "dno_full", "dni_full"),
    )
    copy_input = _pick_file(
        "전부하 원본(2차복사용) 파일을 선택하세요.",
        fl_dynamic + [
            os.path.join(_RUN_DIR, "DNO_FULLLOAD_MEANSURE_수정.xlsx"),
            os.path.join(_RUN_DIR, "DNI", "DNO_FULLLOAD_MEANSURE_수정.xlsx"),
            os.path.join(_DNI_DIR, "DNO_FULLLOAD_MEANSURE_수정.xlsx"),
            os.path.join(_RUN_DIR, "DNO_FULLLOAD_MEASURE.xlsx"),
            os.path.join(_RUN_DIR, "DNI", "DNO_FULLLOAD_MEASURE.xlsx"),
        ],
    )
    csv_scan_dirs = [
        _RUN_DIR,
        os.path.join(_RUN_DIR, "split"),
        os.path.join(_BASE_DIR, "split"),
        os.path.join(os.path.dirname(_BASE_DIR), "split"),
    ]
    csv_dynamic = _collect_files(csv_scan_dirs, exts=(".csv",), name_keywords=("school_reg_list_dni",))
    school_csv = _pick_file(
        "학교 리스트 CSV(school_reg_list_DNI.csv)를 선택하세요.",
        csv_dynamic + [
            os.path.join(_RUN_DIR, "school_reg_list_DNI.csv"),
            os.path.join(_RUN_DIR, "split", "school_reg_list_DNI.csv"),
            os.path.join(_BASE_DIR, "split", "school_reg_list_DNI.csv"),
            os.path.join(os.path.dirname(_BASE_DIR), "split", "school_reg_list_DNI.csv"),
        ],
    )
    final_input = _pick_file(
        "전부하 최종선정 대상 파일(DNI_FULLLOAD_MEASURE.xlsx)을 선택하세요.",
        [
            os.path.join(_RUN_DIR, "DNI_FULLLOAD_MEASURE.xlsx"),
            os.path.join(_RUN_DIR, "DNI", "DNI_FULLLOAD_MEASURE.xlsx"),
            os.path.join(_DNI_DIR, "DNI_FULLLOAD_MEASURE.xlsx"),
            copy_input,
        ],
    )
    return copy_input, school_csv, final_input


def _run_fullload_flow():
    copy_input, school_csv, final_input = _pick_fullload_inputs()
    ok1 = _run_task(
        "전부하 2차 복사",
        lambda: run_fullload_copy(input_path=copy_input, school_list_csv=school_csv),
    )
    ok2 = _run_task(
        "전부하 최종 선정",
        lambda: run_fullload_final(input_path=final_input),
    ) if ok1 else False
    return ok1 and ok2


def _pick_source_total():
    candidates = [
        os.path.join(_RUN_DIR, "DNI_TOTAL_MEASURE_LIST_V1.xlsx"),
        os.path.join(_RUN_DIR, "TOTAL_MEASURE_LIST_V1.xlsx"),
        os.path.join(_DNI_DIR, "DNI_TOTAL_MEASURE_LIST_V1.xlsx"),
        os.path.join(_DNI_DIR, "TOTAL_MEASURE_LIST_V1.xlsx"),
    ]
    existing = [p for p in candidates if os.path.isfile(p)]
    if existing:
        _log("\n원본 TOTAL 파일을 선택하세요.")
        for i, p in enumerate(existing, 1):
            _log(f"  {i}. {p}")
        _log("  0. 직접 경로 입력")
        s = _input("번호 선택 (Enter: 1번): ")
        if not s:
            return existing[0]
        if s == "0":
            return _input("원본 TOTAL 파일 전체 경로 입력: ")
        try:
            idx = int(s)
            if 1 <= idx <= len(existing):
                return existing[idx - 1]
        except ValueError:
            pass
        _log("[경고] 잘못된 입력. 1번 사용")
        return existing[0]
    return _input("\n원본 TOTAL 파일 전체 경로 입력: ")


def _pick_target_total():
    _log("\n출력 TOTAL 파일명을 선택하세요.")
    _log("  1. DNI_TOTAL_MEASURE_LIST_V1.xlsx")
    _log("  2. TOTAL_MEASURE_LIST_V1.xlsx")
    _log("  0. 직접 파일명/경로 입력")
    s = _input("번호 선택 (Enter: 1번): ")
    if not s or s == "1":
        return os.path.join(_RUN_DIR, "DNI_TOTAL_MEASURE_LIST_V1.xlsx")
    if s == "2":
        return os.path.join(_RUN_DIR, "TOTAL_MEASURE_LIST_V1.xlsx")
    if s == "0":
        p = _input("출력 파일명 또는 전체 경로 입력: ")
        if os.path.isabs(p):
            return p
        return os.path.join(_RUN_DIR, p)
    _log("[경고] 잘못된 입력. 1번 사용")
    return os.path.join(_RUN_DIR, "DNI_TOTAL_MEASURE_LIST_V1.xlsx")


def _sync_total_filename():
    src = _pick_source_total()
    if not src or not os.path.isfile(src):
        _log(f"[오류] 원본 TOTAL 파일 없음: {src}")
        return False
    dst = _pick_target_total()
    if not dst.lower().endswith(".xlsx"):
        dst += ".xlsx"
    try:
        os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
        shutil.copy2(src, dst)
        _log(f"[완료] TOTAL 저장: {dst}")
        return True
    except Exception as e:
        _log(f"[오류] TOTAL 저장 실패: {e}")
        return False


def main():
    _log("=" * 58)
    _log(" DNI 측정 자료 빌더 (TOTAL 파일명 선택 지원)")
    _log("=" * 58)
    _log(f"실행 폴더: {_RUN_DIR}")

    while True:
        _log("\n작업을 선택하세요.")
        _log("  1. ISP 처리 실행 (meansure -> measure, 315학교 병합)")
        _log("  2. 전부하 처리 실행 (2차복사, 최종선정)")
        _log("  3. TOTAL 파일명 선택 저장")
        _log("  4. 전체 실행 (1 -> 2 -> 3)")
        _log("  0. 종료")
        s = _input("선택: ")

        if s == "1":
            ok = _run_isp_flow()
            _log("[결과] ISP 처리 완료" if ok else "[결과] ISP 처리 중 오류")
        elif s == "2":
            ok = _run_fullload_flow()
            _log("[결과] 전부하 처리 완료" if ok else "[결과] 전부하 처리 중 오류")
        elif s == "3":
            _sync_total_filename()
        elif s == "4":
            ok = True
            ok = _run_isp_flow() and ok
            ok = _run_fullload_flow() and ok
            ok = _sync_total_filename() and ok
            _log("[결과] 전체 실행 완료" if ok else "[결과] 전체 실행 중 일부 실패")
        elif s == "0":
            _log("종료합니다.")
            break
        else:
            _log("잘못된 선택입니다.")


if __name__ == "__main__":
    main()

