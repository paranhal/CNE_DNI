# -*- coding: utf-8 -*-
"""
BSSID 매칭: CSV 폴더에서 memo1(장비 코드)과 BSSID 열을 읽어
충남_BSSID.csv, 대전_BSSID.csv로 저장하는 실행 파일.
"""

from __future__ import annotations

import csv
import os
import sys

# 기본 CSV 폴더 (Google Drive 어비스)
DEFAULT_CSV_DIR = os.path.expanduser(
    "/Users/paranhal/Library/CloudStorage/GoogleDrive-paranhanl66@gmail.com/내 드라이브/어비스"
)
OUTPUT_충남 = "충남_BSSID.csv"
OUTPUT_대전 = "대전_BSSID.csv"


def _norm_key(s: str) -> str:
    """CSV 헤더 비교용 정규화."""
    if not s:
        return ""
    return str(s).strip().lower().replace(" ", "").replace("_", "").replace(".", "")


def _find_col(fieldnames: list[str], *candidates: str) -> int | None:
    """필드명 리스트에서 후보 중 하나와 매칭되는 열 인덱스. 없으면 None."""
    norm_candidates = [_norm_key(c) for c in candidates]
    for i, h in enumerate(fieldnames):
        n = _norm_key(h)
        if n in norm_candidates:
            return i
    return None


def _collect_bssid_by_region(csv_dir: str) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    """
    csv_dir 내 모든 CSV에서 memo1(또는 momo1)과 BSSID 열을 읽어
    (memo1, BSSID) 쌍을 수집. 파일명에 '충남' 포함 -> 충남, '대전' 포함 -> 대전.
    반환: (충남 리스트, 대전 리스트). 중복 제거된 (memo1, bssid) 튜플 리스트.
    """
    충남_pairs: set[tuple[str, str]] = set()
    대전_pairs: set[tuple[str, str]] = set()

    if not os.path.isdir(csv_dir):
        return [], []

    for fname in sorted(os.listdir(csv_dir)):
        if not fname.lower().endswith(".csv") or fname.startswith(".") or fname.startswith("~$"):
            continue
        path = os.path.join(csv_dir, fname)
        if not os.path.isfile(path):
            continue
        is_충남 = "충남" in fname
        is_대전 = "대전" in fname
        if not is_충남 and not is_대전:
            continue

        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if not header:
                    continue
                memo_col = _find_col(header, "memo1", "momo1", "메모1", "장비관리번호")
                bssid_col = _find_col(header, "BSSID", "bssid")
                if memo_col is None or bssid_col is None:
                    continue
                for row in reader:
                    if max(memo_col, bssid_col) >= len(row):
                        continue
                    memo = (row[memo_col] or "").strip()
                    bssid = (row[bssid_col] or "").strip()
                    if not memo or not bssid:
                        continue
                    pair = (memo, bssid)
                    if is_충남:
                        충남_pairs.add(pair)
                    if is_대전:
                        대전_pairs.add(pair)
        except Exception:
            continue

    return (sorted(충남_pairs), sorted(대전_pairs))


def main() -> None:
    csv_dir = DEFAULT_CSV_DIR
    if len(sys.argv) > 1 and sys.argv[1].strip():
        csv_dir = os.path.abspath(sys.argv[1].strip())

    if not os.path.isdir(csv_dir):
        print(f"[오류] CSV 폴더가 없습니다: {csv_dir}")
        sys.exit(1)

    print(f"[BSSID 매칭] CSV 폴더: {csv_dir}")
    충남_list, 대전_list = _collect_bssid_by_region(csv_dir)
    print(f"  충남: {len(충남_list)}개 (memo1–BSSID 쌍)")
    print(f"  대전: {len(대전_list)}개 (memo1–BSSID 쌍)")

    # 저장: CSV 폴더에 쓰기 시도, 실패 시 현재 작업 디렉터리
    out_dir = csv_dir
    try:
        test_path = os.path.join(out_dir, ".write_test")
        with open(test_path, "w") as _:
            pass
        os.remove(test_path)
    except (OSError, PermissionError):
        out_dir = os.getcwd()
        print(f"  (CSV 폴더에 쓰기 불가 → 출력을 현재 폴더로 저장: {out_dir})")

    충남_path = os.path.join(out_dir, OUTPUT_충남)
    대전_path = os.path.join(out_dir, OUTPUT_대전)

    with open(충남_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["memo1", "BSSID"])
        w.writerows(충남_list)
    print(f"[저장] {충남_path}")

    with open(대전_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["memo1", "BSSID"])
        w.writerows(대전_list)
    print(f"[저장] {대전_path}")

    print("[완료]")


if __name__ == "__main__":
    main()
