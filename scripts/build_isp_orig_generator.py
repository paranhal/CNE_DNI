# -*- coding: utf-8 -*-
"""
ISP 원데이터 생성기 실행 파일 빌드 스크립트

실행: python scripts/build_isp_orig_generator.py
또는: pyinstaller scripts/isp_orig_data_generator.spec

주의: 빌드 후에도 src/measure/ 내 소스 파일은 삭제하지 말 것.
"""

from __future__ import annotations

import os
import subprocess
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_MEASURE = os.path.join(PROJECT_ROOT, "src", "measure")


def main():
    os.chdir(PROJECT_ROOT)
    # PyInstaller onefile로 단일 실행 파일 생성
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", "ISP_원데이터_생성기",
        "--add-data", f"{SRC_MEASURE}{os.pathsep}measure",  # measure_utils 등
        "--hidden-import", "openpyxl",
        "--hidden-import", "measure_utils",
        "--console",
        os.path.join(SRC_MEASURE, "isp_orig_data_generator.py"),
    ]
    print("실행:", " ".join(cmd))
    subprocess.check_call(cmd)
    print("완료: dist/ISP_원데이터_생성기 (또는 .exe)")


if __name__ == "__main__":
    main()
