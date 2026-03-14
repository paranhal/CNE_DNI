# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec: ISP 원데이터 생성기
# 빌드: cd 프로젝트루트 && pyinstaller scripts/isp_orig_data_generator.spec

import os
# 프로젝트 루트 = 현재 작업 디렉터리 (반드시 cd CNE_DNI 후 pyinstaller 실행)
PROJECT_ROOT = os.getcwd()
SRC_MEASURE = os.path.join(PROJECT_ROOT, 'src', 'measure')
_MAIN_SCRIPT = os.path.join(SRC_MEASURE, 'isp_orig_data_generator.py')
if not os.path.isfile(_MAIN_SCRIPT):
    raise SystemExit('스크립트 없음: %r (프로젝트 루트 CNE_DNI에서 실행하세요)' % _MAIN_SCRIPT)
a = Analysis(
    [_MAIN_SCRIPT],
    pathex=[SRC_MEASURE, PROJECT_ROOT],
    binaries=[],
    datas=[(SRC_MEASURE, 'measure')],
    hiddenimports=['openpyxl', 'measure_utils', 'et_xmlfile'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ISP_원데이터_생성기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
