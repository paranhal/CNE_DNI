# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['d:\\CNE_DNI\\src\\measure\\dni_measure_builder.py'],
    pathex=['d:\\CNE_DNI\\src\\measure'],
    binaries=[],
    datas=[],
    hiddenimports=['dni_isp_meansure_to_measure', 'dni_isp_merge_with_school_list', 'dni_fullload_copy_second_to_center_v2', 'dni_fullload_select_final'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='dni_measure_builder_new2',
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
