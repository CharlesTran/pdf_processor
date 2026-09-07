# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller macOS 打包配置：产出 PDF处理工具.app（无控制台窗口）。

用法（macOS 上，推荐配合 build_dmg.sh）:
    pip install -r requirements.txt pyinstaller
    pyinstaller --noconfirm --clean pdf_tool_mac.spec
产物: dist_mac/PDF处理工具.app
"""

import os

block_cipher = None

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PDF处理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 纯 GUI，不带终端窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 图标在 BUNDLE 中设置
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='PDF处理工具',
)

app = BUNDLE(
    coll,
    name='PDF处理工具.app',
    icon='assets/app.icns',
    bundle_identifier='com.charles.pdfprocessor',
    info_plist={
        'CFBundleName': 'PDF 处理工具',
        'CFBundleDisplayName': 'PDF 处理工具',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13',
    },
)
