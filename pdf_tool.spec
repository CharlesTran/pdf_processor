# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：产出单文件无控制台窗口的 Windows exe。

用法（Windows 上）:
    pip install -r requirements.txt pyinstaller
    pyinstaller --noconfirm --clean pdf_tool.spec
"""

import os

block_cipher = None

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # doc2pdf（Word COM）在函数内延迟导入，需显式收集
        'win32com', 'win32com.client', 'win32com.client.gencache',
        'pythoncom', 'pywintypes', 'win32timezone',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'unittest'],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF处理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 纯 GUI，不带黑窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/app.ico' if os.path.exists('assets/app.ico') else None,
)
