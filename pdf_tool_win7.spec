# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：面向 Windows 7 的单文件 exe。

必须用 Python 3.8 + PyInstaller 5.x 构建（见 build-windows-win7.yml），
语法刻意保持 PyInstaller 5.x 兼容。
"""

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
    runtime_hooks=[],
    excludes=[],
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
    upx=False,
    runtime_tmpdir=None,
    console=False,          # 纯 GUI，不带黑窗口
    icon='assets/app.ico',
)
