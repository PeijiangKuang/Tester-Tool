# -*- mode: python ; coding: utf-8 -*-

import platform
import os

block_cipher = None

# 根据平台选择入口
if platform.system() == "Windows":
    entry = ["src\\tester\\__main__.py"]
else:
    entry = ["src/tester/__main__.py"]

a = Analysis(
    entry,
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['PyQt6', 'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets', 'openpyxl', 'et_xmlfile'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ===== Windows EXE =====
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TesterTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TesterTool',
)

# ===== macOS App Bundle =====
app = BUNDLE(
    coll,
    name='TesterTool.app',
    icon=None,
    bundle_identifier='com.testtool.app',
    info_plist={
        'CFBundleName': 'TesterTool',
        'CFBundleDisplayName': '试验数据处理工具',
        'CFBundleIdentifier': 'com.testtool.app',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundlePackageType': 'APPL',
        'CFBundleExecutable': 'TesterTool',
        'LSMinimumSystemVersion': '10.14',
    },
)
