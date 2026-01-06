# -*- mode: python ; coding: utf-8 -*-
"""
PDF Locker PyInstaller spec file

tkinterdnd2を正しくバンドルするための設定を含みます。
"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# tkinterdnd2のデータファイルを収集
try:
    tkdnd_datas = collect_data_files('tkinterdnd2')
except Exception:
    tkdnd_datas = []
    print("Warning: tkinterdnd2 not found, drag & drop will be disabled")

# tkinterdnd2のサブモジュールを収集
try:
    tkdnd_imports = collect_submodules('tkinterdnd2')
except Exception:
    tkdnd_imports = []

block_cipher = None

a = Analysis(
    ['pdf_locker.py'],
    pathex=[],
    binaries=[],
    datas=tkdnd_datas,
    hiddenimports=[
        'tkinterdnd2',
        'pypdf',
        'pypdf.generic',
        'pypdf._crypt_providers',
        'pypdf._crypt_providers._cryptography',
    ] + tkdnd_imports,
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'cv2',
        'torch',
        'tensorflow',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
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
    name='PDF_Locker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUIアプリケーションなのでコンソールを非表示
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# macOS向けの設定
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='PDF_Locker.app',
        icon=None,
        bundle_identifier='com.pdflocker.app',
        info_plist={
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleName': 'PDF Locker',
            'NSHighResolutionCapable': True,
        },
    )
