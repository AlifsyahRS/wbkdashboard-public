# -*- mode: python ; coding: utf-8 -*-

# File used for debugging the executable

import sys

sys.setrecursionlimit(sys.getrecursionlimit() * 5)


block_cipher = None


a = Analysis(
    ['onesheet.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pyodbc'],
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tensorflow', 'torch', 'hook', 'hooks', 'PyQt5', 'PyQt', 'notebook', 'pyarrow', 'statsmodels', 'nbconvert', 'babel', 'botocore', 'jedi', 'llvmlite', 'sphinx', 'gevent', 'lxml', 'tcl8', 'Cython', 'brotli', 'cryptography', 'nbformat', 'tables', 'numba', 'zmq', 'numexpr', 'certifi', 'docutils', 'IPython', 'nacl', 'bcrypt', 'sqlalchemy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='wbkdashboard-debug',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    name='wbkdashboard-debug',
)
