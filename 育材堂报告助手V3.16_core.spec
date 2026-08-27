# -*- mode: python ; coding: utf-8 -*-

# The core application deliberately has no PyInstaller Splash.  The native
# launcher owns the startup experience, avoiding the blank Tcl/Tk window that
# PyInstaller's splash can display on some Windows installations.
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('拉伸模板.pptx', '.'),
        ('VDA弯曲角模板.pptx', '.'),
    ],
    hiddenimports=[
        'tkinterdnd2',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'pandas',
        'openpyxl',
        'pptx',
        'pdfplumber',
        'PIL',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch', 'torchvision', 'tensorflow', 'scipy', 'numpy.core._dotblas',
        'matplotlib', 'IPython', 'jupyter', 'notebook', 'pytest', 'sphinx',
        'docutils', 'jedi', 'parso', 'pyarrow', 'numba', 'llvmlite', 'boto3',
        'botocore', 's3fs', 'fsspec', 'tables', 'numexpr',
    ],
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
    name='育材堂报告助手V3.16_core',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
