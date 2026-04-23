# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Studio Flow

import sys
from pathlib import Path

ROOT = Path(SPECPATH)

a = Analysis(
    [str(ROOT / 'tracker.py')],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        # Bundle ffmpeg next to the executable
        (str(ROOT / 'vendor' / 'ffmpeg.exe'), '.'),
        # Bundle the icon
        (str(ROOT / 'assets' / 'icon.ico'), 'assets'),
    ],
    hiddenimports=[
        # pycaw / COM
        'pycaw',
        'pycaw.pycaw',
        'comtypes',
        'comtypes.client',
        'comtypes.server',
        # pystray Windows backend
        'pystray._win32',
        'pystray.backend.win32',
        # Pillow
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        # psutil
        'psutil',
        # tkinter (usually included but be explicit)
        'tkinter',
        'tkinter.filedialog',
        'tkinter.ttk',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Keep bundle small — these aren't used
        'matplotlib', 'numpy', 'scipy', 'pandas',
        'IPython', 'notebook', 'sphinx',
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='StudioFlow',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # no console window — it's a tray app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / 'assets' / 'icon.ico'),
    version_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[
        'ffmpeg.exe',     # already compressed; UPX on it causes issues
        'vcruntime*.dll',
        'api-ms-*.dll',
    ],
    name='StudioFlow',
)
