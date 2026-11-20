# -*- mode: python ; coding: utf-8 -*-
# Shine - recette de construction PyInstaller (utilisee par build_exe.bat : `pyinstaller config.spec`).
# Resultat : dist\Shine.exe, fenetre sans console, un seul fichier.
# config.yaml est embarque comme configuration par defaut ; un config.yaml pose a cote de Shine.exe
# la remplace (fusion cle par cle) sans reconstruire l'executable.
import os

ROOT = SPECPATH  # dossier de ce fichier
ICON = os.path.join(ROOT, 'assets', 'shine.ico')  # icone de Shine.exe (ignoree si le fichier est absent)

a = Analysis(
    [os.path.join(ROOT, 'shine.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[(os.path.join(ROOT, 'config.yaml'), '.')],
    hiddenimports=[],
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
    name='Shine',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=ICON if os.path.exists(ICON) else None,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
