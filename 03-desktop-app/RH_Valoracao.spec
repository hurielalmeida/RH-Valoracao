# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

pymupdf_datas, pymupdf_binaries, pymupdf_hiddenimports = collect_all('pymupdf')

a = Analysis(
    ['desktop_launcher.py'],
    pathex=['.'],
    binaries=pymupdf_binaries,
    datas=[('frontend', 'frontend'), ('core', 'core'), ('backend', 'backend')] + pymupdf_datas,
    hiddenimports=['backend.app', 'core.rh_valoracao_core'] + pymupdf_hiddenimports,
    noarchive=False,
)

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RH_Valoracao',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
