# PyInstaller spec for Certificate Auto desktop executable.
# Build: pyinstaller certificate_auto.spec

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Data files: templates and static (images, fonts if present) so they are inside the bundle
# SPECPATH = folder that contains this .spec file (e.g. certificate_auto_desktop)
import os
project_root = SPECPATH
templates = (os.path.join(project_root, 'templates'), 'templates')
static = (os.path.join(project_root, 'static'), 'static')
datas = [templates, static]

# Hidden imports for jinja2, pdfkit, openpyxl (Excel), tkinter (built-in)
hiddenimports = [
    'jinja2',
    'jinja2.loaders',
    'pdfkit',
    'dotenv',
    'openpyxl',
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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

# Exe icon: use .ico from static/images (e.g. logo-nobg.ico)
# Absolute path so PyInstaller finds it regardless of cwd
icon_path = os.path.abspath(os.path.join(project_root, 'static', 'images', 'logo-nobg.ico'))
if not os.path.isfile(icon_path):
    icon_path = None

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CertificateAuto',
    debug=False,
    icon=icon_path,
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
    name='CertificateAuto',
)
