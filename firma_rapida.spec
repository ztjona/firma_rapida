# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Firma Rapida
# Build with: pyinstaller firma_rapida.spec

from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# Collect all submodules and data for heavy packages
datas_pyhanko, binaries_pyhanko, hiddenimports_pyhanko = collect_all('pyhanko')
datas_certval, binaries_certval, hiddenimports_certval = collect_all('pyhanko_certvalidator')
datas_asn1, binaries_asn1, hiddenimports_asn1 = collect_all('asn1crypto')
datas_fitz, binaries_fitz, hiddenimports_fitz = collect_all('fitz')

all_datas = (
    datas_pyhanko
    + datas_certval
    + datas_asn1
    + datas_fitz
)
all_binaries = (
    binaries_pyhanko
    + binaries_certval
    + binaries_asn1
    + binaries_fitz
)
all_hidden = (
    hiddenimports_pyhanko
    + hiddenimports_certval
    + hiddenimports_asn1
    + hiddenimports_fitz
    + collect_submodules('cryptography')
    + collect_submodules('PIL')
    + [
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        'email.mime.multipart',
        'email.mime.text',
        'certifi',
        'qrcode',
        'qrcode.image.pure',
        'qrcode.image.base',
    ]
)

a = Analysis(
    ['firma_rapida.py'],
    pathex=[],
    binaries=all_binaries,
    datas=all_datas,
    hiddenimports=all_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy', 'IPython', 'jupyter'],
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
    name='FirmaRapida',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No terminal window (GUI app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',      # Uncomment and add icon.ico to use a custom icon
)
