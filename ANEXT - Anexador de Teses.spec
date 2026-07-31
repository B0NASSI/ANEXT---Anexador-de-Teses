# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [
    ('assets/pdf.ico', 'assets'),
    ('assets/folder interno.ico', 'assets'),
    ('assets/Logo RS completa colorida.png', 'assets'),
    ('modelo', 'modelo'),
]
binaries = []
hiddenimports = ['win32com', 'win32com.client', 'pythoncom']
tmp_ret = collect_all('ttkbootstrap')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['src/app.py'],
    pathex=['src'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy',
        'pywinauto',
        'adodbapi',
        'isapi',
        'pythonwin',
        'setuptools',
        'pip',
        'unittest',
        'email',
        'http',
        'xmlrpc',
        'ftplib',
        'multiprocessing',
    ],
    noarchive=False,
    optimize=0,
)
# binários de formatos/recursos não usados pelo app (avif/webp/cms/imagemath do
# Pillow, lxml.html.diff, isoschematron, pythonwin) — cortados manualmente pois
# os hooks de terceiros os incluem sempre, mesmo sem uso no código (o app só
# abre .ico/.png e PNGs gerados internamente pelo fitz).
_prefixos_nao_usados = (
    'pil\\_avif', 'pil/_avif',
    'pil\\_webp', 'pil/_webp',
    'pil\\_imagingcms', 'pil/_imagingcms',
    'pil\\_imagingmath', 'pil/_imagingmath',
    'lxml\\html', 'lxml/html',
    'lxml\\isoschematron', 'lxml/isoschematron',
    'pythonwin\\', 'pythonwin/',
    'win32\\win32trace', 'win32/win32trace',
)
a.binaries = [x for x in a.binaries if not x[0].lower().startswith(_prefixos_nao_usados)]
a.datas = [x for x in a.datas if not x[0].lower().startswith(_prefixos_nao_usados)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ANEXT - Anexador de Teses',
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
    icon=['assets/pdf.ico'],
)
