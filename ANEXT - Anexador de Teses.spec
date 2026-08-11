# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

assets_datas = [
    ('assets/pdf.ico', 'assets'),
    ('assets/folder interno.ico', 'assets'),
    ('assets/Logo RS completa colorida.png', 'assets'),
    ('assets/icons8-informações-50.png', 'assets'),
]

# ---------------------------------------------------------------------------
# ANEXT.exe - programa principal
# ---------------------------------------------------------------------------
app_datas = list(assets_datas) + [('modelo', 'modelo'), ('NOTAS DE ATUALIZAÇÃO', 'NOTAS DE ATUALIZAÇÃO')]
app_binaries = []
app_hiddenimports = ['win32com', 'win32com.client', 'pythoncom']
# 'requests' entrou aqui porque app.py importa launcher.py (pra reaproveitar
# get_latest_release/is_newer/read_local_version na checagem de versão fora
# do launcher) - precisa do collect_all igual ao do launcher, senão fica só
# com o hidden import "requests" mas sem os módulos que ele carrega por
# baixo dos panos (ver exclusão de 'email'/'http' abaixo, removida por isso)
for pacote in ('ttkbootstrap', 'PIL', 'fitz', 'requests'):
    tmp_ret = collect_all(pacote)
    app_datas += tmp_ret[0]; app_binaries += tmp_ret[1]; app_hiddenimports += tmp_ret[2]

a_app = Analysis(
    ['src/app.py'],
    pathex=['src'],
    binaries=app_binaries,
    datas=app_datas,
    hiddenimports=app_hiddenimports,
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
        # 'email' e 'http' NÃO podem ser excluídos aqui: requests (usado na
        # checagem de versão fora do launcher) depende deles por baixo dos
        # panos (via urllib3) - excluir quebra a checagem em silêncio, só
        # detectável com log no except (já aconteceu no REQUERID)
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
a_app.binaries = [x for x in a_app.binaries if not x[0].lower().startswith(_prefixos_nao_usados)]
a_app.datas = [x for x in a_app.datas if not x[0].lower().startswith(_prefixos_nao_usados)]

pyz_app = PYZ(a_app.pure)

exe_app = EXE(
    pyz_app,
    a_app.scripts,
    [],
    exclude_binaries=True,
    name='ANEXT',
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
    icon=['assets/pdf.ico'],
)

# onedir: pasta com o .exe + dependências soltas ao lado (_internal), em vez
# de um único .exe que se autoextrai pra uma pasta temporária a cada execução
# — isso é o que tornava a abertura lenta (~30s, reextraindo tudo sempre).
coll_app = COLLECT(
    exe_app,
    a_app.binaries,
    a_app.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ANEXT',
)

# ---------------------------------------------------------------------------
# ANEXT Launcher.exe - checa atualizações no GitHub antes de abrir o ANEXT.exe
# ---------------------------------------------------------------------------
launcher_datas = list(assets_datas)
launcher_binaries = []
launcher_hiddenimports = []
for pacote in ('ttkbootstrap', 'PIL', 'requests'):
    tmp_ret = collect_all(pacote)
    launcher_datas += tmp_ret[0]; launcher_binaries += tmp_ret[1]; launcher_hiddenimports += tmp_ret[2]

a_launcher = Analysis(
    ['src/launcher.py'],
    pathex=['src'],
    binaries=launcher_binaries,
    datas=launcher_datas,
    hiddenimports=launcher_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
a_launcher.binaries = [x for x in a_launcher.binaries if not x[0].lower().startswith(_prefixos_nao_usados)]
a_launcher.datas = [x for x in a_launcher.datas if not x[0].lower().startswith(_prefixos_nao_usados)]

pyz_launcher = PYZ(a_launcher.pure)

# contents_directory: nome de pasta próprio (_internal_launcher) pra não
# colidir com o _internal do ANEXT.exe quando os dois ficam lado a lado na
# mesma pasta de instalação
exe_launcher = EXE(
    pyz_launcher,
    a_launcher.scripts,
    [],
    exclude_binaries=True,
    name='ANEXT Launcher',
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
    icon=['assets/pdf.ico'],
    contents_directory='_internal_launcher',
)

coll_launcher = COLLECT(
    exe_launcher,
    a_launcher.binaries,
    a_launcher.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ANEXT Launcher',
)
