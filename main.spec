# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

project_path = Path('./')
sys.path.insert(0, str(project_path))
from version import __version__


def _make_version_info(version_str):
    parts = (version_str.split('.') + ['0', '0', '0', '0'])[:4]
    v = tuple(int(x) for x in parts)
    vs = ', '.join(str(x) for x in v)
    vdot = '.'.join(str(x) for x in v)
    return f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({vs}),
    prodvers=({vs}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(u'040904B0', [
        StringStruct(u'CompanyName',      u'EzFacture'),
        StringStruct(u'FileDescription',  u'EzFacture - Logiciel de facturation'),
        StringStruct(u'FileVersion',      u'{vdot}'),
        StringStruct(u'InternalName',     u'ezfacture'),
        StringStruct(u'LegalCopyright',   u''),
        StringStruct(u'OriginalFilename', u'ezfacture-{version_str}.exe'),
        StringStruct(u'ProductName',      u'EzFacture'),
        StringStruct(u'ProductVersion',   u'{vdot}'),
      ])
    ]),
    VarFileInfo([VarStruct(u'Translation', [0x0409, 0x04B0])])
  ]
)
"""

_version_file = str(project_path / 'file_version_info.txt')
with open(_version_file, 'w', encoding='utf-8') as _f:
    _f.write(_make_version_info(__version__))


hidden_imports = collect_submodules('xlwings') + ['tzdata']
facturx_datas = collect_data_files('facturx')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas = [
        (str(Path("images/logo.png")), "images"),
    ] + facturx_datas,
    hiddenimports=[],
    hookspath=[str(project_path / 'hooks')],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=f'ezfacture-{__version__}',
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
    version=_version_file,
)
