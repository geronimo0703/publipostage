# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

a = Analysis(
    ['app.py'],
    pathex=['scripts'],
    binaries=[],
    datas=[
        ('config/config.yaml', 'config'),
        ('VERSION', '.'),
        ('templates', 'templates'),
        ('static', 'static'),
        *collect_data_files('mammoth'),
    ],
    hiddenimports=[
        'email.mime.multipart', 'email.mime.text', 'email.mime.base',
        'publipostage_stock_ed7',
        *collect_submodules('mammoth'),
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='publipostage',
    debug=False,
    console=True,
)
