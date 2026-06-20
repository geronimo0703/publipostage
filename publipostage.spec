# -*- mode: python ; coding: utf-8 -*-
a = Analysis(
    ['app.py'],
    pathex=['scripts'],
    binaries=[],
    datas=[
        ('config/config.yaml', 'config'),
        ('templates', 'templates'),
    ],
    hiddenimports=[
        'email.mime.multipart', 'email.mime.text', 'email.mime.base',
        'publipostage_stock_ed7',
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
    console=True,  # mets False si tu veux masquer la console (Windows)
)
