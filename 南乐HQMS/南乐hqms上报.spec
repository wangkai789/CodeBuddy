# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['南乐hqms上报_web.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('模板', '模板'),
        ('logs', 'logs'),
        ('输出', '输出'),
    ],
    hiddenimports=[
        'flask', 'pyodbc', 'pandas', 'werkzeug', 'jinja2', 'markupsafe',
        'itsdangerous', 'click', 'blinker', 'dateutil', 'pytz'
    ],
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='南乐HQMS上报系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
