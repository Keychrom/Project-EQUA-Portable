# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['pyraxis.py'],
    pathex=[],
    binaries=[],
    # 'datas'にインストールしたいファイルを追加します。
    # ('ソースパス', '実行ファイル内の格納先フォルダ')
    datas=[('app', 'app')],
    hiddenimports=['win32com.client', 'winreg', 'uuid'],
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
    [], # runtime_hooks
    name='EQUA_Installer', # 生成されるインストーラーのファイル名
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,          # GUIアプリなのでFalse
    icon='installer.ico',   # インストーラー自体のアイコン
    uac_admin=True,         # 管理者権限を要求する
)