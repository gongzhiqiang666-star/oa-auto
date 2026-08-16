# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置：生成 Windows 单文件程序 交易汇总工具.exe
# 用法：pyinstaller --noconfirm --clean 交易汇总工具.spec

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('交易汇总工具.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='交易汇总工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,            # 不弹出黑色命令行窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['交易汇总工具.ico'],
)
