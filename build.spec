# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 规格（Python 3.8, Windows 64 位, 目标 Win7 SP1+）
# 构建： py -3.8 -m PyInstaller --noconfirm build.spec
from PyInstaller.utils.hooks import collect_data_files
import os

block_cipher = None


def _dnd_datas():
    """tkinterdnd2 自带 linux/macOS/Windows 全平台 tkdnd 二进制；
    本程序只在 Windows 用，只保留 win-x64，省掉几 MB 死重量。"""
    out = []
    for src, dst in collect_data_files('tkinterdnd2'):
        if (os.sep + 'tkdnd' + os.sep) in src and 'win-x64' not in src:
            continue
        out.append((src, dst))
    return out


a = Analysis(
    ['gwps_main.py'],
    pathex=['.'],
    binaries=[],
    # tkinterdnd2 没有官方 hook：它的 tkdnd 二进制必须手动带上，否则拖拽在 exe 里失效
    datas=[('fonts', 'fonts'), ('assets', 'assets')] + _dnd_datas(),
    hiddenimports=['win32timezone', 'docx', 'lxml', 'lxml._elementpath',
                   'win32api', 'win32com', 'win32com.client', 'pythoncom', 'pywintypes',
                   'tkinterdnd2'],
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy'],
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
    name='公文一键排版助手',
    debug=False,
    strip=False,
    upx=False,
    console=False,      # 窗口程序
    icon='assets/app.ico',
)
