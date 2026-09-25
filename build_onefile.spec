# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 单文件打包配置（适用于 PyInstaller 6.x）

用法：
    pyinstaller build_onefile.spec --noconfirm
产物：
    dist/DocFormatTool.exe   （单文件，无需 Python 环境）
"""
import os
import sys

sys.setrecursionlimit(5000)

# 以 spec 文件所在目录为项目根目录，避免硬编码绝对路径
project_dir = os.path.abspath(SPECPATH)

a = Analysis(
    ['main.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[
        # 目录树节点图标，运行时通过 sys._MEIPASS 读取
        (os.path.join(project_dir, 'folder.png'), '.'),
        # 程序图标：既作为 exe 的图标资源，也供运行时设置窗口/任务栏图标
        (os.path.join(project_dir, 'app.ico'), '.'),
    ],
    hiddenimports=[
        # win32com 为运行时动态导入，需显式声明
        'pythoncom',
        'pywintypes',
        'win32com',
        'win32com.client',
        'win32com.client.dynamic',
        'win32com.client.gencache',
        'win32timezone',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # pythonwin 的 MFC 界面库：由 win32com.client.gencache -> makepy 的
        # GUIProgress.__init__ 惰性导入引入，本程序不会走到该分支，
        # 且其中 win32ui.pyd 依赖缺失的 mfc140u.dll，属于无效负载。
        'pythonwin',
        'win32ui',
        'win32uiole',
        'pywin.mfc',
        'pywin.dialogs',
        'pywin.tools',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DocFormatTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # 由 make_icon.py 生成，内嵌 16/24/32/48/64/128/256 七种尺寸
    icon=os.path.join(project_dir, 'app.ico'),
    # 版本资源：右键"属性-详细信息"可看到"文档格式设置程序060930"
    version=os.path.join(project_dir, 'version_info.txt'),
)
