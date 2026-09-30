# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 文件夹（onedir）打包配置（适用于 PyInstaller 6.x）

与 build_onefile.spec 的唯一区别：产物是「一个文件夹」而不是「一个 exe」。

为什么要留这个版本
------------------
单文件 exe 每次启动都要把内嵌的 Python 运行时解压到 %TEMP%\\_MEIxxxxxx，
这一步在某些机器上会失败并弹出
    Could not create temporary directory!
常见诱因是安全软件拦截 _MEI 目录创建，或 %TEMP% 被优化工具改成了
不可写的路径。文件夹版启动时**完全不解压、不碰 %TEMP%**，从根上绕开
这个错误；代价是必须把整个文件夹一起分发，不能只发 exe。

分发说明（发给最终用户时务必写清楚）：
    整个 DocFormatTool 文件夹一起拷贝，双击里面的 DocFormatTool.exe。
    只拷 exe 是跑不起来的。

用法：
    pyinstaller build_onedir.spec --noconfirm
产物：
    dist/DocFormatTool/DocFormatTool.exe   （及同目录依赖，无需 Python 环境）
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
        # （onedir 模式下 _MEIPASS 就是 exe 所在目录）
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

# onedir 模式：EXE 只放引导程序与脚本，依赖全部交给 COLLECT 平铺到文件夹
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DocFormatTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # 由 make_icon.py 生成，内嵌 16/24/32/48/64/128/256 七种尺寸
    icon=os.path.join(project_dir, 'app.ico'),
    # 版本资源：右键"属性-详细信息"可看到"文档格式设置程序060929"
    version=os.path.join(project_dir, 'version_info.txt'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='DocFormatTool',
)
