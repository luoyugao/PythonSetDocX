import tkinter as tk
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import doc_parameters_manager as dpm

# 自检报告的默认输出文件（打包成 --noconsole 的 exe 时没有控制台可供阅读，
# 因此把结果同时写进文件，便于打包后验证产物是否真的可用）。
SELFTEST_REPORT = os.path.join(
    os.environ.get('TEMP', os.path.expanduser('~')),
    'DocFormatTool_selftest.txt')


def _emit(lines, sink):
    text = "\n".join(lines) + "\n"
    sink.write(text)
    # exe 以 --noconsole 打包时 stdout 可能不存在或为 None，逐个尝试
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                stream.write(text)
                stream.flush()
        except Exception:
            pass


def selftest():
    """轻量自检：只做导入与资源检查，不创建窗口、不连接 Word。

    用于验证打包产物在**没有 Python 环境**的机器上能否正常启动：
    打包遗漏的模块/资源会在这里以 FAIL 暴露，而不是等用户点按钮时才报错。

    用法：
        DocFormatTool.exe --selftest
        DocFormatTool.exe --selftest D:\\out\\report.txt
    """
    report_path = SELFTEST_REPORT
    if len(sys.argv) > 2 and sys.argv[2].strip():
        report_path = sys.argv[2].strip()

    lines = []
    failures = []

    lines.append("=== 文档格式设置程序 自检 ===")
    lines.append(f"frozen(打包运行) : {getattr(sys, 'frozen', False)}")
    lines.append(f"executable      : {sys.executable}")
    lines.append(f"_MEIPASS        : {getattr(sys, '_MEIPASS', '(无)')}")
    lines.append(f"python          : {sys.version.split()[0]}")
    lines.append(f"工作目录        : {os.getcwd()}")
    lines.append(f"设置文件        : {dpm.SETTINGS_FILENAME}")

    # ---- 1) 关键第三方依赖 ----
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        lines.append("PASS  pywin32 (pythoncom / win32com.client) 可导入")
    except Exception as ex:
        failures.append(f"pywin32 导入失败: {ex}")
        lines.append(f"FAIL  pywin32 导入失败: {ex}")

    # ---- 2) 程序自身的模块 ----
    for name in ("word_constants", "universal", "doc_parameters_manager",
                 "auto_numbering", "page_number_manager",
                 "set_document_format", "main_form"):
        try:
            __import__(name)
            lines.append(f"PASS  模块可导入: {name}")
        except Exception as ex:
            failures.append(f"模块 {name} 导入失败: {ex}")
            lines.append(f"FAIL  模块 {name} 导入失败: {ex}")

    # ---- 3) 打包进来的资源文件 ----
    try:
        from main_form import resource_path
    except Exception as ex:
        resource_path = None
        failures.append(f"无法导入 resource_path: {ex}")
        lines.append(f"FAIL  无法导入 resource_path: {ex}")

    if resource_path is not None:
        for res in ("app.ico", "folder.png"):
            path = resource_path(res)
            if os.path.exists(path):
                lines.append(f"PASS  资源存在: {res} ({os.path.getsize(path)} 字节)")
            else:
                failures.append(f"资源缺失: {res} -> {path}")
                lines.append(f"FAIL  资源缺失: {res} -> {path}")

    # ---- 4) 界面与窗口标题（Tk 可用性） ----
    try:
        root = tk.Tk()
        root.withdraw()
        title = None
        try:
            from main_form import MainForm
            app = MainForm(root)
            title = root.title()
            root.update_idletasks()
            del app
        finally:
            root.destroy()
        expected = "文档格式设置程序060921"
        if title == expected:
            lines.append(f"PASS  窗口标题正确: {title}")
        else:
            failures.append(f"窗口标题不符: {title!r} != {expected!r}")
            lines.append(f"FAIL  窗口标题不符: {title!r} != {expected!r}")
    except Exception as ex:
        failures.append(f"Tk 界面初始化失败: {ex}")
        lines.append(f"FAIL  Tk 界面初始化失败: {ex}")

    # ---- 5) Word/WPS 可用性（非致命，仅提示） ----
    try:
        from universal import Universal
        app_obj = Universal().get_active_word_app()
        if app_obj is not None:
            lines.append("INFO  已检测到运行中的 Word/WPS，可执行格式调整")
        else:
            lines.append("INFO  未检测到运行中的 Word/WPS"
                         "（目标机器需安装 Microsoft Word 或 WPS 文字）")
    except Exception as ex:
        lines.append(f"INFO  检测 Word/WPS 时出错（不影响自检结论）: {ex}")

    lines.append("")
    if failures:
        lines.append(f"自检结论：失败（{len(failures)} 项）")
        lines.append(f"报告文件：{report_path}")
        try:
            with open(report_path, 'w', encoding='utf-8') as fh:
                _emit(lines, fh)
        except Exception:
            _emit(lines, sys.stderr)
        return 1

    lines.append("自检结论：全部通过")
    lines.append(f"报告文件：{report_path}")
    try:
        with open(report_path, 'w', encoding='utf-8') as fh:
            _emit(lines, fh)
    except Exception:
        _emit(lines, sys.stderr)
    return 0


def main():
    dpm.load_settings()
    
    root = tk.Tk()
    
    from main_form import MainForm
    app = MainForm(root)
    
    root.mainloop()


if __name__ == "__main__":
    # --selftest / --version 只在需要时执行，正常启动路径完全不受影响
    if len(sys.argv) > 1 and sys.argv[1] in ("--selftest", "-selftest"):
        sys.exit(selftest())
    if len(sys.argv) > 1 and sys.argv[1] in ("--version", "-v"):
        print("文档格式设置程序060921")
        sys.exit(0)
    main()
