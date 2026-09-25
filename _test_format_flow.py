"""端到端流程自测：不开 Word，用假文档 + 假 SetDocumentFormat 走一遍
MainForm._on_format_adjust，验证：

  1. 「变更正文格式」勾选时确实是流程的第一步（第 8 条）；
  2. 流程里每一个 set_xxx(progress=…) 调用都能被接受（不会再出现
     "got an unexpected keyword argument 'progress'"）；
  3. 正文、图片与表格、标题、收尾各步骤的调用顺序符合设计。

运行：python _test_format_flow.py
"""
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk  # noqa: E402
from tkinter import messagebox  # noqa: E402

import main_form  # noqa: E402

# 出错时不要弹模态对话框卡住自测，改为记录错误信息
messagebox.showerror = lambda title, msg: main_form.__dict__.setdefault(
    "_test_errors", []).append(msg)
main_form.messagebox.showerror = messagebox.showerror
main_form.messagebox.showwarning = lambda title, msg: None


# ---------------------------------------------------------------
# 假的 SetDocumentFormat：记录调用顺序与关键字参数
# ---------------------------------------------------------------
class FakeSetDocumentFormat:
    def __init__(self, full_name=None, doc=None):
        self.work_doc = doc
        FakeSetDocumentFormat.calls.append(("__init__", ()))

    calls = []

    def _record(self, name, **kwargs):
        FakeSetDocumentFormat.calls.append((name, kwargs))

    # 每个方法都显式声明 progress —— 与真实实现保持同样的签名约束
    def snapshot_protected_paragraph_format(self, **kw):
        self._record("snapshot_protected_paragraph_format", **kw)
        return {}

    def restore_protected_paragraph_format(self, snapshots, **kw):
        self._record("restore_protected_paragraph_format", **kw)

    def set_content_format(self, indent_style, alignment, font, font_size,
                           code_paragraphs=None, is_delete_empty_lines=False,
                           is_standard_line_spacing=False, progress=None):
        self._record("set_content_format", progress=progress)
        # 进度回调必须真的能调用
        assert callable(progress), "set_content_format 未收到可调用的 progress"
        progress("正在变更正文格式… 5/500")

    def set_content_style(self, indent_style=None, alignment=None,
                          font=None, font_size=None, progress=None):
        self._record("set_content_style", progress=progress)
        assert callable(progress), "set_content_style 未收到可调用的 progress"
        progress("正在同步「正文」样式…")

    def set_page_margins(self, *args):
        self._record("set_page_margins")

    def set_main_title_format(self, font, font_size, is_bold, progress=None):
        self._record("set_main_title_format", progress=progress)
        assert callable(progress)

    def set_images_and_tables(self, wrap_as_inline=True, no_indent=True,
                              max_width=True, table_only=False, progress=None):
        self._record("set_images_and_tables", table_only=table_only,
                     progress=progress)
        assert callable(progress), "set_images_and_tables 未收到 progress"

    def add_image_border(self, progress=None):
        self._record("add_image_border", progress=progress)
        assert callable(progress)

    def center_all_images(self, progress=None):
        self._record("center_all_images", progress=progress)
        assert callable(progress)

    def set_tables_auto_adjust_and_align(self, progress=None):
        self._record("set_tables_auto_adjust_and_align", progress=progress)
        assert callable(progress)

    def set_title_styles(self, level_or_para, font=None, font_size=None,
                         number_style=None, indent_style=None, progress=None):
        self._record("set_title_styles", level=level_or_para, progress=progress)
        assert callable(progress)
        return None

    def set_image_paragraph_no_indent(self, progress=None):
        self._record("set_image_paragraph_no_indent", progress=progress)
        assert callable(progress)
        return set()

    def set_table_paragraph_no_indent(self, progress=None):
        self._record("set_table_paragraph_no_indent", progress=progress)
        assert callable(progress)
        return set()

    def set_table_surrounding_spacing(self, points=18.0, progress=None):
        self._record("set_table_surrounding_spacing", progress=progress)
        assert callable(progress)

    # 标题流程内部会用到的方法
    def get_paragraph_outline_level(self, paragraph):
        return 0

    def set_toc_line_spacing(self, *a, **k):
        self._record("set_toc_line_spacing")


class FakeRange:
    Text = "测试文档标题\r"


class FakeParagraphs:
    Count = 3

    def __call__(self, index):
        return types.SimpleNamespace(Range=FakeRange())

    def __getitem__(self, index):
        return self(index)


class FakeDocument:
    FullName = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "假文档.docx")
    Paragraphs = FakeParagraphs()

    def Activate(self):
        pass


class FakeWordApp:
    ActiveDocument = FakeDocument()


def install_fake_module():
    """把假的 set_document_format 模块塞进 sys.modules。

    _on_format_adjust 内部用 `from set_document_format import SetDocumentFormat`
    延迟导入，因此替换 sys.modules 里的条目即可注入假实现。
    """
    fake = types.ModuleType("set_document_format")
    fake.SetDocumentFormat = FakeSetDocumentFormat
    sys.modules["set_document_format"] = fake


def install_fake_page_number_module():
    fake = types.ModuleType("page_number_manager")

    def add_page_numbers_custom(doc):
        FakeSetDocumentFormat.calls.append(("add_page_numbers_custom", {}))

    fake.PageNumberManager = types.SimpleNamespace(
        add_page_numbers_custom=add_page_numbers_custom)
    sys.modules["page_number_manager"] = fake


def build_form(table_only):
    """构建主窗体并勾选各步骤；table_only 控制"只处理表格"。

    每次只用一个 Tk root：窗体内部有 1 秒周期的活动文档轮询，销毁窗口后
    再跑第二遍会让该轮询命中已销毁的控件，故本测试只跑单遍。
    """
    root = tk.Tk()
    app = main_form.MainForm(root)
    root.withdraw()

    # 勾选需要走到的各步骤
    app.chk_change_content_format.set(True)
    app.chk_change_page_margin.set(True)
    app.chk_add_page_num.set(True)
    app.chk_change_main_title_format.set(True)
    app.chk_change_image_and_table_format.set(True)
    app.chk_change_level_title_format.set(True)
    app.chk_image_no_indent.set(True)
    app.chk_table_only.set(table_only)
    app.cmb_content_indent.set("首行缩进2字符")
    app.cmb_content_align.set("两端对齐")

    # 记录状态栏依次收到的文字（最终文字会被"完成"覆盖，故需留痕）
    status_history = []
    original_set_status = app._set_status

    def recording_set_status(text):
        status_history.append(text)
        original_set_status(text)

    app._set_status = recording_set_status
    app.status_history = status_history
    return root, app


def main():
    install_fake_module()
    install_fake_page_number_module()

    # "只处理表格"=否：图片边框、图片居中、图片段落缩进等步骤全部走到，
    # 每一个 set_xxx(progress=…) 调用都在此被验证（本次报错的根因就在这里）
    root, app = build_form(table_only=False)

    app.word_app = FakeWordApp()
    FakeSetDocumentFormat.calls = []
    app._on_format_adjust()

    names = [name for name, _ in FakeSetDocumentFormat.calls
             if name != "__init__"]
    errors = getattr(main_form, "_test_errors", [])
    print("调用顺序：")
    for n in names:
        print("   ", n)
    if errors:
        print("处理过程中的异常：", errors)
    print("最终状态栏：", app.status_bar.cget("text"))

    assert not errors, f"流程中抛出异常：{errors}"

    # ---- 断言 1：正文格式必须是第一步 ----
    assert names[0] == "snapshot_protected_paragraph_format", names
    assert names[1] == "set_content_format", \
        f"「变更正文格式」应为第一步，实际顺序：{names}"
    assert names[2] == "set_content_style", names

    # ---- 断言 2：每一步都走到了 ----
    for expected in ("set_page_margins", "add_page_numbers_custom",
                     "set_main_title_format", "set_images_and_tables",
                     "add_image_border", "center_all_images",
                     "set_tables_auto_adjust_and_align",
                     "set_image_paragraph_no_indent",
                     "set_table_paragraph_no_indent",
                     "set_table_surrounding_spacing"):
        assert expected in names, f"缺少步骤 {expected}：{names}"

    # ---- 断言 3：收尾步骤在正文/标题之后 ----
    assert names.index("set_table_paragraph_no_indent") > \
        names.index("set_content_style")
    assert names.index("set_table_surrounding_spacing") == len(names) - 1, \
        f"表格间距应为最后一步：{names}"

    # ---- 断言 4：需要进度上报的步骤都收到了可调用的 progress ----
    # snapshot/restore_protected_paragraph_format 与 set_page_margins、
    # add_page_numbers_custom 不承担细粒度进度上报，不在检查范围内。
    must_report = ("set_content_format", "set_content_style",
                   "set_main_title_format", "set_images_and_tables",
                   "add_image_border", "center_all_images",
                   "set_tables_auto_adjust_and_align",
                   "set_image_paragraph_no_indent",
                   "set_table_paragraph_no_indent",
                   "set_table_surrounding_spacing")
    reported = {n for n, kw in FakeSetDocumentFormat.calls if "progress" in kw}
    missing = [n for n in must_report if n not in reported]
    assert not missing, f"以下步骤没收到 progress：{missing}"

    # ---- 断言 5：状态栏确实收到了细粒度进度文字 ----
    assert any("正文" in s for s in app.status_history), app.status_history
    fine = [s for s in app.status_history if "第" in s and "项" in s]
    print("细粒度进度示例：", fine[:3] if fine else "(无)")
    assert fine, f"未收到细粒度进度：{app.status_history}"

    root.destroy()
    print("\n全部通过：流程顺序与 progress 参数传递均正确")


if __name__ == "__main__":
    main()
