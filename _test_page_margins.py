"""页边距自测：不需要 Word，用假 PageSetup 验证
SetDocumentFormat.set_page_margins 的行为。

运行：python _test_page_margins.py

覆盖的场景：
  1. 普通多节文档 → 每节边距都写成新值，返回 True；
  2. Word 2007 / 部分 WPS 没有 CharsPerLine、LayoutMode 属性（写入抛异常）
     → 不能连累后面的清理项与其它节，边距照样写成功；
  3. 先清镜像边距/装订线/文档网格，再写四边边距（顺序不能反）；
  4. 某节把第一次写入静默吞掉（页面设置对象被缓存）→ 读回校验后重写一次；
  5. 某节逐节怎么写都不生效 → 用整篇文档级 PageSetup 兜底；
  6. 某节彻底写不进去 → 返回 False，并把节号报给 progress；
  7. 写完要重排 + 强制刷新（Repaginate / 页面视图 / ScreenRefresh）；
  8. 没有文档时返回 False。
"""
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 本模块只用到 COM 常量，不需要真的 pywin32：先塞一个空壳，保证
# set_document_format 能被导入（真机上有 pywin32 时同样兼容）。
if "win32com" not in sys.modules:
    win32com = types.ModuleType("win32com")
    win32com.client = types.ModuleType("win32com.client")
    sys.modules["win32com"] = win32com
    sys.modules["win32com.client"] = win32com.client

import word_constants as wc          # noqa: E402
import set_document_format as sdf    # noqa: E402


CM_TO_PT = 28.35
MARGIN_NAMES = ("TopMargin", "BottomMargin", "LeftMargin", "RightMargin")
OTHER_NAMES = ("MirrorMargins", "Gutter", "LayoutMode", "CharsPerLine", "LinesPage")
PROPS = MARGIN_NAMES + OTHER_NAMES


def cm(value):
    return value * CM_TO_PT


class FakePageSetup:
    """假的 PageSetup。

    unsupported: 这些属性在"这个 Word 版本"上不存在，读/写都抛异常
                 （Word 2007 与部分 WPS 就是没有 CharsPerLine / LayoutMode）
    swallow_names: 这些属性的写入被静默吞掉，文档里的值不变
    swallow_first_margin_write: 每个边距属性的第一次写入被吞掉（缓存节）
    hard_reject: 所有写入都不生效，整篇文档级写入也不生效
    """

    def __init__(self, left=3.17, unsupported=(), swallow_names=(),
                 swallow_first_margin_write=False, hard_reject=False):
        object.__setattr__(self, "values", {})
        object.__setattr__(self, "writes", [])
        object.__setattr__(self, "unsupported", set(unsupported))
        object.__setattr__(self, "swallow_names", set(swallow_names))
        object.__setattr__(self, "swallow_first", swallow_first_margin_write)
        object.__setattr__(self, "swallowed_once", set())
        object.__setattr__(self, "hard_reject", hard_reject)
        for name in MARGIN_NAMES:
            self.values[name] = cm(left)
        self.values.update(MirrorMargins=0, Gutter=0.0, LayoutMode=2,
                           CharsPerLine=40, LinesPage=40)

    def __getattr__(self, name):
        if name in PROPS:
            if name in self.unsupported:
                raise AttributeError(f"<unknown>.{name}")
            return self.values.get(name)
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name not in PROPS:
            object.__setattr__(self, name, value)
            return
        self.writes.append((name, value))
        if name in self.unsupported:
            raise AttributeError(f"Property '<unknown>.{name}' can not be set.")
        if self.hard_reject or name in self.swallow_names:
            return
        if self.swallow_first and name in MARGIN_NAMES \
                and name not in self.swallowed_once:
            self.swallowed_once.add(name)
            return
        self.values[name] = value

    # 整篇文档级 PageSetup 写入：走另一条通路，能救回"逐节写不进"的节
    def write_from_document(self, name, value):
        if self.hard_reject:
            return
        self.values[name] = value

    @property
    def margins(self):
        return tuple(self.values[n] for n in MARGIN_NAMES)


class FakeSection:
    def __init__(self, page_setup):
        self.PageSetup = page_setup


class FakeSections:
    def __init__(self, sections):
        self._sections = list(sections)

    @property
    def Count(self):
        return len(self._sections)

    def __call__(self, index):
        return self._sections[index - 1]

    def __getitem__(self, index):
        return self._sections[index - 1]


class FakeDocPageSetup:
    """文档级 PageSetup：在 Word/WPS 里写它是整篇生效的。"""

    def __init__(self, sections):
        object.__setattr__(self, "_sections", sections)
        object.__setattr__(self, "writes", [])

    def __getattr__(self, name):
        if name in PROPS:
            return getattr(self._sections[0], name)
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name not in PROPS:
            object.__setattr__(self, name, value)
            return
        self.writes.append((name, value))
        for section in self._sections:
            section.write_from_document(name, value)


class FakeView:
    def __init__(self):
        self.Type = 1


class FakeWindow:
    def __init__(self):
        self.View = FakeView()


class FakeApplication:
    def __init__(self):
        self.screen_refresh_count = 0

    def ScreenRefresh(self):
        self.screen_refresh_count += 1


class FakeDocument:
    def __init__(self, page_setups):
        self.sections = [FakeSection(ps) for ps in page_setups]
        self.Sections = FakeSections(self.sections)
        self.PageSetup = FakeDocPageSetup([s.PageSetup for s in self.sections])
        self.repaginate_count = 0
        self.window = FakeWindow()
        self.Application = FakeApplication()

    @property
    def ActiveWindow(self):
        return self.window

    def Repaginate(self):
        self.repaginate_count += 1


def make_setter(page_setups):
    """构造 SetDocumentFormat 实例，但绕过 __init__（不需要真 Word）。"""
    doc = FakeDocument(page_setups)
    formatter = sdf.SetDocumentFormat.__new__(sdf.SetDocumentFormat)
    formatter.work_doc = doc
    formatter.level_list_templates = {}
    return formatter, doc


def assert_margins(page_setup, value_cm, message):
    expected = cm(value_cm)
    for name, got in zip(MARGIN_NAMES, page_setup.margins):
        assert abs(float(got) - expected) < 0.01, \
            f"{message}：{name} 期望 {expected:.2f}，实际 {got}"


def test_plain_multisection():
    setups = [FakePageSetup() for _ in range(3)]
    formatter, doc = make_setter(setups)

    assert formatter.set_page_margins(1.5, 1.5, 1.5, 1.5) is True

    for index, ps in enumerate(setups, start=1):
        assert_margins(ps, 1.5, f"第{index}节")
    assert doc.repaginate_count == 1, "应调用 Repaginate 刷新分页"
    assert doc.window.View.Type == wc.wdPrintView, "应切到页面视图"
    assert doc.Application.screen_refresh_count == 1, "应强制刷新窗口"


def test_missing_grid_properties_do_not_break_margins():
    """Word 2007 / WPS 没有 CharsPerLine、LayoutMode：抛异常也得把边距写完。"""
    setups = [FakePageSetup(unsupported=("CharsPerLine", "LayoutMode")),
              FakePageSetup(unsupported=("CharsPerLine", "LayoutMode"))]
    formatter, _ = make_setter(setups)

    assert formatter.set_page_margins(2.0, 2.0, 2.0, 2.0) is True

    for index, ps in enumerate(setups, start=1):
        assert_margins(ps, 2.0, f"第{index}节")
        # 不支持的属性被跳过后，后面的清理项仍要执行到
        assert ps.values["LinesPage"] == 0, \
            f"第{index}节：CharsPerLine 失败后 LinesPage 也应按要求清零"


def test_clear_overrides_before_writing_margins():
    """镜像边距/装订线/文档网格必须先清，再写四边边距。"""
    ps = FakePageSetup()
    formatter, _ = make_setter([ps])

    assert formatter.set_page_margins(1.5, 1.5, 1.5, 1.5) is True

    names = [name for name, _ in ps.writes]
    first_margin = min(names.index(n) for n in MARGIN_NAMES)
    for name in ("MirrorMargins", "Gutter", "LayoutMode", "CharsPerLine"):
        assert names.index(name) < first_margin, \
            f"{name} 应在写边距之前清掉：{names}"


def test_cached_section_needs_second_write():
    """第一节把第一次写入吞掉（页面设置对象被缓存），读回校验后要重写。"""
    setups = [FakePageSetup(swallow_first_margin_write=True), FakePageSetup()]
    formatter, _ = make_setter(setups)

    assert formatter.set_page_margins(1.5, 1.5, 1.5, 1.5) is True

    for index, ps in enumerate(setups, start=1):
        assert_margins(ps, 1.5, f"第{index}节")


def test_document_level_fallback():
    """某节逐节怎么写都不生效时，用整篇文档级 PageSetup 兜底。"""
    stubborn = FakePageSetup(swallow_names=MARGIN_NAMES)
    normal = FakePageSetup()
    formatter, _ = make_setter([stubborn, normal])

    assert formatter.set_page_margins(1.5, 1.5, 1.5, 1.5) is True

    assert_margins(stubborn, 1.5, "逐节写不进的节应由文档级写入救回")
    assert_margins(normal, 1.5, "正常节")


def test_unfixable_section_is_reported():
    bad = FakePageSetup(hard_reject=True)
    good = FakePageSetup()
    formatter, _ = make_setter([bad, good])
    messages = []

    assert formatter.set_page_margins(1.5, 1.5, 1.5, 1.5,
                                      progress=messages.append) is False

    assert_margins(good, 1.5, "能写的节仍要写完")
    assert any("1" in m and "页边距" in m for m in messages), \
        f"应把失败的节号报给 progress：{messages}"


def test_no_document():
    formatter = sdf.SetDocumentFormat.__new__(sdf.SetDocumentFormat)
    formatter.work_doc = None
    formatter.level_list_templates = {}
    assert formatter.set_page_margins(1.5, 1.5, 1.5, 1.5) is False


def main():
    tests = [
        test_plain_multisection,
        test_missing_grid_properties_do_not_break_margins,
        test_clear_overrides_before_writing_margins,
        test_cached_section_needs_second_write,
        test_document_level_fallback,
        test_unfixable_section_is_reported,
        test_no_document,
    ]
    for test in tests:
        test()
        print(f"通过：{test.__name__}")
    print(f"\n全部通过（{len(tests)} 项）")


if __name__ == "__main__":
    main()
