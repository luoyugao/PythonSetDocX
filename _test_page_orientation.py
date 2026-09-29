"""纸张方向自测：不需要 Word，用假 PageSetup 验证
SetDocumentFormat.set_page_orientation 的行为。

运行：python _test_page_orientation.py

覆盖的场景：
  1. 纵向 A4 选"横向" → Orientation=1，宽高互换；
  2. 已经是横向再选"横向" → 不重复写宽高（避免 Word 重排与数值舍入）；
  3. Orientation 置位后 Word 没有自动换宽高（自定义纸张常见）→ 显式兜住；
  4. 非标准纸张（26 x 18.5 厘米）→ 按原尺寸交换，不被统一成 A4；
  5. 多节文档（含一节本来就是横向）→ 全部统一，逐节独立交换；
  6. 每行字符数 / 每页行数清零，并调用 Repaginate 与切到页面视图；
  7. 设置失败的单节不影响其它节；没有文档时返回 False。
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


A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
CM_TO_PT = 28.35


def cm(value):
    return value * CM_TO_PT


class FakePageSetup:
    """假的 PageSetup，模拟 Word 换方向时的宽高交换。

    word_swaps_on_orientation=False 用来模拟"置位 Orientation 但不换宽高"
    的顽固纸张；fail_on_orientation=True 用来模拟该节无法设置方向。
    """

    def __init__(self, width, height, orientation=None,
                 word_swaps_on_orientation=True, fail_on_orientation=False):
        # 记录容器必须最先建立，否则下面写 PageWidth / PageHeight 时
        # __setattr__ 找不到列表（内部用 setdefault 兜底，但语义要清楚）
        self.width_writes = []
        self.height_writes = []
        # 记下 Orientation 被显式置位的次数：用来断言"没有重复置位"。
        # 注意不能用计数器代替 width_writes —— 置位时本类会交换宽高，
        # width_writes 只在代码显式写 PageWidth 时才增加。
        self.orientation_writes = []
        # 构造时写宽高也会被记录：测试里先 clear() 再断言即可
        self.PageWidth = width
        self.PageHeight = height
        self.CharsPerLine = 40
        self.LinesPage = 40
        self._orientation = orientation
        self.word_swaps = word_swaps_on_orientation
        self.fail_on_orientation = fail_on_orientation

    @property
    def Orientation(self):
        return self._orientation

    @Orientation.setter
    def Orientation(self, value):
        if self.fail_on_orientation:
            raise RuntimeError("该节无法设置纸张方向")
        self.orientation_writes.append(value)
        changed = (self._orientation != value)
        self._orientation = value
        if changed and self.word_swaps:
            self.PageWidth, self.PageHeight = self.PageHeight, self.PageWidth

    def __setattr__(self, name, value):
        # 记录显式写宽高的次数：断言"已经摆正就不写"要用到
        if name == "PageWidth":
            self.__dict__.setdefault("width_writes", []).append(value)
        elif name == "PageHeight":
            self.__dict__.setdefault("height_writes", []).append(value)
        object.__setattr__(self, name, value)


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


class FakeView:
    def __init__(self):
        self.Type = 1


class FakeWindow:
    def __init__(self):
        self.View = FakeView()


class FakeDocument:
    def __init__(self, sections):
        self.Sections = FakeSections(sections)
        self.repaginate_count = 0
        self.window = FakeWindow()

    @property
    def ActiveWindow(self):
        return self.window

    def Repaginate(self):
        self.repaginate_count += 1


def make_setter(sections):
    """构造 SetDocumentFormat 实例，但绕过 __init__（不需要真 Word）。"""
    doc = FakeDocument(sections)
    formatter = sdf.SetDocumentFormat.__new__(sdf.SetDocumentFormat)
    formatter.work_doc = doc
    formatter.level_list_templates = {}
    return formatter, doc


def assert_close(actual, expected, message):
    assert abs(actual - expected) < 0.01, \
        f"{message}：期望 {expected:.2f}，实际 {actual:.2f}"


def test_portrait_to_landscape():
    ps = FakePageSetup(cm(A4_WIDTH_CM), cm(A4_HEIGHT_CM), wc.wdOrientPortrait)
    formatter, doc = make_setter([FakeSection(ps)])

    assert formatter.set_page_orientation(False) is True
    assert ps.Orientation == wc.wdOrientLandscape, ps.Orientation
    assert_close(ps.PageWidth, cm(A4_HEIGHT_CM), "横向后纸张宽度应为长边")
    assert_close(ps.PageHeight, cm(A4_WIDTH_CM), "横向后纸张高度应为短边")
    assert ps.CharsPerLine == 0 and ps.LinesPage == 0, "页面网格应清零"
    assert doc.repaginate_count == 1, "应调用 Repaginate 刷新分页"
    assert doc.window.View.Type == wc.wdPrintView, "应切到页面视图"


def test_landscape_stays_landscape_without_rewrite():
    ps = FakePageSetup(cm(A4_HEIGHT_CM), cm(A4_WIDTH_CM), wc.wdOrientLandscape)
    formatter, _ = make_setter([FakeSection(ps)])

    assert formatter.set_page_orientation(False) is True
    ps.width_writes.clear()
    ps.height_writes.clear()
    ps.orientation_writes.clear()

    assert formatter.set_page_orientation(False) is True
    assert ps.width_writes == [], f"已摆正的页面不应重复写宽度：{ps.width_writes}"
    assert ps.height_writes == [], f"已摆正的页面不应重复写高度：{ps.height_writes}"
    # Orientation 始终按目标值置位（幂等）：不这样写，一旦"方向属性"与"纸张宽高"
    # 已经互相矛盾（属性=横向、宽<高），只改宽高就会留下不一致的页面设置。
    # 这里断言置位确实发生了、且值正确——真正的"不重复写"针对的是宽高，见上两行。
    assert ps.orientation_writes == [wc.wdOrientLandscape], \
        f"应把方向幂等置位为横向：{ps.orientation_writes}"


def test_custom_paper_not_swapped_by_word():
    """Orientation 置位但 Word 没换宽高时必须显式交换，否则是"横向的纵向纸"。"""
    ps = FakePageSetup(cm(A4_WIDTH_CM), cm(A4_HEIGHT_CM), wc.wdOrientPortrait,
                       word_swaps_on_orientation=False)
    formatter, _ = make_setter([FakeSection(ps)])

    assert formatter.set_page_orientation(False) is True
    assert ps.Orientation == wc.wdOrientLandscape
    assert_close(ps.PageWidth, cm(A4_HEIGHT_CM), "未自动交换时应由代码兜住")
    assert_close(ps.PageHeight, cm(A4_WIDTH_CM), "未自动交换时应由代码兜住")


def test_nonstandard_paper_keeps_its_size():
    """非标准纸张按自身宽高交换，不能被统一成 A4。"""
    width, height = cm(26.0), cm(18.5)
    ps = FakePageSetup(width, height, wc.wdOrientLandscape)
    formatter, _ = make_setter([FakeSection(ps)])

    assert formatter.set_page_orientation(True) is True
    assert ps.Orientation == wc.wdOrientPortrait
    assert_close(ps.PageWidth, height, "纵向后宽度应为原高度")
    assert_close(ps.PageHeight, width, "纵向后高度应为原宽度")


def test_multiple_sections_mixed_orientation():
    first = FakePageSetup(cm(A4_WIDTH_CM), cm(A4_HEIGHT_CM), wc.wdOrientPortrait)
    second = FakePageSetup(cm(A4_HEIGHT_CM), cm(A4_WIDTH_CM), wc.wdOrientLandscape)
    third = FakePageSetup(cm(26.0), cm(18.5), wc.wdOrientLandscape)
    formatter, _ = make_setter([FakeSection(first), FakeSection(second),
                                FakeSection(third)])

    assert formatter.set_page_orientation(False) is True

    for index, ps in enumerate((first, second, third), start=1):
        assert ps.Orientation == wc.wdOrientLandscape, f"第{index}节方向不对"
        assert ps.PageWidth > ps.PageHeight, f"第{index}节宽高没摆成横向"
    assert_close(third.PageWidth, cm(26.0), "第三节应保留自己的纸张尺寸")


def test_failure_in_one_section_does_not_break_others():
    bad = FakePageSetup(cm(A4_WIDTH_CM), cm(A4_HEIGHT_CM), wc.wdOrientPortrait,
                        fail_on_orientation=True)
    good = FakePageSetup(cm(A4_WIDTH_CM), cm(A4_HEIGHT_CM), wc.wdOrientPortrait)
    formatter, doc = make_setter([FakeSection(bad), FakeSection(good)])

    assert formatter.set_page_orientation(False) is True
    assert bad.PageWidth < bad.PageHeight, "失败的节不应被改动"
    assert good.PageWidth > good.PageHeight, "其它节仍应完成设置"
    assert doc.repaginate_count == 1


def test_no_document():
    formatter = sdf.SetDocumentFormat.__new__(sdf.SetDocumentFormat)
    formatter.work_doc = None
    formatter.level_list_templates = {}
    assert formatter.set_page_orientation(False) is False


def main():
    tests = [
        test_portrait_to_landscape,
        test_landscape_stays_landscape_without_rewrite,
        test_custom_paper_not_swapped_by_word,
        test_nonstandard_paper_keeps_its_size,
        test_multiple_sections_mixed_orientation,
        test_failure_in_one_section_does_not_break_others,
        test_no_document,
    ]
    for test in tests:
        test()
        print(f"通过：{test.__name__}")
    print(f"\n全部通过（{len(tests)} 项）")


if __name__ == "__main__":
    main()
