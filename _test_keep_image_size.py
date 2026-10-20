"""「不改变图片大小」自测：不需要 Word。

用假 Document / InlineShapes / PageSetup 验证
SetDocumentFormat.set_images_and_tables(keep_image_size=...) 的行为：

  1. keep_image_size=True（勾选"不改变图片大小"）时，即使"最大宽度"仍
     勾选，也不写 InlineShape.Width，图片原宽度不变；
  2. 不勾选时保持原行为：宽图按页面正文宽度缩放；
  3. max_width=False（"最大宽度"未勾选）时本来就不改宽度；
  4. "只处理表格"时完全不碰图片；
  5. 不改变大小时，与尺寸无关的处理（段落不缩进）仍然执行；
  6. 非图片的内嵌对象（图表等）任何情况下都不改宽度。

运行：python _test_keep_image_size.py
"""
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 只需要 COM 常量，不真的启动 Word；没有 pywin32 时也能跑。
if "win32com" not in sys.modules:
    win32com = types.ModuleType("win32com")
    win32com.client = types.ModuleType("win32com.client")
    sys.modules["win32com"] = win32com
    sys.modules["win32com.client"] = win32com.client

import set_document_format as sdf  # noqa: E402

PAGE_WIDTH = 595.0
LEFT_MARGIN = 72.0
RIGHT_MARGIN = 72.0
BODY_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN     # 451.0

INDENT_ATTRS = ('LeftIndent', 'RightIndent', 'FirstLineIndent',
                'CharacterUnitLeftIndent', 'CharacterUnitRightIndent',
                'CharacterUnitFirstLineIndent')


class FakeParagraphFormat:
    def __init__(self):
        object.__setattr__(self, 'values', {})

    def __setattr__(self, name, value):
        if name == 'values':
            object.__setattr__(self, name, value)
            return
        self.values[name] = value

    def __getattr__(self, name):
        # clear_indents 只写不读；读不到时按 0 处理
        return self.values.get(name, 0)


class FakeRange:
    def __init__(self):
        self.ParagraphFormat = FakeParagraphFormat()


class FakeInlineShape:
    """假嵌入型对象；width_writes 记录每一次 Width 赋值。"""

    def __init__(self, width=200.0, shape_type=3):
        object.__setattr__(self, 'shape_type', shape_type)
        object.__setattr__(self, '_width', width)
        object.__setattr__(self, 'width_writes', [])
        self.Range = FakeRange()

    @property
    def Type(self):
        return self.shape_type

    @property
    def Width(self):
        return self._width

    @Width.setter
    def Width(self, value):
        self.width_writes.append(value)
        object.__setattr__(self, '_width', value)


class FakeInlineShapes:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def __iter__(self):
        return iter(self.items)


class FakeEmptyCollection:
    Count = 0

    def __iter__(self):
        return iter(())

    def __call__(self, index):
        raise IndexError(index)


class FakePageSetup:
    PageWidth = PAGE_WIDTH
    LeftMargin = LEFT_MARGIN
    RightMargin = RIGHT_MARGIN


class FakeApplication:
    @staticmethod
    def LinesToPoints(lines):
        return float(lines) * 12.0


class FakeDocument:
    def __init__(self, inline_shapes):
        self.InlineShapes = FakeInlineShapes(inline_shapes)
        self.Shapes = FakeEmptyCollection()
        self.Tables = FakeEmptyCollection()
        self.PageSetup = FakePageSetup()
        self.Application = FakeApplication()


def run(inline_shapes, keep_image_size=False, max_width=True,
        table_only=False, no_indent=True):
    doc = FakeDocument(inline_shapes)
    formatter = sdf.SetDocumentFormat.__new__(sdf.SetDocumentFormat)
    formatter.work_doc = doc
    formatter.set_images_and_tables(
        wrap_as_inline=True,
        no_indent=no_indent,
        max_width=max_width,
        table_only=table_only,
        keep_image_size=keep_image_size)
    return doc


def test_keep_image_size_skips_resize_even_with_max_width():
    shape = FakeInlineShape(width=200.0)
    run([shape], keep_image_size=True, max_width=True)

    assert shape.width_writes == [], \
        f"勾选「不改变图片大小」后仍写了宽度：{shape.width_writes}"
    assert shape.Width == 200.0, f"图片宽度被改变：{shape.Width}"
    print("PASS 勾选「不改变图片大小」时，即使「最大宽度」勾选也不改宽度")


def test_without_flag_still_resizes():
    shape = FakeInlineShape(width=200.0)
    run([shape], keep_image_size=False, max_width=True)

    assert shape.width_writes == [BODY_WIDTH], \
        f"未勾选时应按正文宽度缩放，实际写入：{shape.width_writes}"
    assert shape.Width == BODY_WIDTH
    print(f"PASS 未勾选时仍按正文宽度缩放：{BODY_WIDTH:.0f} 磅")


def test_max_width_off_does_not_resize():
    shape = FakeInlineShape(width=200.0)
    run([shape], keep_image_size=False, max_width=False)

    assert shape.width_writes == [], \
        f"「最大宽度」未勾选时不应写宽度：{shape.width_writes}"
    assert shape.Width == 200.0
    print("PASS 「最大宽度」未勾选时不改图片宽度")


def test_table_only_does_not_resize():
    shape = FakeInlineShape(width=200.0)
    run([shape], keep_image_size=False, max_width=True, table_only=True)

    assert shape.width_writes == [], \
        f"「只处理表格」时不应碰图片：{shape.width_writes}"
    assert shape.Width == 200.0
    print("PASS 「只处理表格」时不改图片宽度")


def test_no_indent_still_applied_with_keep_size():
    shape = FakeInlineShape(width=200.0)
    run([shape], keep_image_size=True, no_indent=True)

    values = shape.Range.ParagraphFormat.values
    for attr in INDENT_ATTRS:
        assert attr in values and abs(float(values[attr])) < 0.01, \
            f"不改变大小时「不缩进」未执行：{attr}={values.get(attr)}"
    print("PASS 不改变图片大小时，段落不缩进等其它处理仍然执行")


def test_non_picture_never_resized():
    chart = FakeInlineShape(width=300.0, shape_type=1)   # 非图片
    run([chart], keep_image_size=False, max_width=True)

    assert chart.width_writes == [], \
        f"非图片对象不应改宽度：{chart.width_writes}"
    print("PASS 非图片内嵌对象任何情况下都不改宽度")


if __name__ == "__main__":
    test_keep_image_size_skips_resize_even_with_max_width()
    test_without_flag_still_resizes()
    test_max_width_off_does_not_resize()
    test_table_only_does_not_resize()
    test_no_indent_still_applied_with_keep_size()
    test_non_picture_never_resized()
    print("\n全部通过")
