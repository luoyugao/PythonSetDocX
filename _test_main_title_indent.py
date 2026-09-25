"""总标题（首段）"缩进方式=无"的离线自测（不需要 Word）。

用假对象模拟 Document / Paragraph / Style / Range，验证 set_main_title_format：
  1. MainTitle 样式与首段段落的 6 个缩进属性全部被写为 0；
  2. 首段原有的缩进（含悬挂缩进：LeftIndent>0 且 FirstLineIndent<0，
     以及中文的字符单位"缩进2字符"）不会残留；
  3. 字体/字号选"不变更"时，缩进清零与字体字号按原值写回互不影响。

运行：python _test_main_title_indent.py
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from set_document_format import SetDocumentFormat  # noqa: E402

_INDENT_ATTRS = ('LeftIndent', 'RightIndent', 'FirstLineIndent',
                 'CharacterUnitLeftIndent', 'CharacterUnitRightIndent',
                 'CharacterUnitFirstLineIndent')


class FakeParagraphFormat:
    """记录缩进写入的假 ParagraphFormat。"""

    def __init__(self, values=None):
        object.__setattr__(self, '_values', dict(values or {}))
        object.__setattr__(self, 'writes', [])

    def __getattr__(self, name):
        if name in _INDENT_ATTRS:
            values = object.__getattribute__(self, '_values')
            if name in values:
                return values[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name in _INDENT_ATTRS:
            self._values[name] = value
            self.writes.append((name, value))
            return
        object.__setattr__(self, name, value)


class FakeFont:
    def __init__(self, name=None, size=None):
        self.Name = name
        self.NameFarEast = name
        self.NameAscii = name
        self.NameOther = name
        self.Size = size
        self.Bold = None


class FakeStyle:
    """假 Style：Font 与 ParagraphFormat 均可写。"""

    def __init__(self):
        self.Font = FakeFont()
        self.ParagraphFormat = FakeParagraphFormat()


class FakeRange:
    def __init__(self, font=None, paragraph_format=None):
        self.Font = font or FakeFont()
        self.ParagraphFormat = paragraph_format or FakeParagraphFormat()
        self.style = None

    def set_Style(self, style):
        self.style = style


class FakeParagraph:
    def __init__(self, font=None, paragraph_format=None):
        self.Range = FakeRange(font, paragraph_format)
        self.Alignment = None
        self.SpaceBefore = None
        self.SpaceAfter = None


class FakeStyles:
    def __init__(self):
        self.main_title = FakeStyle()

    def __call__(self, name):       # doc.Styles("MainTitle")
        if name == "MainTitle":
            return self.main_title
        raise KeyError(name)

    def Add(self, name, style_type):    # 样式不存在时新建
        return self.main_title


class FakeDocument:
    def __init__(self, paragraph):
        self._paragraph = paragraph
        self.Styles = FakeStyles()

    def Paragraphs(self, index):
        assert index == 1, "总标题只处理首段"
        return self._paragraph


def make_formatter(doc):
    formatter = SetDocumentFormat.__new__(SetDocumentFormat)
    formatter.work_doc = doc
    return formatter


def test_hanging_indent_cleared():
    """首段是悬挂缩进 + 字符单位缩进时，全部归零。"""
    original = {
        'LeftIndent': 2 * 0.35 * 28.35,
        'RightIndent': 1.0,
        'FirstLineIndent': -2 * 0.35 * 28.35,
        'CharacterUnitLeftIndent': 2.0,
        'CharacterUnitRightIndent': 1.0,
        'CharacterUnitFirstLineIndent': -2.0,
    }
    para = FakeParagraph(paragraph_format=FakeParagraphFormat(original))
    doc = FakeDocument(para)
    make_formatter(doc).set_main_title_format("黑体", "二号", True)

    style_indents = {a: getattr(doc.Styles.main_title.ParagraphFormat, a)
                     for a in _INDENT_ATTRS}
    assert style_indents == {a: 0 for a in _INDENT_ATTRS}, \
        f"MainTitle 样式缩进未全部清零：{style_indents}"
    para_indents = {a: getattr(para.Range.ParagraphFormat, a)
                    for a in _INDENT_ATTRS}
    assert para_indents == {a: 0 for a in _INDENT_ATTRS}, \
        f"首段段落缩进未全部清零：{para_indents}"
    print("PASS 悬挂缩进的首段：样式与段落的 6 个缩进属性全部归零")


def test_first_line_indent_2chars_cleared():
    """中文的"首行缩进2字符"（存在字符单位属性里）也被清掉。"""
    original = {'FirstLineIndent': 2 * 0.35 * 28.35,
                'CharacterUnitFirstLineIndent': 2.0}
    para = FakeParagraph(paragraph_format=FakeParagraphFormat(original))
    doc = FakeDocument(para)
    make_formatter(doc).set_main_title_format(None, None, False)

    fmt = para.Range.ParagraphFormat
    assert getattr(fmt, 'CharacterUnitFirstLineIndent') == 0, \
        "字符单位的首行缩进未清零"
    assert getattr(fmt, 'FirstLineIndent') == 0, "点值首行缩进未清零"
    print("PASS 首段\"首行缩进2字符\"被清零（点值与字符单位）")


def test_keep_font_and_size_still_clears_indent():
    """字体/字号选"不变更"时：原字体字号写回，缩进照样清零。"""
    para = FakeParagraph(
        font=FakeFont(name="仿宋", size=22.0),
        paragraph_format=FakeParagraphFormat({'LeftIndent': 10.0}))
    doc = FakeDocument(para)
    make_formatter(doc).set_main_title_format(None, None, True)

    style = doc.Styles.main_title
    assert style.Font.Name == "仿宋", f"原字体未写回：{style.Font.Name}"
    assert style.Font.Size == 22.0, f"原字号未写回：{style.Font.Size}"
    assert style.Font.Bold == 1, "加粗未生效"
    assert getattr(para.Range.ParagraphFormat, 'LeftIndent') == 0
    print("PASS \"不变更\"字体字号时，缩进仍被清零")


def test_center_and_spacing_applied():
    """居中、段前0磅、段后28磅属于固有格式，顺带校验未被缩进改动影响。"""
    para = FakeParagraph()
    doc = FakeDocument(para)
    make_formatter(doc).set_main_title_format("黑体", "二号", True)

    assert para.Alignment is not None, "段落未设为居中"
    assert para.SpaceBefore == 0.0 and para.SpaceAfter == 28.0, \
        f"段间距异常：{para.SpaceBefore} / {para.SpaceAfter}"
    print("PASS 首段居中、段前0磅、段后28磅正常应用")


if __name__ == "__main__":
    test_hanging_indent_cleared()
    test_first_line_indent_2chars_cleared()
    test_keep_font_and_size_still_clears_indent()
    test_center_and_spacing_applied()
    print("\n全部通过")
