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
from word_constants import (  # noqa: E402
    clear_indents, _CHAR_UNIT_INDENT_ATTRS, _POINT_INDENT_ATTRS)

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


class InheritingParagraphFormat(FakeParagraphFormat):
    """模拟 Word 的一条实测规则：**字符单位值 = 0 视为"未设置直接格式"**。

    「正文」样式带着"首行缩进2字符"，首段自己没有直接缩进。若先写点值 0、
    再写字符单位 0，Word 会把刚写好的直接 0 当成"未设置"，段落退回继承
    「正文」的 2 字符缩进——总标题于是又缩进了，居中显示也跟着偏右。
    只有把点值 0 放在最后写，直接 0 才钉得住。
    """

    def __init__(self, inherited_first_line=2 * 0.35 * 28.35):
        object.__setattr__(self, '_values', {})
        object.__setattr__(self, 'writes', [])
        object.__setattr__(self, '_inherited_first_line', inherited_first_line)

    def __getattr__(self, name):
        values = object.__getattribute__(self, '_values')
        if name == 'FirstLineIndent':
            if name in values:
                return values[name]
            # 没有直接格式 -> 继承「正文」的 2 字符缩进
            return object.__getattribute__(self, '_inherited_first_line')
        if name == 'CharacterUnitFirstLineIndent':
            if 'FirstLineIndent' in values:
                return 0.0
            return 2.0
        if name in _INDENT_ATTRS:
            return values.get(name, 0.0)
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name in _INDENT_ATTRS:
            values = object.__getattribute__(self, '_values')
            values[name] = value
            self.writes.append((name, value))
            if name.startswith('CharacterUnit') and value == 0:
                # Word 把"字符单位值=0"当成未设置直接格式，连带抹掉点值直接格式
                values.pop('FirstLineIndent', None)
                values.pop('LeftIndent', None)
            if name == 'FirstLineIndent' and value == 0:
                values['CharacterUnitFirstLineIndent'] = 0
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


def test_clear_indents_char_unit_first_point_last():
    """缩进清零必须"先字符单位、后点值"：字符单位 0 最后写会被 Word 忽略。"""
    fmt = FakeParagraphFormat()
    clear_indents(FakeRange(paragraph_format=fmt))
    order = [name for name, _ in fmt.writes]
    assert order == list(_CHAR_UNIT_INDENT_ATTRS + _POINT_INDENT_ATTRS), \
        f"缩进清零的写入顺序不对：{order}"
    print("PASS 缩进清零按\"先字符单位、后点值\"执行")


def test_no_indent_wins_over_inherited_body_style():
    """"变更正文格式"未勾选时「正文」仍是首行缩进2字符，首段必须仍无缩进。"""
    fmt = InheritingParagraphFormat()
    para = FakeParagraph(paragraph_format=fmt)
    doc = FakeDocument(para)
    make_formatter(doc).set_main_title_format("黑体", "二号", True)

    assert abs(fmt.FirstLineIndent) < 0.01, \
        f"首段仍继承到首行缩进：{fmt.FirstLineIndent}"
    assert abs(fmt.CharacterUnitFirstLineIndent) < 0.01, \
        f"首段的字符单位首行缩进未清零：{fmt.CharacterUnitFirstLineIndent}"
    assert para.Alignment == 1, f"首段未水平居中：{para.Alignment}"
    print("PASS 父样式带\"首行缩进2字符\"时，首段仍无缩进且水平居中")


def test_layout_only_keeps_font_size_and_spacing():
    """总标题组未勾选时：只强制居中+无缩进，字体/字号/段距保持原样。"""
    font = FakeFont(name="仿宋", size=22.0)
    fmt = InheritingParagraphFormat()
    para = FakeParagraph(font=font, paragraph_format=fmt)
    para.Alignment = 0            # 原来是左对齐
    para.SpaceBefore = 12.0       # 作者自己定的段前段后
    para.SpaceAfter = 6.0
    doc = FakeDocument(para)

    assert make_formatter(doc).set_main_title_layout() is True
    assert para.Alignment == 1, f"首段未强制居中：{para.Alignment}"
    assert abs(fmt.FirstLineIndent) < 0.01, \
        f"首段首行缩进未清零：{fmt.FirstLineIndent}"
    assert abs(fmt.CharacterUnitFirstLineIndent) < 0.01, \
        f"首段字符单位首行缩进未清零：{fmt.CharacterUnitFirstLineIndent}"
    assert font.Name == "仿宋" and font.Size == 22.0, \
        f"字体字号被改动：{font.Name} / {font.Size}"
    assert para.SpaceBefore == 12.0 and para.SpaceAfter == 6.0, \
        f"段前段后被改动：{para.SpaceBefore} / {para.SpaceAfter}"
    print("PASS 总标题组未勾选时：首行强制居中、无缩进，字体字号与段距保持原样")


def test_layout_without_document_returns_false():
    """没有文档时只返回 False，不抛异常。"""
    formatter = SetDocumentFormat.__new__(SetDocumentFormat)
    formatter.work_doc = None
    assert formatter.set_main_title_layout() is False
    print("PASS 无文档时 set_main_title_layout 安全返回 False")


if __name__ == "__main__":
    test_hanging_indent_cleared()
    test_first_line_indent_2chars_cleared()
    test_keep_font_and_size_still_clears_indent()
    test_center_and_spacing_applied()
    test_clear_indents_char_unit_first_point_last()
    test_no_indent_wins_over_inherited_body_style()
    test_layout_only_keeps_font_size_and_spacing()
    test_layout_without_document_returns_false()
    print("\n全部通过")
