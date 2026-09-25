"""缩进"已满足则跳过"逻辑的离线自测（不需要 Word）。

用假对象模拟 Range / ParagraphFormat，验证：
  1. 已是目标缩进状态时不再写入缩进；
  2. 需要改动时照常写入点值与字符单位值；
  3. 读取不到缩进（对象不支持）时不误判，照常写入。

运行：python _test_indent_skip.py
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from set_document_format import SetDocumentFormat, _make_progress  # noqa: E402
from word_constants import _PT_PER_CHAR  # noqa: E402


class FakeParagraphFormat:
    """记录属性写入次数的假 ParagraphFormat。"""

    _ATTRS = ('LeftIndent', 'RightIndent', 'FirstLineIndent',
              'CharacterUnitLeftIndent', 'CharacterUnitRightIndent',
              'CharacterUnitFirstLineIndent')

    def __init__(self, values=None, writable=True):
        self._values = dict(values or {})
        self._writable = writable
        self.writes = []

    def __getattr__(self, name):
        if name in FakeParagraphFormat._ATTRS:
            if name in self._values:
                return self._values[name]
            # 模拟 wdUndefined：取值抛异常（上层按"读不到"处理）
            raise AttributeError(name)
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name in ('_values', '_writable', 'writes'):
            object.__setattr__(self, name, value)
            return
        if not self._writable:
            raise AttributeError(name)
        self._values[name] = value
        self.writes.append((name, value))


class FakeRange:
    def __init__(self, paragraph_format):
        self.ParagraphFormat = paragraph_format


def make_formatter():
    return SetDocumentFormat.__new__(SetDocumentFormat)


def test_skip_when_already_indented():
    fmt = FakeParagraphFormat({
        'LeftIndent': 0.0,
        'RightIndent': 0.0,
        'FirstLineIndent': 2 * _PT_PER_CHAR,
        'CharacterUnitLeftIndent': 0.0,
        'CharacterUnitFirstLineIndent': 2.0,
    })
    formatter = make_formatter()
    wrote = formatter._apply_indent_style_if_needed(
        FakeRange(fmt), "首行缩进2字符")
    assert wrote is False, "已是首行缩进2字符，应跳过写入"
    assert fmt.writes == [], f"不应有任何写入，实际：{fmt.writes}"
    print("PASS 已缩进段落不再叠加设置缩进")


def test_skip_even_for_different_style():
    """严格模式：已有缩进时，即使所选方式不同也不再改动缩进。"""
    fmt = FakeParagraphFormat({
        'LeftIndent': 2 * _PT_PER_CHAR,
        'RightIndent': 0.0,
        'FirstLineIndent': -2 * _PT_PER_CHAR,
        'CharacterUnitLeftIndent': 2.0,
        'CharacterUnitFirstLineIndent': -2.0,
    })
    formatter = make_formatter()
    wrote = formatter._apply_indent_style_if_needed(
        FakeRange(fmt), "首行缩进2字符")
    assert wrote is False, "段落已有（悬挂）缩进，应跳过写入"
    assert fmt.writes == []
    print("PASS 已有缩进（悬挂）时跳过，即使所选方式不同")


def test_skip_no_indent_does_not_clear_existing_indent():
    """严格模式：选"无缩进"时，已有缩进的段落保持原样（不再清零）。"""
    fmt = FakeParagraphFormat({
        'LeftIndent': 0.0,
        'RightIndent': 0.0,
        'FirstLineIndent': 2 * _PT_PER_CHAR,
        'CharacterUnitLeftIndent': 0.0,
        'CharacterUnitFirstLineIndent': 2.0,
    })
    formatter = make_formatter()
    wrote = formatter._apply_indent_style_if_needed(FakeRange(fmt), "无缩进")
    assert wrote is False, "已有缩进时不再写入（含不再清零）"
    assert fmt.writes == [], f"已有缩进不应被清零，实际：{fmt.writes}"
    print("PASS 已有缩进时选\"无缩进\"也保持原样")


def test_write_when_no_indent_at_all():
    """完全没有缩进的段落按所选方式写入。"""
    fmt = FakeParagraphFormat({
        'LeftIndent': 0.0,
        'RightIndent': 0.0,
        'FirstLineIndent': 0.0,
        'CharacterUnitLeftIndent': 0.0,
        'CharacterUnitFirstLineIndent': 0.0,
    })
    formatter = make_formatter()
    wrote = formatter._apply_indent_style_if_needed(
        FakeRange(fmt), "首行缩进2字符")
    assert wrote is True, "无缩进 -> 首行缩进2字符，应写入"
    assert ('CharacterUnitFirstLineIndent', 2) in fmt.writes
    assert ('FirstLineIndent', 2 * _PT_PER_CHAR) in fmt.writes
    print("PASS 无缩进段落照常设置所选缩进方式")


def test_write_when_unreadable():
    fmt = FakeParagraphFormat({})      # 所有缩进属性都读不到
    formatter = make_formatter()
    wrote = formatter._apply_indent_style_if_needed(
        FakeRange(fmt), "首行缩进2字符")
    assert wrote is True, "读不到缩进时必须照常设置，不能误判为已满足"
    assert ('CharacterUnitFirstLineIndent', 2) in fmt.writes
    print("PASS 缩进不可读时照常设置（不误跳过）")


def test_unknown_indent_style_never_writes():
    fmt = FakeParagraphFormat({'FirstLineIndent': 0.0})
    formatter = make_formatter()
    wrote = formatter._apply_indent_style_if_needed(FakeRange(fmt), "不变更")
    assert wrote is False, "未识别的缩进方式不应写入"
    assert fmt.writes == []
    print("PASS 未识别的缩进方式不做任何写入")


def test_progress_throttle():
    seen = []
    report = _make_progress(seen.append, 5)
    for i in range(1, 13):
        report(f"item {i}")
    assert seen == ["item 5", "item 10"], f"节流异常：{seen}"
    silent = _make_progress(None, 5)
    silent("nothing happens")
    print("PASS 进度上报按步长节流，无回调时静默")


if __name__ == "__main__":
    test_skip_when_already_indented()
    test_skip_even_for_different_style()
    test_skip_no_indent_does_not_clear_existing_indent()
    test_write_when_no_indent_at_all()
    test_write_when_unreadable()
    test_unknown_indent_style_never_writes()
    test_progress_throttle()
    print("\n全部通过")
