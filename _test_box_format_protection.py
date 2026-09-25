"""端到端自测：文本框内 / 单单元格表格（代码框）内的段落只允许改字号。

规则（本次实现）：
    - 文本框（形状 TextFrame）、图文框（Frame）与只有一个单元格的表格
      ——代码框的常见做法——里的段落，**除字号外**的格式一律保持原样：
      字体、对齐、缩进、行距、段前段后、样式都不动；
    - 字号跟随「正文」样式/界面"变更正文格式"里的正文字号一起变；
      段落自己带直接字号的，直接格式优先，仍保持原字号。

测试办法：造一份含上述容器的文档，跑一遍完整流程，逐属性比对处理前后的
值。流程步骤与 main_form._on_format_adjust 中的顺序一致。

运行：python _test_box_format_protection.py
（需要本机装有 Word；脚本自己新起一个 Word 实例并在结束时退出，
  不会影响你已经打开的文档。）
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import win32com.client as w  # noqa: E402

import set_document_format as sdf  # noqa: E402
import word_constants as wc  # noqa: E402

TEXT_BOX_NAME = 'TEST_TEXT_BOX'

# 逐属性比对：这些都应当保持不变（字号单独检查，不在其中）
PARAGRAPH_ATTRS = ('Alignment', 'LeftIndent', 'RightIndent', 'FirstLineIndent',
                   'CharacterUnitLeftIndent', 'CharacterUnitRightIndent',
                   'CharacterUnitFirstLineIndent', 'LineSpacingRule',
                   'SpaceBefore', 'SpaceAfter')
FONT_ATTRS = ('Name', 'NameFarEast', 'Size', 'Bold')

CONTENT_FONT = '宋体'
CONTENT_SIZE_NAME = '小四'
CONTENT_SIZE_POINTS = 12.0      # 小四 = 12磅
CONTENT_ALIGN = '两端对齐'
CONTENT_INDENT = '首行缩进2字符'


# ---------------------------------------------------------------
# 本机 Word 2013 的 Application.LinesToPoints 经 COM 调用会报错
# （"未指定的错误"），补一个等价实现，只影响"0.5行 → 磅"的折算。
# ---------------------------------------------------------------
class _AppShim:
    def __init__(self, app):
        self._app = app

    def __getattr__(self, name):
        return getattr(self._app, name)

    def LinesToPoints(self, lines):
        return float(lines) * 12.0


class _DocShim:
    def __init__(self, doc, app):
        self._doc = doc
        self.Application = app

    def __getattr__(self, name):
        return getattr(self._doc, name)


def _append_paragraph(doc, text=''):
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    if text:
        rng.InsertAfter(text)
    rng.InsertParagraphAfter()


def _append_table(doc, rows, cols):
    """在文末追加表格。

    必须先补一个空段落再插表：直接在上一个单元格的位置插表会嵌成嵌套表格，
    文档结构就不是被测的样子了。
    """
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    rng.InsertParagraphAfter()
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    return doc.Tables.Add(rng, rows, cols)


def _set_direct_format(paragraph, font, size, alignment):
    """给段落写一套"作者自己定"的直接格式。"""
    paragraph.Range.Font.Name = font
    paragraph.Range.Font.NameFarEast = font
    paragraph.Range.Font.Size = size
    paragraph.Alignment = alignment
    paragraph.LineSpacingRule = 0
    paragraph.SpaceBefore = 0
    paragraph.SpaceAfter = 0


def build_doc(app):
    """造文档，返回 (doc, {名称: 取段落的函数})。"""
    doc = app.Documents.Add()
    doc.Content.Text = ''
    rng = doc.Range(0, 0)
    rng.InsertAfter('文档总标题')
    rng.InsertParagraphAfter()

    _append_paragraph(doc, '正文第一段，用来占位。')
    _append_paragraph(doc, '正文第二段，用来占位。')

    # 普通多单元格表格（表格格式由"变更图片与表格的格式"负责，不应被改字号）
    normal_table = _append_table(doc, 2, 2)
    normal_table.Range.Text = 'A1\x07B1\x07A2\x07B2\x07'

    # 代码框一：单单元格表格，作者写了直接格式（楷体、9磅、右对齐、单倍行距）
    box_a = _append_table(doc, 1, 1)
    box_a.Cell(1, 1).Range.Text = 'code box with own format'
    _set_direct_format(box_a.Cell(1, 1).Range.Paragraphs(1), '楷体', 9, 2)

    # 代码框二：单单元格表格，没有任何直接格式（全部继承「正文」样式），
    # 用来验证"字号跟随正文设置"、其余格式仍保持处理前的样子
    box_b = _append_table(doc, 1, 1)
    box_b.Cell(1, 1).Range.Text = 'code box inheriting style'

    _append_paragraph(doc, '正文第三段，用来占位。')

    # 图文框：框内段落 + 一个空段落（空段落不能被"删除空行"删掉）
    frame = doc.Frames.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1))
    frame.Range.Text = '图文框内的段落'
    frame.Range.InsertParagraphAfter()

    # 文本框：作者写了直接格式（黑体、8磅、右对齐、单倍行距）
    text_box = doc.Shapes.AddTextbox(1, 60, 60, 240, 90)
    text_box.Name = TEXT_BOX_NAME
    text_box.TextFrame.TextRange.Text = '文本框里的段落'
    _set_direct_format(text_box.TextFrame.TextRange.Paragraphs(1), '黑体', 8, 2)

    getters = {
        '正文第二段': lambda: _find_paragraph(doc, '正文第二段'),
        '普通表格单元格': lambda: doc.Tables(1).Cell(1, 1).Range.Paragraphs(1),
        '代码框一': lambda: doc.Tables(2).Cell(1, 1).Range.Paragraphs(1),
        '代码框二': lambda: doc.Tables(3).Cell(1, 1).Range.Paragraphs(1),
        '图文框内段落': lambda: doc.Frames(1).Range.Paragraphs(1),
        '文本框内段落': lambda: _find_shape(doc, TEXT_BOX_NAME)
                                    .TextFrame.TextRange.Paragraphs(1),
    }
    return doc, getters


def _find_paragraph(doc, text):
    """按文字找正文段落。

    不按下标取：处理后段落数可能变化（删除空行等），下标会对不上；
    也不缓存段落对象：跨多次文档操作后旧对象可能已失效。
    """
    for index in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(index)
        if para.Range.Text.startswith(text) and not get_in_table(para.Range):
            return para
    raise AssertionError('文档里找不到段落：%s' % text)


def get_in_table(range_obj):
    try:
        return bool(range_obj.Information(wc.wdWithInTable))
    except Exception:
        return False


def _find_shape(doc, name):
    """按名称找形状（文本框）。图文框也可能出现在 Shapes 里，故不能按下标取。"""
    for index in range(1, doc.Shapes.Count + 1):
        if doc.Shapes(index).Name == name:
            return doc.Shapes(index)
    raise AssertionError('文档里找不到形状：%s' % name)


def read_format(paragraph):
    """读取一个段落的完整格式（逐属性读，读不到记 ERR）。"""
    values = {}
    try:
        pf = paragraph.Range.ParagraphFormat
        font = paragraph.Range.Font
    except Exception as ex:
        return {'读不到格式': str(ex)}
    for attr in PARAGRAPH_ATTRS:
        try:
            values[attr] = getattr(pf, attr)
        except Exception as ex:
            values[attr] = 'ERR:%s' % ex
    for attr in FONT_ATTRS:
        try:
            values['Font.' + attr] = getattr(font, attr)
        except Exception as ex:
            values['Font.' + attr] = 'ERR:%s' % ex
    try:
        values['Style'] = paragraph.Range.Style.NameLocal
    except Exception as ex:
        values['Style'] = 'ERR:%s' % ex
    return values


def run_flow(fmt, progress_texts):
    """按 main_form._on_format_adjust 的顺序跑一遍流程。"""
    def report(message):
        progress_texts.append(message)

    snapshots = fmt.snapshot_protected_paragraph_format()
    fmt.set_content_format(CONTENT_INDENT, CONTENT_ALIGN, CONTENT_FONT,
                           CONTENT_SIZE_NAME, None, True, True,
                           progress=report)
    try:
        fmt.set_content_style(CONTENT_INDENT, CONTENT_ALIGN, CONTENT_FONT,
                              CONTENT_SIZE_NAME, progress=report)
    finally:
        fmt.restore_protected_paragraph_format(snapshots)

    fmt.set_images_and_tables(wrap_as_inline=True, no_indent=True,
                              max_width=True, progress=report)
    fmt.set_tables_auto_adjust_and_align(progress=report)
    fmt.set_table_paragraph_no_indent(progress=report)
    fmt.set_table_surrounding_spacing(progress=report)


def main():
    app = w.DispatchEx('Word.Application')
    app.Visible = False
    doc = None
    failed = []
    try:
        doc, getters = build_doc(app)

        for i in range(1, doc.Tables.Count + 1):
            table = doc.Tables(i)
            print('表格 %d：Rows=%s Columns=%s' % (
                i, table.Rows.Count, table.Columns.Count))
        print('图文框数：%d，框内段落数：%d' % (
            doc.Frames.Count, doc.Frames(1).Range.Paragraphs.Count))

        before = {name: read_format(get()) for name, get in getters.items()}
        frame_paras_before = doc.Frames(1).Range.Paragraphs.Count

        fmt = sdf.SetDocumentFormat(doc=_DocShim(doc, _AppShim(app)))
        progress_texts = []
        run_flow(fmt, progress_texts)

        # 处理结束后重新取段落对象：跨多次文档操作后旧对象可能已失效
        after = {name: read_format(getters[name]()) for name in getters}

        print('\n================ 比对结果 ================')
        for name in before:
            changed = [attr for attr in before[name]
                       if before[name][attr] != after[name][attr]]
            print('%-14s 变化的属性：%s' % (name, changed or '（无）'))
            for attr in changed:
                print('    %-30s %s -> %s'
                      % (attr, before[name][attr], after[name][attr]))

        print('\n================ 断言 ================')

        # 1) 文本框 / 图文框 / 单单元格表格：除字号外一律不变
        for name in ('文本框内段落', '图文框内段落', '代码框一', '代码框二'):
            changed = [attr for attr in before[name]
                       if attr != 'Font.Size'
                       and before[name][attr] != after[name][attr]]
            assert not changed, f'{name}的格式被改动：{changed}'
            print(f'  OK  {name}：除字号外的格式均未变')

        # 2) 代码框二 / 图文框内段落没有直接字号 → 跟随正文设置（小四=12磅）
        for name in ('代码框二', '图文框内段落'):
            size = after[name]['Font.Size']
            assert size == CONTENT_SIZE_POINTS, \
                f'{name}的字号应跟随正文设置（{CONTENT_SIZE_POINTS}磅），实际 {size}'
            print(f'  OK  {name}：字号跟随正文设置为 {size} 磅')

        # 3) 代码框一 / 文本框内的段落自带直接字号 → 保持原字号
        assert after['代码框一']['Font.Size'] == 9, after['代码框一']['Font.Size']
        assert after['文本框内段落']['Font.Size'] == 8, \
            after['文本框内段落']['Font.Size']
        print('  OK  代码框一 = 9 磅、文本框内段落 = 8 磅：自带直接字号保持不变')

        # 4) 图文框里的空段落还在（没被"删除空行"删掉）
        frame_paras_after = doc.Frames(1).Range.Paragraphs.Count
        assert frame_paras_after == frame_paras_before, \
            f'图文框内段落数变了：{frame_paras_before} -> {frame_paras_after}'
        print(f'  OK  图文框内段落数不变（{frame_paras_after} 段，空段落未被删除）')

        # 5) 正文段落照常被处理（功能没有因为上面的保护而失效）
        body = after['正文第二段']
        assert body['Font.Name'] == CONTENT_FONT, body['Font.Name']
        assert body['Font.Size'] == CONTENT_SIZE_POINTS, body['Font.Size']
        assert body['CharacterUnitFirstLineIndent'] == 2, body
        assert body['Alignment'] == 3, body['Alignment']      # 3 = 两端对齐
        print('  OK  正文段落：宋体、12磅、首行缩进2字符、两端对齐')

        # 6) 多单元格表格仍按"表格格式"处理：字体、字号不被正文设置改掉
        #    （对齐与段前段后由"图片与表格格式"那几步负责，本测试不比对）
        cell_before = before['普通表格单元格']
        cell = after['普通表格单元格']
        for attr in ('Font.Name', 'Font.NameFarEast', 'Font.Size',
                     'LineSpacingRule', 'LeftIndent', 'FirstLineIndent',
                     'CharacterUnitFirstLineIndent'):
            assert cell[attr] == cell_before[attr], \
                f'普通表格单元格的 {attr} 被改动：{cell_before[attr]} -> {cell[attr]}'
        print('  OK  普通表格单元格：字体 %s、字号 %s 未被正文设置覆盖'
              % (cell['Font.Name'], cell['Font.Size']))

        # 7) 流程确实跑完了（状态栏文字有上报）
        assert progress_texts, '流程没有上报任何进度'
        print('  OK  流程上报了 %d 条进度文字' % len(progress_texts))
    except AssertionError as ex:
        failed.append(str(ex))
        print('\n[失败] %s' % ex)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            app.Quit(SaveChanges=0)
        except Exception:
            pass

    if failed:
        print('\n自测失败：%d 项' % len(failed))
        return 1
    print('\n全部通过。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
