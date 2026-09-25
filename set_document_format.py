import os
import win32com.client
import word_constants as wc
from word_constants import (set_range_style, get_range_information, get_range_style,
                            get_list_gallery, apply_indent_style,
                            _INDENT_STYLE_MAP,
                            get_effective_font_names, get_effective_font_size,
                            get_effective_indents, apply_font_names,
                            apply_font_size, apply_indents,
                            apply_indents_pinned, clear_indents)


# 界面的"对齐方式"文本 → Word 段落对齐常量
_ALIGNMENT_MAP = {
    "左对齐": wc.wdAlignParagraphLeft,
    "右对齐": wc.wdAlignParagraphRight,
    "居中对齐": wc.wdAlignParagraphCenter,
    "两端对齐": wc.wdAlignParagraphJustify,
    "分散对齐": wc.wdAlignParagraphDistribute,
}

# 总标题（文档首段）的段间距：段前0磅、段后28磅。
# 用固定磅值而非"行"：总标题的字号可能被改成二号/三号，按行折算会让
# 段后间距随字号变化，得不到稳定的28磅。
_MAIN_TITLE_SPACE_BEFORE = 0.0
_MAIN_TITLE_SPACE_AFTER = 28.0


def _no_progress(_message):
    """默认进度回调：不做任何事（供未传 progress 的调用方使用）。"""
    return None


def _make_progress(callback, every=20):
    """把进度回调包一层节流：每 every 次调用只向状态栏发送一次。

    设置正文、图片、表格与标题都是逐段/逐个对象的 COM 调用，耗时较长。
    逐次刷新状态栏会让界面重绘占掉可观的时间，因此按固定步长抽样上报，
    既让用户看得到进展，又不至于拖慢处理。

    Args:
        callback: 形如 f(message) 的可调用对象；None 表示不需要上报。
        every: 每多少次调用上报一次（至少 1）。

    Returns:
        callable: f(message) —— 未到上报点时静默返回。
    """
    if callback is None:
        return _no_progress

    step = max(1, int(every))
    state = {"n": 0}

    def report(message):
        state["n"] += 1
        if state["n"] % step:
            return
        try:
            callback(message)
        except Exception:
            # 上报进度失败绝不影响格式化处理本身
            pass

    return report


# ========================================================
# 文本框 / 代码框（单单元格表格）识别
#
# 这两类容器里的段落，除字号外一律保持原样：字体、对齐、缩进、行距、
# 段前段后都不该被本工具改掉。原因是它们是"框"——框里的排版由作者自己
# 定（代码框常是等宽字体、不缩进、单倍行距），跟着正文走就会散架。
# ========================================================

def range_in_text_frame(range_obj):
    """Range 是否位于图文框（Frame）内。

    图文框里的段落就在 Document.Paragraphs 里（不像文本框那样自成一段
    故事），不单独识别就会被当成正文段落处理。
    """
    try:
        return range_obj.Frames.Count > 0
    except Exception:
        return False


def is_single_cell_table(table):
    """表格是否只有一个单元格——代码框最常见的做法。

    优先用 Range.Cells.Count 判"只有一个单元格"，取不到时退回"一行一列"：
    本机 Word 2013 上 Table.Cells 属性经 COM 调用会报错
    （AttributeError: <unknown>.Cells），Table.Range.Cells 正常。
    """
    try:
        return int(table.Range.Cells.Count) == 1
    except Exception:
        pass
    try:
        return int(table.Rows.Count) == 1 and int(table.Columns.Count) == 1
    except Exception:
        return False


def range_in_single_cell_table(range_obj):
    """Range 是否位于单单元格表格（代码框）内。"""
    try:
        if not get_range_information(range_obj, wc.wdWithInTable):
            return False
    except Exception:
        return False
    try:
        tables = range_obj.Tables
        if int(tables.Count) < 1:
            return False
        # Tables(1) 是最内层的表格：嵌套表格时按最内层的规模判断
        return is_single_cell_table(tables(1))
    except Exception:
        return False


def _iter_shapes(shapes):
    """递归展平形状集合：组合形状（Group）里的子形状也会被列出。

    文本框被"组合"后不出现在顶层 Shapes 里，不递归就会漏掉。
    """
    try:
        count = int(shapes.Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            shape = shapes(index)
        except Exception:
            continue
        try:
            is_group = int(shape.Type) == wc.msoGroup
        except Exception:
            is_group = False
        if is_group:
            try:
                for item in _iter_shapes(shape.GroupItems):
                    yield item
            except Exception:
                continue
        else:
            yield shape


def iter_text_frame_ranges(doc):
    """产出文档中所有文本框（形状 TextFrame）的文本 Range。

    形状很多时逐个读 TextFrame 会抛异常（图片、线条等本就没有文本框），
    取不到的直接跳过。
    """
    for shape in _iter_shapes(doc.Shapes):
        try:
            if not shape.TextFrame.HasText:
                continue
            yield shape.TextFrame.TextRange
        except Exception:
            continue


class SetDocumentFormat:
    def __init__(self, full_name=None, doc=None):
        self.work_doc = None
        self.level_list_templates = {}
        if doc is not None:
            self.work_doc = doc
            print(f"正在激活文档：{self.work_doc.FullName}")
            self.work_doc.Activate()
        elif full_name is not None:
            os.startfile(full_name)
            from universal import Universal
            self.work_doc = Universal().get_active_word_app().Documents[full_name]
            self.work_doc.Activate()
    
    def set_page_margins(self, top, bottom, left, right):
        if self.work_doc is None:
            print("没有找到目标Word文档。")
            return
        
        margin_points = 28.35
        right_margin_points = margin_points * right
        
        for i in range(1, self.work_doc.Sections.Count + 1):
            try:
                section = self.work_doc.Sections(i)
                ps = section.PageSetup
                print(f"Section {i}: RightMargin before = {ps.RightMargin:.2f}")
                
                ps.TopMargin = margin_points * top
                ps.BottomMargin = margin_points * bottom
                ps.LeftMargin = margin_points * left
                ps.RightMargin = right_margin_points
                ps.MirrorMargins = False
                ps.Gutter = 0
                ps.CharsPerLine = 0
                ps.LinesPage = 0
                ps.LayoutMode = 0
                
                print(f"Section {i}: RightMargin after = {ps.RightMargin:.2f} (expected {right_margin_points:.2f})")
            except Exception as ex:
                print(f"Section {i}: 设置边距失败 - {ex}")
        
        self.work_doc.Repaginate()
        
        try:
            self.work_doc.ActiveWindow.View.Type = 3
        except:
            pass
    
    def set_standard_line_spacing(self, progress=None):
        """把所有正文段落设为1.5倍行距、段前段后0磅。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            print("没有找到目标Word文档。")
            return

        report = _make_progress(progress, 20)
        paragraphs = self.work_doc.Paragraphs
        para_total = paragraphs.Count
        for i in range(1, para_total + 1):
            report(f"正在设置正文行距与段间距… {i}/{para_total}")
            # 首段为文档总标题，不属于正文：正文处理不改变它的行距与段间距
            # （总标题格式由 set_main_title_format 单独负责）
            if i == 1:
                continue

            paragraph = paragraphs(i)

            # 表格内的段落不按正文处理，保持表格格式中设置的单倍行距
            if get_range_information(paragraph.Range, wc.wdWithInTable):
                continue

            # 图文框（Frame）内的段落与文本框一样：除字号外一律保持原样。
            # 图文框的段落就在 Document.Paragraphs 里，不跳过就会被套上
            # 1.5倍行距、段前段后0磅。（文本框里的段落不在
            # Document.Paragraphs 里，这里碰不到它们。）
            if range_in_text_frame(paragraph.Range):
                continue

            is_title = False
            try:
                paragraph_style = get_range_style(paragraph.Range)
                if paragraph_style is not None:
                    style_name = paragraph_style.NameLocal
                    is_title = "标题" in style_name or "MainTitle" in style_name
            except:
                pass
            # 样式名不含"标题"但已设大纲级别（1~9）的段落同样是标题：
            # 判定标准与"变更各章节标题格式"保持一致，避免标题被当正文处理。
            if not is_title:
                is_title = self._has_heading_outline_level(paragraph)

            if not is_title:
                try:
                    try:
                        paragraph.LineSpacingRule = wc.wdLineSpace1pt5
                        paragraph.DisableLineHeightGrid = True
                    except:
                        try:
                            paragraph.Format.LineSpacingRule = wc.wdLineSpace1pt5
                            paragraph.Format.DisableLineHeightGrid = True
                        except:
                            pass
                    try:
                        paragraph.SpaceBefore = 0
                    except:
                        pass
                    try:
                        paragraph.SpaceAfter = 0
                    except:
                        pass
                except:
                    pass

    # ========================================================
    # 缩进状态判定 / 按需设置缩进
    #
    # 段落已经处于目标缩进状态时不再重复写缩进：重复写入不会改变显示，
    # 却会把同一组缩进属性反复写进段落（每次都触发 Word 重排），并且容易
    # 因"点值 / 字符单位值"两组属性不同步而把段落弄成自相矛盾的状态。
    # 判定用段落**当前有效值**（含从样式继承来的值），只要它已经等于目标
    # 缩进就跳过写入。
    # ========================================================

    def _get_current_indents(self, range_obj):
        """读取 Range 所在段落当前的缩进属性；读取失败时返回 None。

        返回 None 表示"状态未知"——此时调用方必须照常写入缩进，不能因为
        读不到就误判为"已经缩进"。
        """
        try:
            indents = get_effective_indents(range_obj)
        except Exception:
            return None
        return indents if indents else None

    def _indent_style_already_satisfied(self, indents, indent_style=None):
        """判断段落是否已经处于缩进状态（已设置缩进方式）。

        规则（对应"正文格式已经为缩进状态时，不要再叠加设置缩进"）：
        段落当前只要已经带有任何缩进——首行缩进、左缩进或右缩进——即视为
        "已为缩进状态"，不再对它写入缩进（含"无缩进"，不再把已有缩进清零）。

        Args:
            indents: _get_current_indents() 的返回值；None 表示状态未知，
                直接返回 False（照常设置缩进）。
            indent_style: 界面上的缩进方式文本。仅用于"该选项不在缩进方式
                表中"的情形——未识别时不动缩进，按已满足处理。

        Returns:
            bool: True 表示已有缩进（或缩进方式未识别），无需再设置缩进。
        """
        if indent_style is not None and indent_style not in _INDENT_STYLE_MAP:
            # 缩进方式未识别：不该动缩进，"已满足"按 True 处理以避免误写
            return True

        if not indents:
            return False

        def _value(attr):
            """返回属性当前值；未记录（继承或读不到）返回 None。"""
            return indents.get(attr)

        def _nonzero(attr):
            value = _value(attr)
            if value is None:
                return False
            return abs(value) > 0.01

        # 首行缩进（含字符单位值）
        if _nonzero('FirstLineIndent') or _nonzero('CharacterUnitFirstLineIndent'):
            return True

        # 左缩进（悬挂缩进在 Word 里表现为左缩进 > 0 且首行缩进 < 0）
        if _nonzero('LeftIndent') or _nonzero('CharacterUnitLeftIndent'):
            return True

        # 右缩进同样属于"已有缩进设置"
        if _nonzero('RightIndent') or _nonzero('CharacterUnitRightIndent'):
            return True

        return False

    def _apply_indent_style_if_needed(self, range_obj, indent_style):
        """按需为 Range 所在段落设置缩进方式。

        只用于**段落**（Range）：段落已经是缩进状态（已设置缩进方式）时
        直接返回，不做任何写入——不在已有缩进之上叠加设置缩进，也不把已有
        缩进清零。只有完全没有缩进的段落才按所选方式写入。

        样式（Styles 的 ParagraphFormat）不走这里：样式的缩进属性多是从
        基准样式继承来的（读取值可能是别的样式给的），据此判断"已缩进"
        会漏掉真正需要写入的缩进，因此样式一律照常显式设置。

        Returns:
            bool: True 表示本次确实写入了缩进；False 表示段落已有缩进
                （或状态未知而无需写入），未做任何修改。
        """
        if self._indent_style_already_satisfied(
                self._get_current_indents(range_obj), indent_style):
            return False
        apply_indent_style(range_obj.ParagraphFormat, indent_style)
        return True

    def _is_between_tables(self, index):
        """判断第 index 段是否夹在两张表格之间（删除它会让两张表合并）。

        Word 的规则是：两张表格之间没有段落间隔就会被合并成一张。删空行时
        必须避开这种空段落——代码框常用单单元格表格做成，被并进相邻表格后
        就再也不是"单单元格表格"，会被当成普通表格处理，与"框内除字号外
        不改动"的要求相悖（更糟的是文档结构本身被改坏了）。

        以"前一段与后一段都在表格里"为准；任一侧取不到（首段、末段、
        空文档）都按不在表格里处理——宁可多删一个空行，也不误判。
        """
        for side in (index - 1, index + 1):
            if side < 1:
                return False
            try:
                side_para = self.work_doc.Paragraphs(side)
            except Exception:
                return False
            if not get_range_information(side_para.Range, wc.wdWithInTable):
                return False
        return True

    def set_content_format(self, indent_style, alignment, font, font_size,
                           code_paragraphs=None, is_delete_empty_lines=False,
                           is_standard_line_spacing=False, progress=None):
        """按界面"变更正文格式"组的设置处理文档正文段落。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度；
                为 None（默认）时不上报，行为与加入该参数之前完全一致。
        """
        report = _make_progress(progress, 20)
        code_paragraph_indexes = []
        if code_paragraphs is not None:
            for row in code_paragraphs:
                start_idx = row['StartIndex']
                end_idx = row['EndIndex']
                for m in range(start_idx, end_idx + 1):
                    code_paragraph_indexes.append(m)
        
        word_alignment = _ALIGNMENT_MAP.get(alignment, wc.wdAlignParagraphLeft)
        
        paragraphs = self.work_doc.Paragraphs

        # 浮动图片（Shapes）的锚点段落也要按"图片段落"跳过。
        # 浮动图片不计入 InlineShapes，仅凭下面的 InlineShapes.Count 判断会漏判，
        # 带浮动图片的段落就会被当成正文套上首行缩进与对齐方式，
        # 图片随之被顶出缩进——这正是"未勾选图片格式，图片却被动了"的原因。
        shape_anchor_starts = set()
        try:
            for shape in self.work_doc.Shapes:
                try:
                    anchor = shape.Anchor
                except Exception:
                    continue
                if anchor is None:
                    continue
                try:
                    shape_anchor_starts.add(int(anchor.Start))
                except Exception:
                    pass
                # 锚点 Range 可能覆盖多段，取它所在的第一段起点更稳妥
                try:
                    shape_anchor_starts.add(int(anchor.Paragraphs(1).Range.Start))
                except Exception:
                    pass
        except Exception:
            pass
        
        para_total = paragraphs.Count
        for i in range(1, para_total + 1):
            # 段落级进度上报（每20段一次），长文档处理期间状态栏保持可见进展
            report(f"正在变更正文格式… {i}/{para_total}")
            paragraph = paragraphs(i)

            # 首段为文档总标题，不属于正文：正文格式不处理它
            # （总标题格式由 set_main_title_format 单独负责）
            if i == 1:
                continue

            if paragraph.Range.Start in code_paragraph_indexes:
                continue
            
            if paragraph.Range.InlineShapes.Count > 0 or paragraph.Range.Tables.Count > 0:
                continue

            # 浮动图片的锚点段落同样跳过（见上方 shape_anchor_starts 的说明）
            if paragraph.Range.Start in shape_anchor_starts:
                continue

            # 表格内的段落不按正文处理（表格及其段落格式由"变更图片与表格的格式"负责）
            if get_range_information(paragraph.Range, wc.wdWithInTable):
                continue

            # 图文框（Frame）内的段落与文本框一样：除字号外一律保持原样。
            # 图文框的段落就在 Document.Paragraphs 里，不跳过就会被当成正文
            # 段落套上对齐、缩进、字体与字号。（文本框里的段落不在
            # Document.Paragraphs 里，这里碰不到它们，但会顺着「正文」样式
            # 继承而改变，由 snapshot/restore_protected_paragraph_format 保护。）
            if range_in_text_frame(paragraph.Range):
                continue

            try:
                paragraph_style = get_range_style(paragraph.Range)
                style_name = paragraph_style.NameLocal
                is_title = "标题" in style_name or "MainTitle" in style_name
            except:
                is_title = False

            # 样式名不含"标题"但已设大纲级别（1~9）的段落同样是标题。
            # 否则这类标题会被当作正文写入直接字号（如"五号"），
            # 而直接格式优先于样式，会覆盖后续为标题套用的"标题 N"样式字号。
            if not is_title:
                is_title = self._has_heading_outline_level(paragraph)

            # 标题段落整体跳过正文格式处理：对齐、缩进、字体、字号都不应被动。
            # 特别注意不能保留原来的 else 分支（FirstLineIndent = 0）——
            # 它在"缩进方式=不变更"或"是标题"时都会执行，会把标题首行缩进
            # 强制清零，这正是标题缩进被改掉的原因。
            if not is_title:
                paragraph.Alignment = word_alignment

                # 缩进处理：段落已经处于缩进状态（已设置缩进方式）时
                # **不再叠加设置缩进**，也不把已有缩进清零；
                # 只有完全没有缩进的段落才按所选方式写入。
                # "缩进方式=不变更"时保持原有行为：只把首行缩进点值清零。
                if indent_style is not None:
                    if self._apply_indent_style_if_needed(
                            paragraph.Range, indent_style):
                        paragraph.RightIndent = 0
                else:
                    paragraph.RightIndent = 0
                    paragraph.FirstLineIndent = 0

                if font is not None:
                    # 只给本段落写直接格式，**不要**改 paragraph_style.Font.Name。
                    # 段落所用样式（多为「正文」）是全文共用的：改样式字体
                    # 会连带改掉表格内等所有用该样式的段落——表格文字"被统一
                    # 成与正文相同"就出自这里。字号分支本来就是写 Range 直接
                    # 格式，字体分支与此保持一致。
                    try:
                        paragraph.Range.Font.Name = font
                        # 中文文档里汉字走"中文"字体槽位（NameFarEast），
                        # 只设 Name 改不到汉字，必须一并设置。
                        paragraph.Range.Font.NameFarEast = font
                    except Exception:
                        pass

                if font_size is not None and font_size != "不变更":
                    try:
                        paragraph.Range.Font.Size = self._convert_chinese_font_size_to_points(font_size)
                    except:
                        pass
        
        if is_delete_empty_lines:
            empty_total = self.work_doc.Paragraphs.Count
            for i in range(empty_total, 0, -1):
                # 段落级进度上报（每20段一次）：删除空行同样要遍历全文
                report(f"正在删除空行… {empty_total - i + 1}/{empty_total}")
                # 首段（总标题）不删除：删除空行属于正文处理，总标题另有专门处理
                if i == 1:
                    continue

                para = self.work_doc.Paragraphs(i)
                range_obj = para.Range
                
                if get_range_information(range_obj, wc.wdWithInTable):
                    continue

                # 图文框内的空段落不删除：删掉会改变框内的行数与排版，
                # 而框内（同文本框）除字号外不应有任何改动
                if range_in_text_frame(range_obj):
                    continue

                # 夹在两张表格之间的空段落不删除：删掉会被 Word 合并成一张表，
                # 单单元格的代码框就此失去"单单元格表格"的身份（见 _is_between_tables）
                if self._is_between_tables(i):
                    continue

                has_image = range_obj.InlineShapes.Count > 0
                if not has_image:
                    for shape in self.work_doc.Shapes:
                        if shape.Anchor is not None and shape.Anchor.InRange(range_obj):
                            has_image = True
                            break
                
                if has_image:
                    continue
                
                if not range_obj.Text.strip():
                    try:
                        range_obj.Delete()
                    except:
                        pass
        
        if is_standard_line_spacing:
            self.set_standard_line_spacing(progress=progress)

    def set_content_style(self, indent_style=None, alignment=None,
                          font=None, font_size=None, progress=None):
        """按界面"变更正文格式"组的设置，同步文档的「正文」样式本身。

        与 set_content_format 的分工：
          - set_content_format 只对文档中现有的正文段落写直接格式；
          - 本方法改的是样式本身（Word 内置「正文」/ Normal 样式），
            使之后新输入的段落、以及不带直接格式的段落也能与界面一致。

        传入 None 表示界面上选择了"不变更"，对应项保持样式原样不动。

        Args:
            indent_style: 缩进方式（"首行缩进2字符" / "无缩进" / "悬挂缩进"）
            alignment: 对齐方式（"两端对齐" 等）
            font: 字体名
            font_size: 中文字号名（"五号" 等）
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 1)
        report("正在同步「正文」样式…")

        style = None
        for name in ("正文", "Normal", "normal"):
            try:
                style = self.work_doc.Styles(name)
                break
            except Exception:
                continue
        if style is None:
            return

        try:
            if font is not None:
                report(f"正在同步「正文」样式字体：{font}")
                style.Font.Name = font
                # 中文文档里汉字用的是"中文"字体槽位（NameFarEast），
                # 只设 Font.Name 改不到汉字，必须一并设置。
                try:
                    style.Font.NameFarEast = font
                except Exception:
                    pass

            if font_size is not None:
                report(f"正在同步「正文」样式字号：{font_size}")
                style.Font.Size = self._convert_chinese_font_size_to_points(font_size)

            if alignment is not None:
                report(f"正在同步「正文」样式对齐方式：{alignment}")
                style.ParagraphFormat.Alignment = _ALIGNMENT_MAP.get(
                    alignment, wc.wdAlignParagraphLeft)

            if indent_style is not None:
                report(f"正在同步「正文」样式缩进方式：{indent_style}")
                apply_indent_style(style.ParagraphFormat, indent_style)
        except Exception as ex:
            print(f"设置正文样式时出错: {ex}")

    # ========================================================
    # 受保护容器内段落的格式保护
    #
    # 受保护容器 = 表格（含只有一个单元格的"代码框"）+ 文本框 / 图文框。
    # 它们的共同点是"框里的排版由作者自己定"：容器内段落大多与正文共用
    # 「正文」样式，而 set_content_style 修改的是样式本身，因此会连带改掉
    # 容器内文字的字体、对齐与缩进——容器文字于是"被统一成与正文相同"。
    # 文本框 / 图文框里的文字虽然不参与正文段落的处理，却同样顺着样式
    # 继承被改掉。而框内的格式应当只由"变更图片与表格的格式"（表格）或
    # 作者自己（文本框/代码框）负责。
    #
    # 做法：改「正文」样式之前记录容器内段落的**有效**格式，改完以后以
    # 直接格式写回（直接格式优先于样式，故框内显示保持不变）。
    #
    # 唯一的例外是**字号**：文本框与单单元格表格（代码框）内的字号不写回，
    # 让它跟随「正文」样式与界面"变更正文格式"里的正文字号一起变。
    # ========================================================

    def _read_paragraph_format(self, paragraph):
        """读取段落的有效格式，返回 (fonts, size, indents, alignment)。"""
        try:
            fonts = get_effective_font_names(paragraph.Range)
        except Exception:
            fonts = {}
        size = get_effective_font_size(paragraph.Range)
        indents = get_effective_indents(paragraph.Range)
        try:
            alignment = paragraph.Alignment
        except Exception:
            alignment = None
        return (fonts, size, indents, alignment)

    def _write_paragraph_format(self, paragraph, values, restore_size=True):
        """以直接格式把快照的格式写回段落。

        restore_size=False 时不写回字号：字号是唯一允许作用于文本框 /
        单单元格表格（代码框）内段落的项，把旧字号写成直接格式反而会把
        它们冻在原字号上（直接格式优先于样式）。

        缩进用 apply_indents_pinned（点值最后写）：容器内段落"本来不缩进"
        的情况必须钉住，否则「正文」样式一改成"首行缩进2字符"，文本框内的
        段落就又缩进了。见 apply_indents_pinned 的说明。
        """
        fonts, size, indents, alignment = values
        apply_font_names(paragraph.Range, fonts)
        if restore_size:
            apply_font_size(paragraph.Range, size)
        apply_indents_pinned(paragraph.Range, indents)
        if alignment is not None:
            try:
                paragraph.Alignment = alignment
            except Exception:
                pass

    def _collect_range_paragraphs(self, range_obj):
        """把 Range 里的段落逐个取出来（取不到的跳过）。"""
        paragraphs = []
        try:
            count = int(range_obj.Paragraphs.Count)
        except Exception:
            return paragraphs
        for index in range(1, count + 1):
            try:
                paragraphs.append(range_obj.Paragraphs(index))
            except Exception:
                continue
        return paragraphs

    def _iter_container_paragraphs(self):
        """产出受保护容器里的段落：(容器类型, [段落, ...])。

        顺序固定：先文本框（含组合形状内的，按形状顺序），后图文框。
        快照与还原都调用本方法，两次顺序一致即可按位置一一对应——
        不能像表格内段落那样按 Range.Start 索引：文本框各自是独立的故事
        （story），框内段落的 Start 都从 0 开始，彼此会互相覆盖。

        注意：本方法是生成器，调用方**必须把一组段落当场用完再取下一组**，
        不能先把段落对象收集起来。原因见 snapshot_protected_paragraph_format
        的说明（访问 Document.Frames 会让已取得的文本框段落对象失效）。
        """
        for range_obj in iter_text_frame_ranges(self.work_doc):
            paragraphs = self._collect_range_paragraphs(range_obj)
            if paragraphs:
                yield ('text_box', paragraphs)

        try:
            frame_total = int(self.work_doc.Frames.Count)
        except Exception:
            frame_total = 0
        for index in range(1, frame_total + 1):
            try:
                frame_range = self.work_doc.Frames(index).Range
            except Exception:
                continue
            paragraphs = self._collect_range_paragraphs(frame_range)
            if paragraphs:
                yield ('frame', paragraphs)

    def snapshot_protected_paragraph_format(self):
        """快照受保护容器内段落的有效格式，供「正文」样式变更后还原。

        覆盖范围与"唯一例外是字号"的规则见本节开头的说明。

        Returns:
            dict: {
                'paragraphs': [(Range.Start, fonts, size, indents, alignment,
                                restore_size), ...],   # 表格内段落，文档顺序
                'containers': [(容器类型,
                                [(fonts, size, indents, alignment), ...]), ...],
            }
             fonts    — {字体槽位名: 字体名}
             size     — 字号磅值，取不到为 None
             indents  — {缩进属性名: 值}
             alignment— 段落对齐常量
             restore_size — 是否写回字号：单单元格表格（代码框）为 False，
                            其余表格为 True；containers 里的段落一律不写回。
        """
        if self.work_doc is None:
            return {}

        snapshots = {'paragraphs': [], 'containers': []}

        try:
            total = self.work_doc.Paragraphs.Count
        except Exception:
            total = 0

        # 表格内段落按**文档顺序**记成列表，还原时同样按顺序一一对应。
        # 不按 Range.Start 索引：快照与还原之间隔着 set_content_format，
        # 它可能删掉正文里的空行，其后所有段落的起点整体前移，按起点查找
        # 会全部落空——代码框（单单元格表格）的缩进就是这么又冒出来的。
        # 表格内段落的数量与顺序在这中间不会变，按顺序对应最稳。
        #
        # 这里只快照表格内段落。图文框内段落虽然也在 Document.Paragraphs 里，
        # 但由下面的 containers 负责（那里不写回字号），在这里一并快照的话
        # 会被当成表格段落、把字号写成直接格式。
        for index in range(1, total + 1):
            try:
                para = self.work_doc.Paragraphs(index)
                if not get_range_information(para.Range, wc.wdWithInTable):
                    continue
                key = int(para.Range.Start)
            except Exception:
                continue
            values = self._read_paragraph_format(para)
            restore_size = not range_in_single_cell_table(para.Range)
            snapshots['paragraphs'].append((key, values + (restore_size,)))

        # 每组段落当场读完再取下一组，不要先收集对象：访问 Document.Frames
        # 会让此前取得的文本框段落对象失效（本机 Word 2013 上再读它的属性
        # 会报"发生意外"，COM 错误码 E_OUTOFMEMORY），而 _iter_container_paragraphs
        # 在文本框之后就要碰 Frames。生成器 + 逐组处理正好避开这一点。
        for kind, paragraphs in self._iter_container_paragraphs():
            snapshots['containers'].append(
                (kind, [self._read_paragraph_format(p) for p in paragraphs]))

        return snapshots

    def restore_protected_paragraph_format(self, snapshots):
        """把快照的受保护容器内段落格式以直接格式写回。

        写回的项：字体、缩进、对齐。文本框与单单元格表格（代码框）内的
        字号不写回（见快照方法的说明），因此这些段落的字号始终跟随样式。

        表格内段落按文档顺序与快照一一对应（快照与还原之间可能删过空行，
        起点已整体前移，不能按 Range.Start 查找），文本框 / 图文框按容器
        顺序对应。

        Args:
            snapshots: snapshot_protected_paragraph_format() 的返回值；
                       为空时不做任何操作。
        """
        if self.work_doc is None or not snapshots:
            return

        paragraph_snapshots = snapshots.get('paragraphs') or []
        if paragraph_snapshots:
            # 先按文档顺序把表格内段落的位置摸清：只记 (段落序号, 起点)，
            # 不持有段落对象——后面还要枚举图文框，持有对象有失效的风险
            current = []
            try:
                total = self.work_doc.Paragraphs.Count
            except Exception:
                total = 0
            for index in range(1, total + 1):
                try:
                    para = self.work_doc.Paragraphs(index)
                    if not get_range_information(para.Range, wc.wdWithInTable):
                        continue
                    current.append((index, int(para.Range.Start)))
                except Exception:
                    continue

            # 数量对不上说明文档在快照之后被别的操作动过，退回按起点匹配，
            # 宁可少写几个段落，也不要把格式写到别的段落上
            by_start = None
            if len(current) != len(paragraph_snapshots):
                by_start = {start: index for index, start in current}

            for position, saved in enumerate(paragraph_snapshots):
                start, values = saved
                fonts, size, indents, alignment, restore_size = values
                if by_start is None:
                    doc_index = current[position][0]
                else:
                    doc_index = by_start.get(start)
                if doc_index is None:
                    continue
                try:
                    para = self.work_doc.Paragraphs(doc_index)
                except Exception:
                    continue
                self._write_paragraph_format(
                    para, (fonts, size, indents, alignment), restore_size)

        saved_containers = snapshots.get('containers') or []
        if not saved_containers:
            return

        # 同快照：逐组取、逐组写，段落对象不跨组持有（见快照处的说明）
        for (kind, paragraphs), saved in zip(self._iter_container_paragraphs(),
                                             saved_containers):
            saved_kind, saved_values = saved
            # 容器类型对不上说明两次枚举的顺序已经不一致，宁可不写
            if kind != saved_kind:
                continue
            for para, values in zip(paragraphs, saved_values):
                self._write_paragraph_format(para, values, restore_size=False)

    def set_main_title_format(self, font, font_size, is_bold, progress=None):
        """设置文章首行总标题的格式。

        传入 None 表示界面该项为"不变更"：
          - font / font_size 为 None 时，先读取段落**处理前**的有效字体与
            字号，并以原值写入 MainTitle 样式。总标题的格式由新建的
            MainTitle 样式承载，而新建样式的字体字号会继承文档默认值
            （如宋体五号）；若不以原值写回，段落原有字体字号会被样式默认值
            覆盖——这正是"选了不变更、字体字号却仍被改掉"的原因。
          - 缩进方式固定为**无缩进**（界面上不提供总标题的缩进开关）：
            MainTitle 会继承「正文」样式，而「正文」的缩进（如首行缩进
            2字符）后续会被 set_content_style 改写，若不在样式与段落上
            都清零，总标题会莫名多出首行缩进。
          - 加粗、居中、段前0磅、段后28磅属于总标题的固有格式，
            无论字体字号是否为"不变更"都要执行。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 1)
        report("正在读取总标题原格式…")

        try:
            main_para = self.work_doc.Paragraphs(1)
        except Exception:
            return

        # ---- 处理前快照：仅对"不变更"的项取值 ----
        keep_fonts = get_effective_font_names(main_para.Range) if font is None else None
        keep_size = get_effective_font_size(main_para.Range) if font_size is None else None

        style = None
        try:
            style = self.work_doc.Styles("MainTitle")
        except:
            style = self.work_doc.Styles.Add("MainTitle", wc.wdStyleTypeParagraph)

        report(f"正在写入总标题样式（{'加粗' if is_bold else '不加粗'}）…")
        if font is not None:
            style.Font.Name = font
            # 中文文档里汉字走"中文"字体槽位（NameFarEast），一并设置
            try:
                style.Font.NameFarEast = font
            except Exception:
                pass
        elif keep_fonts:
            apply_font_names(style, keep_fonts)

        if font_size is not None:
            style.Font.Size = self._convert_chinese_font_size_to_points(font_size)
        else:
            apply_font_size(style, keep_size)

        style.Font.Bold = 1 if is_bold else 0
        style.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
        # 总标题的段间距固定为段前0磅、段后28磅
        # （不能写成"0.5行"：折算值随字号变化，得不到稳定的28磅）
        style.ParagraphFormat.SpaceBefore = _MAIN_TITLE_SPACE_BEFORE
        style.ParagraphFormat.SpaceAfter = _MAIN_TITLE_SPACE_AFTER

        # 缩进方式固定为"无缩进"：6 个缩进属性（左/右/首行 + 三个字符单位
        # 版本）成对清零。只清首行缩进不够——"悬挂缩进"在 Word 里是
        # LeftIndent>0 且 FirstLineIndent<0，而中文的"缩进2字符"存在字符
        # 单位属性里。先清样式：MainTitle 基于「正文」，样式上的 0 才能挡住
        # 后续 set_content_style 给「正文」写进去的缩进。
        clear_indents(style)

        set_range_style(main_para.Range, style)

        report("正在应用总标题格式（居中、无缩进、段后28磅）…")
        main_para.Alignment = wc.wdAlignParagraphCenter
        # 段落原有直接格式优先于样式：缩进与段间距再以直接格式写一遍，
        # 否则首段自己的首行缩进/段前段后仍会保持旧值。
        clear_indents(main_para.Range)
        try:
            main_para.SpaceBefore = _MAIN_TITLE_SPACE_BEFORE
            main_para.SpaceAfter = _MAIN_TITLE_SPACE_AFTER
        except Exception:
            pass

    def set_title_styles(self, level_or_para, font=None, font_size=None,
                         number_style=None, indent_style=None, progress=None):
        """设置某一级别标题的样式（level_or_para 为级别整数时）。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return None

        report = _make_progress(progress, 1)

        if isinstance(level_or_para, int):
            level = level_or_para
            report(f"正在读取「标题 {level}」样式与目录行距…")
            self.set_toc_line_spacing()
            style = None
            try:
                style = self.work_doc.Styles(f"标题 {level}")
            except:
                style = self.work_doc.Styles.Add(f"标题{level}", wc.wdStyleTypeParagraph)

            if font is not None:
                style.Font.Name = font

            if font_size is not None:
                style.Font.Size = self._convert_chinese_font_size_to_points(font_size)

            # 链接列表模板前先记下样式自身的字体。LinkToListTemplate 会把
            # 列表模板级别的字体（新建模板默认即文档默认字体，如宋体五号）
            # 带到样式上，静默改掉"标题 N"样式的字体字号。用户选了"不变更"
            # 的项必须在链接之后还原回去。
            style_font_backup = self._snapshot_style_font(style)

            if number_style is not None:
                if number_style == "无序号":
                    report(f"正在解除「标题 {level}」的编号链接…")
                    self._remove_numbering_from_style(style, level)
                    self.level_list_templates.pop(level, None)
                else:
                    report(f"正在生成「标题 {level}」序号样式：{number_style}")
                    list_template = self._create_single_level_template(number_style, level)
                    if list_template is not None:
                        self.level_list_templates[level] = list_template
                        # 用 LinkToListTemplate 显式覆盖样式已有的列表模板链接
                        # （中文 Word 内置"标题 N"样式预链接了多级列表模板，
                        #  仅设置 template.ListLevels(level).LinkedStyle 无法覆盖）
                        try:
                            style.LinkToListTemplate(
                                ListTemplate=list_template, ListLevelNumber=level)
                        except Exception:
                            list_template.ListLevels(level).LinkedStyle = style

            self._restore_style_font(style, style_font_backup, font, font_size)

            # 缩进方式：写在链接列表模板之后，避免被列表级别的
            # NumberPosition / TextPosition 覆盖。界面上选"不变更"时
            # 调用方传 None，此处不做任何修改。
            if indent_style is not None:
                report(f"正在设置「标题 {level}」缩进方式：{indent_style}")
                apply_indent_style(style.ParagraphFormat, indent_style)

            report(f"正在设置「标题 {level}」加粗与段间距…")
            style.Font.Bold = 1
            style.ParagraphFormat.SpaceBefore = self.work_doc.Application.LinesToPoints(0.5)
            style.ParagraphFormat.SpaceAfter = self.work_doc.Application.LinesToPoints(0.5)
            return style
        else:
            para = level_or_para
            if font is None and font_size is None and number_style is None and indent_style is None:
                return

            try:
                style = get_range_style(para)

                if font is not None:
                    style.Font.Name = font

                if font_size is not None:
                    style.Font.Size = self._convert_chinese_font_size_to_points(font_size)

                if number_style is not None:
                    level = self.get_paragraph_outline_level(para)
                    if level >= 1:
                        if number_style == "无序号":
                            self._remove_numbering_from_style(style, level)
                            self.level_list_templates.pop(level, None)
                        else:
                            list_template = self._create_single_level_template(number_style, level)
                            if list_template is not None:
                                self.level_list_templates[level] = list_template
                                try:
                                    style.LinkToListTemplate(
                                        ListTemplate=list_template, ListLevelNumber=level)
                                except Exception:
                                    list_template.ListLevels(level).LinkedStyle = style

                style.Font.Bold = 1
                set_range_style(para.Range, style)
            except Exception as ex:
                print(f"设置标题样式时出错: {ex}")

    def _snapshot_style_font(self, style):
        """记录样式自身的字体属性，供链接列表模板之后还原。

        返回 {属性名: 值} 字典；取不到的属性直接跳过。
        """
        backup = {}
        for attr in ('Size', 'Name', 'NameFarEast', 'NameAscii'):
            try:
                backup[attr] = getattr(style.Font, attr)
            except Exception:
                continue
        return backup

    def _restore_style_font(self, style, backup, font, font_size):
        """还原样式字体中用户未显式指定的部分。

        font / font_size 为 None 表示该级别选择了"不变更"，此时样式的
        字体名 / 字号不应被列表模板带偏；显式指定的项则保持用户设定值，
        不做还原。
        """
        if not backup:
            return

        skip = set()
        if font is not None:
            skip.update(('Name', 'NameFarEast', 'NameAscii'))
        if font_size is not None:
            skip.add('Size')

        for attr, value in backup.items():
            if attr in skip:
                continue
            if attr == 'Size':
                # wdUndefined(9999999) 等无效值不写回，避免破坏样式
                try:
                    if not 0 < float(value) < 1000:
                        continue
                except (TypeError, ValueError):
                    continue
            try:
                setattr(style.Font, attr, value)
            except Exception:
                pass

    def set_toc_line_spacing(self):
        """设置1~9级目录样式的段前、段后行距为0.5行。"""
        if self.work_doc is None:
            return

        half_line = self.work_doc.Application.LinesToPoints(0.5)
        for level in range(1, 10):
            style = self._get_toc_style(level)
            if style is None:
                continue
            try:
                style.ParagraphFormat.SpaceBefore = half_line
                style.ParagraphFormat.SpaceAfter = half_line
            except:
                pass

    def _get_toc_style(self, level):
        """按中文/英文内置名称获取目录样式，找不到返回None。"""
        for name in (f"目录 {level}", f"toc {level}", f"TOC {level}"):
            try:
                return self.work_doc.Styles(name)
            except:
                continue
        return None

    def _create_single_level_template(self, number_style, target_level):
        """创建一个全新的列表模板，仅配置指定级别的编号格式。

        先清空全部9级，再设置目标级别和第1级（作为 ApplyListTemplate
        无级别参数时 Word 默认使用第1级的安全兜底）。
        """
        list_template = self.work_doc.ListTemplates.Add(True)

        # --- 先清空全部9级，防止模板自带默认格式或 Word 自动填充 ---
        for lvl in range(1, 10):
            try:
                level_obj = list_template.ListLevels(lvl)
                level_obj.NumberFormat = ""
                level_obj.NumberStyle = wc.wdListNumberStyleNone
            except:
                pass

        # --- 确定 NumberFormat 与 NumberStyle ---
        # 统一使用 \". \"（点号+空格）作为序号分隔符，
        # 将空格纳入 NumberFormat 字符串，配合 TrailingCharacter=wdTrailingNone
        # 使分隔符后的间距为0字符。
        if number_style == "一.":
            fmt = f"%{target_level}. "
            ns = wc.wdListNumberStyleSimpChinNum
        elif number_style == "一）":
            fmt = f"%{target_level}）"
            ns = wc.wdListNumberStyleSimpChinNum
        elif number_style == "1.":
            fmt = f"%{target_level}. "
            ns = wc.wdListNumberStyleArabic
        elif number_style == "第1章":
            # 章序号（一级、二级标题用）：如"第1章 绪论"、"第2章 需求分析"
            fmt = f"第%{target_level}章 "
            ns = wc.wdListNumberStyleArabic
        elif number_style == "步骤1. ":
            # 步骤序号（一级、二级标题用）：如"步骤1. 绪论"、"步骤2. 需求分析"
            fmt = f"步骤%{target_level}. "
            ns = wc.wdListNumberStyleArabic
        elif number_style == "①":
            fmt = f"%{target_level}"
            ns = wc.wdListNumberStyleNumberInCircle
        elif number_style == "资料1. ":
            fmt = f"资料%{target_level}. "
            ns = wc.wdListNumberStyleArabic
        else:
            return None

        # --- 同时设置目标级别和第1级 ---
        # 目标级别：供 ApplyListTemplateWithLevel 明确指定
        # 第1级：作为 ApplyListTemplate 无级别参数时 Word 默认使用第1级的安全兜底
        for lvl in (1, target_level):
            level_obj = list_template.ListLevels(lvl)
            level_obj.NumberFormat = fmt
            level_obj.NumberStyle = ns
            # 从图库模板中获取通用布局参数
            try:
                list_gallery = get_list_gallery(self.work_doc.Application, wc.wdOutlineNumberGallery)
                ref_template = list_gallery.ListTemplates(1)
                ref_level_obj = ref_template.ListLevels(1)
                for attr in ('NumberPosition', 'Alignment',
                             'ResetOnHigher', 'StartAt'):
                    try:
                        setattr(level_obj, attr, getattr(ref_level_obj, attr))
                    except:
                        pass
            except:
                pass
            # 强制设置 TrailingCharacter 为无尾随字符，
            # 使序号后分隔符（\". \"）与段落文字之间的间距为0字符。
            try:
                level_obj.TrailingCharacter = wc.wdTrailingNone
            except:
                pass
            # 设置 TabPosition 为 0，消除序号与文字间的制表符间距。
            try:
                level_obj.TabPosition = 0
            except:
                pass

        return list_template

    def _remove_numbering_from_style(self, style, level):
        """移除指定样式的编号（实现"无序号"功能）。

        做法：创建全新的空白列表模板，清空其所有级别的 NumberFormat，
        再将目标级别链接到对应样式。仅影响目标级别，不波及其它标题级别。
        注意：仅设置样式层的链接可能不够——对于已应用该样式的段落，
        调用方还需对段落直接执行 ListFormat.RemoveNumbers() 才能彻底清除。
        """
        try:
            empty_template = self.work_doc.ListTemplates.Add(True)
            # 清空所有9个级别的编号格式，防止模板自带默认格式
            for lvl in range(1, 10):
                try:
                    level_obj = empty_template.ListLevels(lvl)
                    level_obj.NumberFormat = ""
                    level_obj.NumberStyle = wc.wdListNumberStyleNone
                except:
                    pass
            # 用 LinkToListTemplate 显式覆盖样式已有的列表模板链接
            try:
                style.LinkToListTemplate(
                    ListTemplate=empty_template, ListLevelNumber=level)
            except Exception:
                empty_template.ListLevels(level).LinkedStyle = style
        except Exception as ex:
            print(f"移除编号时出错: {ex}")
    
    def set_images_and_tables(self, wrap_as_inline=True, no_indent=True,
                              max_width=True, table_only=False, progress=None):
        """设置图片与表格的格式。

        Args:
            wrap_as_inline: 是否将图片文字环绕类型设为嵌入型（默认True）
            no_indent: 是否取消图片/表格所在段落的缩进（默认True）。
                "不缩进"的最终保证不靠这里——本方法执行时正文样式尚未变动，
                清掉的缩进可能随后被「正文」样式的缩进带回来，表格内的段落
                还会被表格格式快照写成直接格式。流程末尾必须再调一次
                set_image_paragraph_no_indent() 与
                set_table_paragraph_no_indent()。
            max_width: 兼容保留参数。表格宽度已统一由 _auto_adjust_tables()
                处理（先“根据内容调整表格”，再撑满页面宽度），不再受该参数影响
            table_only: 是否只处理表格，跳过所有图片相关的处理（默认False）
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 5)

        # 计算页面正文区域宽度（用于max_width）
        page_width = None
        if max_width and not table_only:
            try:
                ps = self.work_doc.PageSetup
                page_width = ps.PageWidth - ps.LeftMargin - ps.RightMargin
            except:
                pass

        if not table_only:
            if wrap_as_inline:
                # 将所有浮动图形（Shapes）转换为嵌入型（InlineShapes）
                # 从后往前遍历，因为ConvertToInlineShape会从Shapes集合中移除当前项
                # Count 随转换递减，处理完一项后计数会变化，故不显示总数
                try:
                    shape_count = self.work_doc.Shapes.Count
                except Exception:
                    shape_count = 0
                for i in range(self.work_doc.Shapes.Count, 0, -1):
                    report(f"正在将浮动图片转为嵌入型… 剩余 {i} 个")
                    try:
                        shape = self.work_doc.Shapes(i)
                        shape.ConvertToInlineShape()
                    except:
                        pass
                if shape_count:
                    report(f"浮动图片转换完成，共处理 {shape_count} 个")

            # 处理嵌入型图片：最大宽度、不缩进
            try:
                inline_total = self.work_doc.InlineShapes.Count
            except Exception:
                inline_total = 0
            for index, inline_shape in enumerate(self.work_doc.InlineShapes, start=1):
                report(f"正在处理图片宽度与缩进… {index}/{inline_total}")
                if inline_shape.Type == wc.wdInlineShapePicture:
                    try:
                        if max_width and page_width is not None:
                            inline_shape.Width = page_width

                        # 图片所在段落一律不缩进：这里清的是段落缩进（点值 +
                        # 字符单位值全部归零，见 clear_indents）。只清
                        # FirstLineIndent 会在"悬挂缩进"下留下左缩进，
                        # 图片仍被顶偏。
                        if no_indent:
                            clear_indents(inline_shape.Range)
                    except:
                        pass

        # 处理表格：嵌入型、列宽按内容自适应后撑满页面宽度、不缩进
        try:
            table_total = self.work_doc.Tables.Count
        except Exception:
            table_total = 0
        for table_index, table in enumerate(self.work_doc.Tables, start=1):
            report(f"正在处理表格环绕与缩进… {table_index}/{table_total}")
            try:
                if wrap_as_inline:
                    # 确保所有表格为嵌入型（无文字环绕）
                    table.Rows.WrapAroundText = False

                if no_indent:
                    # 单单元格表格（代码框）例外：框内缩进是作者自己定的，
                    # 与文本框一样只允许改字号，不参与"表格文字一律不缩进"
                    if not is_single_cell_table(table):
                        # 表格内文字一律不缩进：清零的是全部 6 个缩进属性
                        # （点值 + 字符单位值），只把 FirstLineIndent 置 0 不够——
                        # 中文 Word 的"首行缩进2字符"存在字符单位属性里，点值清零
                        # 后字符单位值仍是 2，单元格文字照旧缩进，并且会被后续的
                        # 表格格式快照当成"原有缩进"写回成直接格式（见 clear_indents）。
                        clear_indents(table.Range)
            except:
                pass

        # 表格宽度：先按内容调整，再撑满页面宽度（统一由 _auto_adjust_tables 负责）
        self._auto_adjust_tables(progress=progress)

        # 表格内段落：单倍行距，段前、段后各空0.5行
        self.set_table_paragraph_spacing(progress=progress)

        # 注意：表格与上下段落的18磅间隔由 set_table_surrounding_spacing() 负责，
        # 它必须在正文格式处理之后调用，否则会被"标准行段间距"清零。

    def set_image_paragraph_no_indent(self, progress=None):
        """把文档中所有图片所在段落的缩进清零（图片段落一律不缩进）。

        为什么需要单独这一步、且必须放在流程末尾：

        1）时机问题。正文处理会改「正文」样式本身（set_content_style），
           而图片段落大多正是「正文」样式。样式一旦被设成"首行缩进2字符"，
           图片段落便会顺着样式继承出缩进——图片于是又被顶偏。
           （"变更正文格式"已按流程要求前置为第一步执行，正文样式因此
           总是先于图片/表格处理被改写。）
        2）直接格式可能"存不下来"。在段落上写"缩进=0"时，如果该值与样式
           当时的值相同，Word 会把这个直接格式省掉（等同于没写），之后
           样式一改，段落就跟着变。

        所以图片的缩进必须在正文/标题样式都处理完之后，再用直接格式钉一次。
        调用方（main_form、event_handlers）在两个处理流程的末尾调用本方法。

        覆盖对象：
            - 嵌入型图片所在段落（InlineShapes，含表格内的图片）；
            - 浮动图片的锚点段落（未转嵌入型时同样要清）；
            - 图片与文字同处一段的段落，整体按"图片段落"处理。
        清零的是全部 6 个缩进属性（点值 + 字符单位值），见 clear_indents。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。

        Returns:
            set: 已处理的段落起始位置集合
        """
        if self.work_doc is None:
            return set()

        report = _make_progress(progress, 5)
        handled = set()

        def _clear_paragraph(paragraph):
            try:
                key = int(paragraph.Range.Start)
            except Exception:
                return
            if key in handled:
                return
            handled.add(key)
            clear_indents(paragraph.Range)

        # 嵌入型图片（含表格内的图片）
        try:
            inline_total = self.work_doc.InlineShapes.Count
        except Exception:
            inline_total = 0
        try:
            for index, inline_shape in enumerate(self.work_doc.InlineShapes, start=1):
                report(f"正在取消图片段落缩进… {index}/{inline_total}")
                try:
                    if inline_shape.Type != wc.wdInlineShapePicture:
                        continue
                    _clear_paragraph(inline_shape.Range.Paragraphs(1))
                except Exception:
                    continue
        except Exception:
            pass

        # 浮动图片的锚点段落
        try:
            shape_total = self.work_doc.Shapes.Count
        except Exception:
            shape_total = 0
        try:
            for index, shape in enumerate(self.work_doc.Shapes, start=1):
                report(f"正在取消浮动图片段落缩进… {index}/{shape_total}")
                try:
                    anchor = shape.Anchor
                    if anchor is None:
                        continue
                    _clear_paragraph(anchor.Paragraphs(1))
                except Exception:
                    continue
        except Exception:
            pass

        return handled

    def set_table_paragraph_no_indent(self, progress=None):
        """把表格内所有单元格段落的缩进清零（表格文字一律不缩进）。

        为什么需要单独这一步、且必须放在流程末尾：

        1）单元格段落大多与正文共用「正文」样式。set_content_style 改的是
           样式本身，样式里的"首行缩进2字符"会顺着继承落到单元格文字上
           ——"处理表格时单元格文字被加上段首缩进"就出自这里。
        2）restore_protected_paragraph_format() 写回的快照取自「正文」样式
           尚未变动的时刻，快照里的缩进本就是当时从样式继承来的值。不在这里
           清零的话，继承来的缩进会被写成直接格式钉死在单元格上，此后即使
           把正文改成"无缩进"，单元格文字仍然缩进。

        因此表格与图片同理：在正文/标题样式处理完之后，用直接格式把缩进
        清零一次。清零的是全部 6 个缩进属性（点值 + 字符单位值），
        与 set_images_and_tables 里对表格的处理保持一致，见 clear_indents。

        单单元格表格（代码框）不在此列：框内缩进是作者自己定的，与文本框
        一样只允许改字号。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。

        Returns:
            set: 已处理的段落起始位置集合
        """
        if self.work_doc is None:
            return set()

        report = _make_progress(progress, 5)
        handled = set()

        try:
            table_total = self.work_doc.Tables.Count
        except Exception:
            table_total = 0
        table_index = 0
        for table in self.work_doc.Tables:
            table_index += 1
            report(f"正在取消表格内段落缩进… {table_index}/{table_total}")
            if is_single_cell_table(table):
                continue
            try:
                paragraphs = table.Range.Paragraphs
                count = paragraphs.Count
            except Exception:
                continue

            for i in range(1, count + 1):
                try:
                    paragraph = paragraphs(i)
                    key = int(paragraph.Range.Start)
                except Exception:
                    continue
                if key in handled:
                    continue
                handled.add(key)
                clear_indents(paragraph.Range)

        return handled

    def set_table_paragraph_spacing(self, progress=None):
        """设置表格内段落为单倍行距，段前、段后各空0.5行。

        单单元格表格（代码框）不在此列：框内的行距与段间距是作者自己定的，
        与文本框一样只允许改字号——代码框被塞进"段前段后0.5行"就会散架。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 1)
        half_line = self.work_doc.Application.LinesToPoints(0.5)
        try:
            table_total = self.work_doc.Tables.Count
        except Exception:
            table_total = 0
        table_index = 0
        for table in self.work_doc.Tables:
            table_index += 1
            report(f"正在设置表格内段落行距与间距… {table_index}/{table_total}")
            if is_single_cell_table(table):
                continue
            try:
                paragraphs = table.Range.Paragraphs
                for i in range(1, paragraphs.Count + 1):
                    paragraph = paragraphs(i)
                    try:
                        paragraph.LineSpacingRule = wc.wdLineSpaceSingle
                        paragraph.DisableLineHeightGrid = True
                    except:
                        try:
                            paragraph.Format.LineSpacingRule = wc.wdLineSpaceSingle
                            paragraph.Format.DisableLineHeightGrid = True
                        except:
                            pass
                    try:
                        paragraph.SpaceBefore = half_line
                    except:
                        pass
                    try:
                        paragraph.SpaceAfter = half_line
                    except:
                        pass
            except:
                pass

    def set_table_surrounding_spacing(self, points=18.0, progress=None):
        """设置表格与其上下段落之间的间距（默认18磅）。

        表格对象本身不提供外部间距属性，改为把表格前一段落的段后间距、
        表格后一段落的段前间距设为指定磅值。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 1)
        try:
            content_end = self.work_doc.Content.End
        except:
            return

        try:
            table_total = self.work_doc.Tables.Count
        except Exception:
            table_total = 0
        table_index = 0
        for table in self.work_doc.Tables:
            table_index += 1
            report(f"正在设置表格与上下段落的间距… {table_index}/{table_total}")
            try:
                table_start = table.Range.Start
                table_end = table.Range.End
            except:
                continue

            # 表格前一段落：表格起始位置前一个字符（即前一段落的段落标记）所属段落
            if table_start > 0:
                try:
                    before_para = self.work_doc.Range(
                        table_start - 1, table_start).Paragraphs(1)
                    # 首段（总标题）的段后间距由 set_main_title_format 固定为
                    # 28磅：表格紧接总标题时不能被这里的18磅改掉。
                    if int(before_para.Range.Start) > \
                            int(self.work_doc.Paragraphs(1).Range.Start):
                        before_para.SpaceAfter = points
                except:
                    pass

            # 表格后一段落：表格结束位置起的段落
            if table_end < content_end:
                try:
                    after_para = self.work_doc.Range(
                        table_end, table_end + 1).Paragraphs(1)
                    # 同上：紧接表格的是首段（总标题）时，它的段前间距固定为
                    # 0磅，不能被这里的18磅改掉。
                    if int(after_para.Range.Start) != \
                            int(self.work_doc.Paragraphs(1).Range.Start):
                        after_para.SpaceBefore = points
                except:
                    pass

    def center_all_images(self, progress=None):
        """将所有图片居中显示（始终执行，不受界面控件影响）

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 5)

        # 处理嵌入式图片：将所在段落设置为居中对齐
        try:
            inline_total = self.work_doc.InlineShapes.Count
        except Exception:
            inline_total = 0
        inline_index = 0
        for inline_shape in self.work_doc.InlineShapes:
            inline_index += 1
            if inline_shape.Type == wc.wdInlineShapePicture:
                report(f"正在居中图片… {inline_index}/{inline_total}")
                try:
                    inline_shape.Range.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
                except:
                    pass

        # 处理浮动图片：水平居中于页面
        try:
            shape_total = self.work_doc.Shapes.Count
        except Exception:
            shape_total = 0
        shape_index = 0
        for shape in self.work_doc.Shapes:
            shape_index += 1
            report(f"正在居中浮动图片… {shape_index}/{shape_total}")
            try:
                shape.RelativeHorizontalPosition = wc.wdRelativeHorizontalPositionPage
                shape.Left = wc.wdShapeCenter
            except:
                pass

    def _auto_adjust_tables(self, progress=None):
        """将所有表格按“根据内容调整表格”自动调整，并撑满页面宽度。

        分两步执行（等价于WPS中先点“自动调整——根据内容调整表格”，
        再让表格占满页面正文宽度）：
            1. AutoFitBehavior(wdAutoFitContent)：列宽按单元格内容自动分配；
            2. AutoFitBehavior(wdAutoFitWindow)：在各列内容比例的基础上，
               把整个表格拉伸到页面正文宽度。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 1)
        try:
            table_total = self.work_doc.Tables.Count
        except Exception:
            table_total = 0
        table_index = 0
        for table in self.work_doc.Tables:
            table_index += 1
            report(f"正在调整表格宽度（按内容自适应并撑满页面）… {table_index}/{table_total}")
            try:
                # 保证表格处于自动调整状态，使列宽跟随内容
                table.AllowAutoFit = True
            except:
                pass

            # 第1步：根据内容调整表格（列宽随内容自适应）
            try:
                table.AutoFitBehavior(wc.wdAutoFitContent)
            except:
                pass

            # 第2步：在内容比例的基础上撑满页面宽度
            try:
                table.AutoFitBehavior(wc.wdAutoFitWindow)
            except:
                pass

    def _apply_table_alignment(self, progress=None):
        """设置表格内容对齐：所有行文字垂直居中，首行文字水平居中。

        - 垂直方向：整表所有单元格垂直居中（保持各单元格原有的水平对齐方式）；
        - 水平方向：仅第一行（表头行）的单元格文字水平居中，其余行保持原样。

        单单元格表格（代码框）只做垂直居中，不做首行水平居中：它那唯一的
        一行就是正文本身，居中会改掉框内段落的对齐（除字号外不该动）。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 1)
        try:
            table_total = self.work_doc.Tables.Count
        except Exception:
            table_total = 0
        table_index = 0
        for table in self.work_doc.Tables:
            table_index += 1
            report(f"正在设置表格对齐方式（行垂直居中、首行水平居中）… {table_index}/{table_total}")
            # 所有行文字垂直方向居中显示
            try:
                table.Range.Cells.VerticalAlignment = wc.wdCellAlignVerticalCenter
            except:
                try:
                    table.Rows.VerticalAlignment = wc.wdCellAlignVerticalCenter
                except:
                    pass

            if is_single_cell_table(table):
                continue

            # 第一行水平方向也居中显示（逐单元格设置，跳过无内容的单元格）
            try:
                first_row = table.Rows(1)
                cell_count = first_row.Cells.Count
            except:
                continue

            for i in range(1, cell_count + 1):
                try:
                    cell_range = first_row.Cells(i).Range
                    # 空单元格没有段落可设置对齐，跳过以免报错
                    if not cell_range.Text.strip():
                        continue
                    cell_range.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
                except:
                    pass

    def set_tables_auto_adjust_and_align(self, progress=None):
        """表格统一格式（始终执行，不受界面控件影响）。

        依次完成：
            1. 根据内容调整表格，再撑满页面宽度（_auto_adjust_tables）；
            2. 所有行文字垂直居中、首行文字水平居中（_apply_table_alignment）。

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        self._auto_adjust_tables(progress=progress)
        self._apply_table_alignment(progress=progress)

    def set_all_tables_max_width(self):
        """将所有表格调整为页面最大宽度（始终执行，不受界面控件影响）。

        保留旧接口名以便兼容；实际表格宽度处理已包含“根据内容调整表格”，
        并同时完成行垂直居中与首行水平居中。
        """
        self.set_tables_auto_adjust_and_align()

    def add_image_border(self, progress=None):
        """为所有图片添加1px宽度的黑色外框（始终执行，不受界面控件影响）

        Args:
            progress: 可选回调 f(message)，用于向状态栏上报细粒度进度。
        """
        if self.work_doc is None:
            return

        report = _make_progress(progress, 5)
        border_edges = [wc.wdBorderTop, wc.wdBorderLeft,
                        wc.wdBorderBottom, wc.wdBorderRight]

        # 处理嵌入式图片：逐条边设置黑色细线边框
        try:
            inline_total = self.work_doc.InlineShapes.Count
        except Exception:
            inline_total = 0
        inline_index = 0
        for inline_shape in self.work_doc.InlineShapes:
            inline_index += 1
            if inline_shape.Type == wc.wdInlineShapePicture:
                report(f"正在为图片添加边框… {inline_index}/{inline_total}")
                try:
                    borders = inline_shape.Borders
                    borders.Enable = True
                    for edge in border_edges:
                        try:
                            b = borders(edge)
                            b.LineStyle = wc.wdLineStyleSingle
                            b.LineWidth = wc.wdLineWidth075pt
                            b.Color = wc.wdColorBlack
                        except:
                            pass
                except:
                    pass

        # 处理浮动图片：设置黑色细线边框
        try:
            shape_total = self.work_doc.Shapes.Count
        except Exception:
            shape_total = 0
        shape_index = 0
        for shape in self.work_doc.Shapes:
            shape_index += 1
            report(f"正在为浮动图片添加边框… {shape_index}/{shape_total}")
            try:
                shape.Line.Visible = True
                shape.Line.Weight = 1.0
                shape.Line.ForeColor.RGB = wc.wdColorBlack
            except:
                pass

    def delete_all_pictures(self):
        if self.work_doc is None:
            return
        
        for i in range(self.work_doc.InlineShapes.Count, 0, -1):
            if self.work_doc.InlineShapes(i).Type == wc.wdInlineShapePicture:
                self.work_doc.InlineShapes(i).Delete()
    
    def cancel_wrap_as_inline(self):
        if self.work_doc is None:
            return
        
        for shape in self.work_doc.Shapes:
            shape.WrapFormat.Type = wc.wdWrapSquare
    
    def get_paragraph_outline_level(self, target_paragraph):
        if target_paragraph is None:
            raise ValueError("目标段落对象不能为空")
        
        outline_level = target_paragraph.OutlineLevel
        return self._convert_outline_level_to_int(outline_level)
    
    def _convert_outline_level_to_int(self, outline_level):
        return outline_level

    def _has_heading_outline_level(self, target_paragraph):
        """判断段落的大纲级别是否为标题级别（1~9）。

        与 get_paragraph_outline_level() 同源，但用于"是否算标题"的判定：
        中文 Word 内置"标题 N"样式之外的段落（例如手动设置大纲级别的段落、
        样式名为英文 Heading N 或自定义名称的段落）同样应按标题对待。
        任何异常（含 WPS 下 OutlineLevel 取不到）都按非标题处理。
        """
        try:
            level = int(target_paragraph.OutlineLevel)
        except Exception:
            return False
        return 1 <= level <= 9
    
    def _convert_chinese_font_size_to_points(self, chinese_font_size):
        # 防御：如果传入 None，返回默认 12.0 磅
        if chinese_font_size is None:
            return 12.0

        size_mapping = {
            "初号": 42.0,
            "小初": 36.0,
            "一号": 26.0,
            "小一": 24.0,
            "二号": 22.0,
            "小二": 18.0,
            "三号": 16.0,
            "小三": 15.0,
            "四号": 14.0,
            "小四": 12.0,
            "五号": 10.5,
            "小五": 9.0,
            "六号": 7.5,
            "小六": 6.5,
            "七号": 5.5,
            "八号": 5.0
        }
        
        return size_mapping.get(chinese_font_size.strip(), 12.0)