import os
import win32com.client
import word_constants as wc
from word_constants import (set_range_style, get_range_information, get_range_style,
                            get_list_gallery, apply_indent_style,
                            get_effective_font_names, get_effective_font_size,
                            get_effective_indents, apply_font_names,
                            apply_font_size, apply_indents)


# 界面的"对齐方式"文本 → Word 段落对齐常量
_ALIGNMENT_MAP = {
    "左对齐": wc.wdAlignParagraphLeft,
    "右对齐": wc.wdAlignParagraphRight,
    "居中对齐": wc.wdAlignParagraphCenter,
    "两端对齐": wc.wdAlignParagraphJustify,
    "分散对齐": wc.wdAlignParagraphDistribute,
}


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
    
    def set_standard_line_spacing(self):
        if self.work_doc is None:
            print("没有找到目标Word文档。")
            return

        paragraphs = self.work_doc.Paragraphs
        for i in range(1, paragraphs.Count + 1):
            # 首段为文档总标题，不属于正文：正文处理不改变它的行距与段间距
            # （总标题格式由 set_main_title_format 单独负责）
            if i == 1:
                continue

            paragraph = paragraphs(i)

            # 表格内的段落不按正文处理，保持表格格式中设置的单倍行距
            if get_range_information(paragraph.Range, wc.wdWithInTable):
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
    
    def set_content_format(self, indent_style, alignment, font, font_size,
                           code_paragraphs=None, is_delete_empty_lines=False, is_standard_line_spacing=False):
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
        
        for i in range(1, paragraphs.Count + 1):
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
                paragraph.RightIndent = 0

                if indent_style is not None:
                    if indent_style == "首行缩进2字符":
                        paragraph.FirstLineIndent = 2 * 0.35 * 28.35
                    elif indent_style == "无缩进":
                        paragraph.FirstLineIndent = 0
                    elif indent_style == "悬挂缩进":
                        paragraph.FirstLineIndent = -2 * 0.35 * 28.35
                else:
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
            for i in range(self.work_doc.Paragraphs.Count, 0, -1):
                # 首段（总标题）不删除：删除空行属于正文处理，总标题另有专门处理
                if i == 1:
                    continue

                para = self.work_doc.Paragraphs(i)
                range_obj = para.Range
                
                if get_range_information(range_obj, wc.wdWithInTable):
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
            self.set_standard_line_spacing()

    def set_content_style(self, indent_style=None, alignment=None,
                          font=None, font_size=None):
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
        """
        if self.work_doc is None:
            return

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
                style.Font.Name = font
                # 中文文档里汉字用的是"中文"字体槽位（NameFarEast），
                # 只设 Font.Name 改不到汉字，必须一并设置。
                try:
                    style.Font.NameFarEast = font
                except Exception:
                    pass

            if font_size is not None:
                style.Font.Size = self._convert_chinese_font_size_to_points(font_size)

            if alignment is not None:
                style.ParagraphFormat.Alignment = _ALIGNMENT_MAP.get(
                    alignment, wc.wdAlignParagraphLeft)

            if indent_style is not None:
                apply_indent_style(style.ParagraphFormat, indent_style)
        except Exception as ex:
            print(f"设置正文样式时出错: {ex}")

    # ========================================================
    # 表格内段落格式保护
    #
    # 表格内段落大多与正文共用「正文」样式。set_content_style 修改的是
    # 样式本身，因此会连带改掉表格内文字的字号、对齐与缩进——表格文字
    # 于是"被统一成与正文相同"。而表格格式应当只由"变更图片与表格的
    # 格式"负责。
    #
    # 做法：改「正文」样式之前记录表格内段落的**有效**格式，改完以后以
    # 直接格式写回（直接格式优先于样式，故表格显示保持不变）。
    # ========================================================

    def snapshot_table_paragraph_format(self):
        """快照表格内段落的有效格式，供「正文」样式变更后还原。

        Returns:
            dict: {段落 Range.Start: (fonts, size, indents, alignment)}
                  fonts  — {字体槽位名: 字体名}
                  size   — 字号磅值，取不到为 None
                  indents— {缩进属性名: 值}
                  alignment — 段落对齐常量
        """
        if self.work_doc is None:
            return {}

        snapshots = {}
        try:
            total = self.work_doc.Paragraphs.Count
        except Exception:
            return snapshots

        for index in range(1, total + 1):
            try:
                para = self.work_doc.Paragraphs(index)
                if not get_range_information(para.Range, wc.wdWithInTable):
                    continue
                fonts = get_effective_font_names(para.Range)
                size = get_effective_font_size(para.Range)
                indents = get_effective_indents(para.Range)
                try:
                    alignment = para.Alignment
                except Exception:
                    alignment = None
                snapshots[int(para.Range.Start)] = (fonts, size, indents, alignment)
            except Exception:
                continue
        return snapshots

    def restore_table_paragraph_format(self, snapshots):
        """把快照的表格内段落格式以直接格式写回。

        Args:
            snapshots: snapshot_table_paragraph_format() 的返回值；
                       为空字典时不做任何操作。
        """
        if self.work_doc is None or not snapshots:
            return

        try:
            total = self.work_doc.Paragraphs.Count
        except Exception:
            return

        for index in range(1, total + 1):
            try:
                para = self.work_doc.Paragraphs(index)
                key = int(para.Range.Start)
            except Exception:
                continue

            if key not in snapshots:
                continue

            fonts, size, indents, alignment = snapshots[key]
            apply_font_names(para.Range, fonts)
            apply_font_size(para.Range, size)
            apply_indents(para.Range, indents)
            if alignment is not None:
                try:
                    para.Alignment = alignment
                except Exception:
                    pass

    def set_main_title_format(self, font, font_size, is_bold):
        """设置文章首行总标题的格式。

        传入 None 表示界面该项为"不变更"：
          - font / font_size 为 None 时，先读取段落**处理前**的有效字体与
            字号，并以原值写入 MainTitle 样式。总标题的格式由新建的
            MainTitle 样式承载，而新建样式的字体字号会继承文档默认值
            （如宋体五号）；若不以原值写回，段落原有字体字号会被样式默认值
            覆盖——这正是"选了不变更、字体字号却仍被改掉"的原因。
          - 段落缩进同样以原值写入样式并显式归零兜底：MainTitle 会继承
            「正文」样式，而「正文」的缩进（如首行缩进2字符）后续会被
            set_content_style 改写，不钉住的话总标题会莫名多出首行缩进。
          - 加粗、居中、段前段后 0.5 行属于总标题的固有格式，按参数执行。
        """
        if self.work_doc is None:
            return

        try:
            main_para = self.work_doc.Paragraphs(1)
        except Exception:
            return

        # ---- 处理前快照：仅对"不变更"的项取值 ----
        keep_fonts = get_effective_font_names(main_para.Range) if font is None else None
        keep_size = get_effective_font_size(main_para.Range) if font_size is None else None
        # 缩进无界面开关，一律按原值钉住（取不到时按无缩进处理）
        keep_indents = get_effective_indents(main_para.Range)

        style = None
        try:
            style = self.work_doc.Styles("MainTitle")
        except:
            style = self.work_doc.Styles.Add("MainTitle", wc.wdStyleTypeParagraph)

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
        style.ParagraphFormat.SpaceBefore = self.work_doc.Application.LinesToPoints(0.5)
        style.ParagraphFormat.SpaceAfter = self.work_doc.Application.LinesToPoints(0.5)

        # 钉住缩进：取到原值的按原值写，取不到的显式归零，
        # 避免总标题继承「正文」样式后续被改写的缩进
        for attr in ('LeftIndent', 'RightIndent', 'FirstLineIndent',
                     'CharacterUnitLeftIndent', 'CharacterUnitRightIndent',
                     'CharacterUnitFirstLineIndent'):
            try:
                setattr(style.ParagraphFormat, attr,
                        float(keep_indents.get(attr, 0.0)))
            except Exception:
                pass

        set_range_style(main_para.Range, style)

        main_para.Alignment = wc.wdAlignParagraphCenter
    
    def set_title_styles(self, level_or_para, font=None, font_size=None,
                         number_style=None, indent_style=None):
        if self.work_doc is None:
            return None

        if isinstance(level_or_para, int):
            level = level_or_para
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
                    self._remove_numbering_from_style(style, level)
                    self.level_list_templates.pop(level, None)
                else:
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
                apply_indent_style(style.ParagraphFormat, indent_style)

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
                              max_width=True, table_only=False):
        """设置图片与表格的格式。

        Args:
            wrap_as_inline: 是否将图片文字环绕类型设为嵌入型（默认True）
            no_indent: 是否取消图片/表格所在段落的缩进（默认True）
            max_width: 兼容保留参数。表格宽度已统一由 _auto_adjust_tables()
                处理（先“根据内容调整表格”，再撑满页面宽度），不再受该参数影响
            table_only: 是否只处理表格，跳过所有图片相关的处理（默认False）
        """
        if self.work_doc is None:
            return

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
                for i in range(self.work_doc.Shapes.Count, 0, -1):
                    try:
                        shape = self.work_doc.Shapes(i)
                        shape.ConvertToInlineShape()
                    except:
                        pass

            # 处理嵌入型图片：最大宽度、不缩进
            for inline_shape in self.work_doc.InlineShapes:
                if inline_shape.Type == wc.wdInlineShapePicture:
                    try:
                        if max_width and page_width is not None:
                            inline_shape.Width = page_width

                        if no_indent:
                            try:
                                inline_shape.Range.ParagraphFormat.FirstLineIndent = 0
                            except:
                                pass
                    except:
                        pass

        # 处理表格：嵌入型、列宽按内容自适应后撑满页面宽度、不缩进
        for table in self.work_doc.Tables:
            try:
                if wrap_as_inline:
                    # 确保所有表格为嵌入型（无文字环绕）
                    table.Rows.WrapAroundText = False

                if no_indent:
                    try:
                        table.Range.ParagraphFormat.FirstLineIndent = 0
                    except:
                        pass
            except:
                pass

        # 表格宽度：先按内容调整，再撑满页面宽度（统一由 _auto_adjust_tables 负责）
        self._auto_adjust_tables()

        # 表格内段落：单倍行距，段前、段后各空0.5行
        self.set_table_paragraph_spacing()

        # 注意：表格与上下段落的18磅间隔由 set_table_surrounding_spacing() 负责，
        # 它必须在正文格式处理之后调用，否则会被"标准行段间距"清零。

    def set_table_paragraph_spacing(self):
        """设置表格内段落为单倍行距，段前、段后各空0.5行。"""
        if self.work_doc is None:
            return

        half_line = self.work_doc.Application.LinesToPoints(0.5)
        for table in self.work_doc.Tables:
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

    def set_table_surrounding_spacing(self, points=18.0):
        """设置表格与其上下段落之间的间距（默认18磅）。

        表格对象本身不提供外部间距属性，改为把表格前一段落的段后间距、
        表格后一段落的段前间距设为指定磅值。
        """
        if self.work_doc is None:
            return

        try:
            content_end = self.work_doc.Content.End
        except:
            return

        for table in self.work_doc.Tables:
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
                    before_para.SpaceAfter = points
                except:
                    pass

            # 表格后一段落：表格结束位置起的段落
            if table_end < content_end:
                try:
                    after_para = self.work_doc.Range(
                        table_end, table_end + 1).Paragraphs(1)
                    after_para.SpaceBefore = points
                except:
                    pass

    def center_all_images(self):
        """将所有图片居中显示（始终执行，不受界面控件影响）"""
        if self.work_doc is None:
            return

        # 处理嵌入式图片：将所在段落设置为居中对齐
        for inline_shape in self.work_doc.InlineShapes:
            if inline_shape.Type == wc.wdInlineShapePicture:
                try:
                    inline_shape.Range.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
                except:
                    pass

        # 处理浮动图片：水平居中于页面
        for shape in self.work_doc.Shapes:
            try:
                shape.RelativeHorizontalPosition = wc.wdRelativeHorizontalPositionPage
                shape.Left = wc.wdShapeCenter
            except:
                pass

    def _auto_adjust_tables(self):
        """将所有表格按“根据内容调整表格”自动调整，并撑满页面宽度。

        分两步执行（等价于WPS中先点“自动调整——根据内容调整表格”，
        再让表格占满页面正文宽度）：
            1. AutoFitBehavior(wdAutoFitContent)：列宽按单元格内容自动分配；
            2. AutoFitBehavior(wdAutoFitWindow)：在各列内容比例的基础上，
               把整个表格拉伸到页面正文宽度。
        """
        if self.work_doc is None:
            return

        for table in self.work_doc.Tables:
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

    def _apply_table_alignment(self):
        """设置表格内容对齐：所有行文字垂直居中，首行文字水平居中。

        - 垂直方向：整表所有单元格垂直居中（保持各单元格原有的水平对齐方式）；
        - 水平方向：仅第一行（表头行）的单元格文字水平居中，其余行保持原样。
        """
        if self.work_doc is None:
            return

        for table in self.work_doc.Tables:
            # 所有行文字垂直方向居中显示
            try:
                table.Range.Cells.VerticalAlignment = wc.wdCellAlignVerticalCenter
            except:
                try:
                    table.Rows.VerticalAlignment = wc.wdCellAlignVerticalCenter
                except:
                    pass

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

    def set_tables_auto_adjust_and_align(self):
        """表格统一格式（始终执行，不受界面控件影响）。

        依次完成：
            1. 根据内容调整表格，再撑满页面宽度（_auto_adjust_tables）；
            2. 所有行文字垂直居中、首行文字水平居中（_apply_table_alignment）。
        """
        self._auto_adjust_tables()
        self._apply_table_alignment()

    def set_all_tables_max_width(self):
        """将所有表格调整为页面最大宽度（始终执行，不受界面控件影响）。

        保留旧接口名以便兼容；实际表格宽度处理已包含“根据内容调整表格”，
        并同时完成行垂直居中与首行水平居中。
        """
        self.set_tables_auto_adjust_and_align()

    def add_image_border(self):
        """为所有图片添加1px宽度的黑色外框（始终执行，不受界面控件影响）"""
        if self.work_doc is None:
            return

        border_edges = [wc.wdBorderTop, wc.wdBorderLeft,
                        wc.wdBorderBottom, wc.wdBorderRight]

        # 处理嵌入式图片：逐条边设置黑色细线边框
        for inline_shape in self.work_doc.InlineShapes:
            if inline_shape.Type == wc.wdInlineShapePicture:
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
        for shape in self.work_doc.Shapes:
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