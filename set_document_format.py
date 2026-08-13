import os
import win32com.client
import word_constants as wc
from word_constants import set_range_style, get_range_information, get_range_style, get_list_gallery


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
            paragraph = paragraphs(i)
            is_title = False
            try:
                paragraph_style = get_range_style(paragraph.Range)
                if paragraph_style is not None:
                    style_name = paragraph_style.NameLocal
                    is_title = "标题" in style_name or "MainTitle" in style_name
            except:
                pass

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
        
        alignment_map = {
            "左对齐": wc.wdAlignParagraphLeft,
            "右对齐": wc.wdAlignParagraphRight,
            "居中对齐": wc.wdAlignParagraphCenter,
            "两端对齐": wc.wdAlignParagraphJustify,
            "分散对齐": wc.wdAlignParagraphDistribute
        }
        word_alignment = alignment_map.get(alignment, wc.wdAlignParagraphLeft)
        
        paragraphs = self.work_doc.Paragraphs
        
        for i in range(1, paragraphs.Count + 1):
            paragraph = paragraphs(i)
            
            if paragraph.Range.Start in code_paragraph_indexes:
                continue
            
            if paragraph.Range.InlineShapes.Count > 0 or paragraph.Range.Tables.Count > 0:
                continue
            
            is_in_table = get_range_information(paragraph.Range, wc.wdWithInTable)
            
            try:
                paragraph_style = get_range_style(paragraph.Range)
                style_name = paragraph_style.NameLocal
                is_title = "标题" in style_name or "MainTitle" in style_name
            except:
                is_title = False
            
            if not is_in_table and not is_title:
                paragraph.Alignment = word_alignment
                paragraph.RightIndent = 0
            if indent_style is not None and not is_title and not is_in_table:
                if indent_style == "首行缩进2字符":
                    paragraph.FirstLineIndent = 2 * 0.35 * 28.35
                elif indent_style == "无缩进":
                    paragraph.FirstLineIndent = 0
                elif indent_style == "悬挂缩进":
                    paragraph.FirstLineIndent = -2 * 0.35 * 28.35
            else:
                paragraph.FirstLineIndent = 0
            
            if not is_title:
                if font is not None and paragraph_style is not None:
                    paragraph_style.Font.Name = font
                
                if font_size != "不变更":
                    try:
                        paragraph.Range.Font.Size = self._convert_chinese_font_size_to_points(font_size)
                    except:
                        pass
        
        if is_delete_empty_lines:
            for i in range(self.work_doc.Paragraphs.Count, 0, -1):
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
    
    def set_main_title_format(self, font, font_size, is_bold):
        if self.work_doc is None:
            return
        
        style = None
        try:
            style = self.work_doc.Styles("MainTitle")
        except:
            style = self.work_doc.Styles.Add("MainTitle", wc.wdStyleTypeParagraph)
        
        if font is not None:
            style.Font.Name = font
        if font_size is not None:
            style.Font.Size = self._convert_chinese_font_size_to_points(font_size)
        style.Font.Bold = 1 if is_bold else 0
        style.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
        style.ParagraphFormat.SpaceBefore = self.work_doc.Application.LinesToPoints(0.5)
        style.ParagraphFormat.SpaceAfter = self.work_doc.Application.LinesToPoints(0.5)
        
        set_range_style(self.work_doc.Paragraphs(1).Range, style)
        
        self.work_doc.Paragraphs(1).Alignment = wc.wdAlignParagraphCenter
    
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
        elif number_style == "1.":
            fmt = f"%{target_level}. "
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
    
    def set_images_and_tables(self, wrap_as_inline=True, no_indent=True, max_width=True):
        """设置图片与表格的格式。

        Args:
            wrap_as_inline: 是否将图片文字环绕类型设为嵌入型（默认True）
            no_indent: 是否取消图片/表格所在段落的缩进（默认True）
            max_width: 是否将图片/表格设为页面最大宽度（默认True）
        """
        if self.work_doc is None:
            return

        # 计算页面正文区域宽度（用于max_width）
        page_width = None
        if max_width:
            try:
                ps = self.work_doc.PageSetup
                page_width = ps.PageWidth - ps.LeftMargin - ps.RightMargin
            except:
                pass

        if wrap_as_inline:
            # 将所有浮动图形（Shapes）转换为嵌入型（InlineShapes）
            # 从后往前遍历，因为ConvertToInlineShape会从Shapes集合中移除当前项
            for i in range(self.work_doc.Shapes.Count, 0, -1):
                try:
                    shape = self.work_doc.Shapes(i)
                    shape.ConvertToInlineShape()
                except:
                    pass

            # 确保所有表格为嵌入型（无文字环绕）
            for table in self.work_doc.Tables:
                try:
                    table.Rows.WrapAroundText = False
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

        # 处理表格：最大宽度、不缩进
        for table in self.work_doc.Tables:
            try:
                if max_width:
                    table.AutoFitBehavior(wc.wdAutoFitWindow)

                if no_indent:
                    try:
                        table.Range.ParagraphFormat.FirstLineIndent = 0
                    except:
                        pass
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

    def set_all_tables_max_width(self):
        """将所有表格调整为页面最大宽度（始终执行，不受界面控件影响）"""
        if self.work_doc is None:
            return

        for table in self.work_doc.Tables:
            try:
                table.AutoFitBehavior(wc.wdAutoFitWindow)
            except:
                pass

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