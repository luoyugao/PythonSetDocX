import os
import win32com.client
import word_constants as wc
from word_constants import set_range_style, get_range_information, get_range_style, get_list_gallery


class SetDocumentFormat:
    def __init__(self, full_name=None, doc=None):
        self.work_doc = None
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
        
        self.work_doc.PageSetup.TopMargin = margin_points * top
        self.work_doc.PageSetup.BottomMargin = margin_points * bottom
        self.work_doc.PageSetup.LeftMargin = margin_points * left
        self.work_doc.PageSetup.RightMargin = margin_points * right
    
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
        
        style.Font.Name = font
        style.Font.Size = self._convert_chinese_font_size_to_points(font_size)
        style.Font.Bold = 1 if is_bold else 0
        style.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
        style.ParagraphFormat.SpaceAfter = 24
        
        set_range_style(self.work_doc.Paragraphs(1).Range, style)
        
        self.work_doc.Paragraphs(1).Alignment = wc.wdAlignParagraphCenter
    
    def set_title_styles(self, level_or_para, font=None, font_size=None,
                         number_style=None, indent_style=None):
        if self.work_doc is None:
            return None
        
        if isinstance(level_or_para, int):
            level = level_or_para
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
                list_gallery = get_list_gallery(self.work_doc.Application, wc.wdOutlineNumberGallery)
                list_template = None
                
                if number_style == "一.":
                    list_template = list_gallery.ListTemplates(1)
                elif number_style == "一）":
                    list_template = list_gallery.ListTemplates(2)
                elif number_style == "1.":
                    list_template = list_gallery.ListTemplates(3)
                elif number_style == "1)":
                    list_template = list_gallery.ListTemplates(4)
                elif number_style == "①":
                    list_template = list_gallery.ListTemplates(5)
                
                if list_template is not None:
                    style.LinkToListTemplate(list_template)
            
            style.Font.Bold = 1
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
                    list_gallery = get_list_gallery(self.work_doc.Application, wc.wdOutlineNumberGallery)
                    list_template = None
                    
                    if number_style == "一.":
                        list_template = list_gallery.ListTemplates(1)
                    elif number_style == "一）":
                        list_template = list_gallery.ListTemplates(2)
                    elif number_style == "1.":
                        list_template = list_gallery.ListTemplates(3)
                    elif number_style == "1)":
                        list_template = list_gallery.ListTemplates(4)
                    elif number_style == "①":
                        list_template = list_gallery.ListTemplates(5)
                    
                    if list_template is not None:
                        style.LinkToListTemplate(list_template)
                
                style.Font.Bold = 1
                set_range_style(para.Range, style)
            except Exception as ex:
                print(f"设置标题样式时出错: {ex}")
    
    def set_images_and_tables(self, wrap_as_inline=True):
        if self.work_doc is None:
            return
        
        for inline_shape in self.work_doc.InlineShapes:
            if inline_shape.Type == wc.wdInlineShapePicture:
                page_width = self.work_doc.PageSetup.PageWidth - \
                            self.work_doc.PageSetup.LeftMargin - \
                            self.work_doc.PageSetup.RightMargin
                inline_shape.Width = page_width
                inline_shape.Borders.Enable = 1
                
                try:
                    shape = inline_shape.ConvertToShape()
                    shape.RelativeHorizontalPosition = wc.wdRelativeHorizontalPositionPage
                    shape.WrapFormat.Type = wc.wdWrapSquare
                except:
                    pass
        
        if wrap_as_inline:
            for shape in self.work_doc.Shapes:
                try:
                    shape.WrapFormat.Type = wc.wdWrapTopBottom
                except:
                    pass
        
        for table in self.work_doc.Tables:
            table.AutoFitBehavior(wc.wdAutoFitWindow)
    
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