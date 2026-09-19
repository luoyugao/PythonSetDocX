import os
import re
import win32com.client
import word_constants as wc
from word_constants import (set_range_style, get_range_style,
                            get_effective_font_size, get_effective_font_names,
                            apply_font_size, apply_font_names,
                            get_effective_indents, apply_indents,
                            apply_indent_style)
from page_number_manager import PageNumberManager


class EventHandlers:
    def __init__(self, main_form):
        self.main_form = main_form
        self.level_style = None
        self.word_app_check_timer = None
        self.is_timer_running = False
        self.WordAppNullDetected = None
    
    def btn_all_title_level_up_click(self):
        if self.main_form.work_doc is None:
            return
        
        doc = self.main_form.work_doc
        
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            style = get_range_style(para)
            style_name = style.NameLocal
            
            new_style = None
            if style_name == "标题 6":
                new_style = "标题 5"
            elif style_name == "标题 5":
                new_style = "标题 4"
            elif style_name == "标题 4":
                new_style = "标题 3"
            elif style_name == "标题 3":
                new_style = "标题 2"
            elif style_name == "标题 2":
                new_style = "标题 1"
            
            if new_style is not None:
                set_range_style(para.Range, new_style)
    
    def btn_all_title_level_down_click(self):
        if self.main_form.work_doc is None:
            return
        
        doc = self.main_form.work_doc
        
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            style = get_range_style(para)
            style_name = style.NameLocal
            
            new_style = None
            if style_name == "标题 1":
                new_style = "标题 2"
            elif style_name == "标题 2":
                new_style = "标题 3"
            elif style_name == "标题 3":
                new_style = "标题 4"
            elif style_name == "标题 4":
                new_style = "标题 5"
            elif style_name == "标题 5":
                new_style = "标题 6"
            
            if new_style is not None:
                set_range_style(para.Range, new_style)
    
    def btn_delete_all_picture_click(self):
        from set_document_format import SetDocumentFormat
        
        if self.main_form.work_doc is None:
            return
        
        set_doc_format = SetDocumentFormat(doc=self.main_form.work_doc)
        set_doc_format.delete_all_pictures()
    
    def button_format_adjust_click(self):
        import doc_parameters_manager as dpm
        
        dpm.work_doc = dpm.word_app.ActiveDocument
        
        from set_document_format import SetDocumentFormat
        set_doc_format = SetDocumentFormat(doc=dpm.work_doc)
        
        if self.main_form.chk_change_page_margin.get():
            set_doc_format.set_page_margins(
                float(self.main_form.top_margin.get()),
                float(self.main_form.bottom_margin.get()),
                float(self.main_form.left_margin.get()),
                float(self.main_form.right_margin.get()))
        
        if self.main_form.chk_add_page_num.get():
            self._add_page_number()
        
        if self.main_form.chk_change_main_title_format.get():
            # "不变更"必须转成 None：set_main_title_format 以 None 表示该项
            # 不修改，并把段落原有字体字号写回样式。若直接把字符串
            # "不变更"传下去，会被当成字体名/字号名写进 Word。
            main_title_font = self.main_form.cmb_main_title_font.get()
            main_title_size = self.main_form.cmb_main_title_font_size.get()
            if main_title_font == "不变更":
                main_title_font = None
            if main_title_size == "不变更":
                main_title_size = None
            set_doc_format.set_main_title_format(
                main_title_font,
                main_title_size,
                self.main_form.chk_main_title_bold.get())
        
        if self.main_form.chk_change_image_and_table_format.get():
            set_doc_format.set_images_and_tables(
                wrap_as_inline=True,
                no_indent=self.main_form.chk_image_no_indent.get(),
                max_width=self.main_form.chk_max_width.get())

            # 表格统一格式：按内容自动调整后撑满页面宽度，
            # 所有行垂直居中、首行水平居中。
            # 放在勾选判断内：未勾选时不应改动图片与表格。
            set_doc_format.set_tables_auto_adjust_and_align()
        
        if self.main_form.chk_change_content_format.get():
            indent_style = self.main_form.cmb_content_indent.get()
            alignment = self.main_form.cmb_content_align.get()
            font = self.main_form.cmb_content_font.get()
            font_size = self.main_form.cmb_content_font_size.get()
            
            if font == "不变更":
                font = None
            if font_size == "不变更":
                font_size = None
            if indent_style == "不变更":
                indent_style = None
            if alignment == "不变更":
                alignment = None
            
            # 表格内段落与正文共用「正文」样式，而 set_content_style 改的是
            # 样式本身——不保护的话表格文字会被一并改成与正文相同的格式。
            # 改样式前快照，改完以直接格式写回。
            table_format_snapshot = \
                set_doc_format.snapshot_table_paragraph_format()

            set_doc_format.set_content_format(
                indent_style, alignment, font, font_size,
                None,
                self.main_form.chk_delete_empty_lines.get(),
                self.main_form.chk_standard_line_spacing.get())

            # 同步设置文档的「正文」样式本身（新版式 / 无直接格式的段落
            # 也随之一致）。传入的 None 表示该项在界面上是"不变更"。
            try:
                set_doc_format.set_content_style(
                    indent_style, alignment, font, font_size)
            finally:
                set_doc_format.restore_table_paragraph_format(
                    table_format_snapshot)
        
        if self.main_form.chk_change_level_title_format.get():
            self._set_level_title_styles(set_doc_format)

        # 图片段落不缩进（收尾钉死）：放在正文/标题格式处理之后。
        # set_content_style 改的是「正文」样式本身，图片段落多为该样式，
        # 样式里的首行缩进会被继承，图片因此被顶偏；此处用直接格式再清零一次。
        if (self.main_form.chk_change_image_and_table_format.get() or
                self.main_form.chk_change_content_format.get()) and \
                self.main_form.chk_image_no_indent.get():
            set_doc_format.set_image_paragraph_no_indent()
    
    def select_active_doc_button_click(self):
        try:
            from universal import Universal
            
            dpm.word_app = Universal().get_active_word_app()
            dpm.work_doc = dpm.word_app.ActiveDocument
            self.main_form._set_full_filename(dpm.work_doc.FullName)
        except:
            import tkinter as tk
            from tkinter import messagebox
            messagebox.showwarning("警告", "请先打开一个文档")
    
    def save_and_close_button_click(self, event=None):
        import tkinter as tk
        
        no_close = False
        
        if isinstance(event, tk.MouseEvent) and event.num == 3:
            no_close = True
        
        self._save_or_move_file(no_close=no_close)
        self.main_form._reflash_document_editor()
    
    def delete_button_click(self):
        if self.main_form.work_doc is None:
            return
        
        filename = self.main_form.work_doc.FullName
        
        try:
            self._close_active_document()
            os.remove(filename)
            self.main_form._reflash_document_editor()
            self.main_form.status_bar.config(text="删除成功")
        except Exception as ex:
            self.main_form.status_bar.config(text=f"删除文件失败: {ex}")
    
    def node_mouse_click(self, event):
        import tkinter as tk
        from tkinter import messagebox, filedialog
        
        if self.main_form._get_full_filename() == "未选择文件":
            messagebox.showwarning("警告", "没有指定需移动的文件")
            return
        
        try:
            target_dir = event.node.tag
            
            if target_dir == "OTHER_FOLDER":
                target_dir = filedialog.askdirectory()
                if not target_dir:
                    return
            
            filename = self.main_form.txt_new_filename.get().strip() or \
                       os.path.basename(self.main_form._get_full_filename())
            ext = os.path.splitext(self.main_form._get_full_filename())[1]
            target_file_path = os.path.join(target_dir, filename + ext)
            
            if not os.path.exists(target_file_path):
                try:
                    doc = self.main_form.work_doc
                    if doc is not None:
                        doc.SaveAs(target_file_path)
                        
                        try:
                            os.remove(self.main_form._get_full_filename())
                        except Exception as delete_ex:
                            messagebox.showwarning("警告", f"删除原文件失败: {delete_ex}\n文件已另存为: {target_file_path}")
                        
                        self.main_form._reflash_document_editor()
                except Exception as ex:
                    messagebox.showerror("错误", f"文件移动失败: {ex}")
                    return
            else:
                if event.num == 1:
                    try:
                        doc = self.main_form.work_doc
                        if doc is not None:
                            doc.Save()
                            
                            reopened_doc = self.main_form.word_app.Documents.Open(target_file_path)
                            self.main_form.work_doc = reopened_doc
                            
                            doc.Close(0)
                            
                            self.main_form._reflash_document_editor()
                    except Exception as ex:
                        messagebox.showerror("错误", f"切换文件失败: {ex}")
                elif event.num == 3:
                    self._close_active_document()
                    self.main_form._set_full_filename("未选择文件")
        except Exception as ex:
            messagebox.showerror("错误", f"操作失败: {ex}")
    
    def initialize_word_app_checker(self):
        import threading
        
        def check_word_app_status():
            import doc_parameters_manager as dpm
            from universal import Universal
            
            while self.is_timer_running:
                dpm.word_app = Universal().get_active_word_app()
                if dpm.word_app is not None:
                    try:
                        if dpm.word_app.Documents.Count > 0:
                            self.is_timer_running = False
                            dpm.work_doc = dpm.word_app.ActiveDocument
                    except:
                        pass
                import time
                time.sleep(1)
        
        self.is_timer_running = True
        thread = threading.Thread(target=check_word_app_status)
        thread.daemon = True
        thread.start()
    
    def start_word_app_checking(self):
        import doc_parameters_manager as dpm
        
        if dpm.word_app is None and not self.is_timer_running:
            self.initialize_word_app_checker()
            if self.WordAppNullDetected:
                self.WordAppNullDetected(None, None)
    
    def _add_page_number(self):
        import doc_parameters_manager as dpm
        
        if not PageNumberManager.is_valid_document_for_page_numbering(dpm.work_doc):
            if PageNumberManager.add_page_number_error_code == PageNumberManager.AddPageNumberErrorCodes.PageNumberAlreadyExists:
                is_custom_format = self._is_page_number_format_custom(dpm.work_doc)
                if not is_custom_format:
                    import tkinter as tk
                    from tkinter import messagebox
                    b = messagebox.askyesno("提示", "文档已存在自定义格式页码，是否需要重置页码？")
                    if not b:
                        return
        
        PageNumberManager.add_page_numbers_custom(dpm.work_doc)
    
    def _is_page_number_format_custom(self, doc):
        try:
            for section in doc.Sections:
                footer = section.Footers(wc.wdHeaderFooterPrimary)
                if not footer.Exists:
                    continue
                
                has_page_field = False
                has_total_pages_field = False
                
                for field in footer.Range.Fields:
                    if field.Type == wc.wdFieldPage:
                        has_page_field = True
                    elif field.Type == wc.wdFieldNumPages:
                        has_total_pages_field = True
                
                if has_page_field and has_total_pages_field:
                    footer_text = footer.Range.Text.strip()
                    if "第" in footer_text and "页" in footer_text and "/" in footer_text and "共" in footer_text:
                        return True
            
            return False
        except Exception as ex:
            print(f"检查页码格式时发生错误: {ex}")
            return False
    
    def _close_active_document(self):
        import doc_parameters_manager as dpm
        
        doc = dpm.word_app.ActiveDocument
        doc.Close(0)
    
    def _save_or_move_file(self, new_path=None, no_close=False):
        import doc_parameters_manager as dpm
        
        word_app = dpm.word_app
        doc = dpm.word_app.ActiveDocument
        
        old_name = os.path.splitext(os.path.basename(self.main_form._get_full_filename()))[0]
        
        if word_app is None:
            return
        
        if new_path is None and old_name == self.main_form.txt_new_filename.get():
            try:
                doc.Save()
                if not no_close:
                    self._close_active_document()
            except:
                self.main_form._set_full_filename("未选择文件")
            return
        
        is_move = False
        if new_path is None:
            new_path = doc.Path
        else:
            is_move = True
        
        if self.main_form.txt_new_filename.get().strip():
            new_name = self.main_form.txt_new_filename.get().strip()
            
            if any(char in new_name for char in '<>:"/\\|?*'):
                import tkinter as tk
                from tkinter import messagebox
                messagebox.showerror("错误", "文件名包含非法字符")
                return
            
            ext = os.path.splitext(self.main_form._get_full_filename())[1]
            new_full_filename = os.path.join(new_path, new_name + ext)
            
            if new_name != old_name or new_path != dpm.WORK_FOLDER:
                try:
                    doc.SaveAs(new_full_filename)
                    if not no_close:
                        self._close_active_document()
                    
                    if new_name != old_name or is_move:
                        os.remove(self.main_form._get_full_filename())
                    
                    self.main_form.txt_new_filename.delete(0, tk.END)
                except Exception as ex:
                    pass
    
    def _set_level_title_styles(self, set_doc_format):
        self.level_style = [None, None, None, None, None]
        # 字体/字号/缩进方式选择"不变更"的级别：这些级别必须保留段落处理前的
        # 字体、字号与缩进。集合在这里（任何样式被创建/改动之前）直接从下拉框
        # 读取，供文末的"处理前快照"使用——晚于 set_title_styles 读取就只能
        # 读到已被改动的值。
        keep_font_levels = set()
        keep_size_levels = set()
        keep_indent_levels = set()
        title_indent_styles = {}        # 级别索引 -> 界面上的"缩进方式"文本
        for i in range(5):
            try:
                if self.main_form.level_title_font_vars[i].get() in ("", "不变更"):
                    keep_font_levels.add(i)
                if self.main_form.level_title_font_size_vars[i].get() in ("", "不变更"):
                    keep_size_levels.add(i)
                indent_value = self.main_form.level_title_indent_vars[i].get()
                title_indent_styles[i] = indent_value
                if indent_value in ("", "不变更"):
                    keep_indent_levels.add(i)
            except Exception:
                pass
        
        try:
            doc = set_doc_format.work_doc
            if doc is None:
                doc = dpm.work_doc
            
            if doc is None:
                return
            
            para_count = doc.Paragraphs.Count
            
            # ---- 处理前快照：记录"不变更"级别段落的字体、字号与缩进 ----
            # 必须早于任何 set_title_styles 调用：它会创建/改动"标题 N"样式
            # 并链接列表模板，可能连带改变段落的有效字体、字号与缩进，
            # 那时再读就只能读到已被改动的值，"保留原格式"也就失效了。
            preserved_fonts = {}        # para_index -> (字号, {槽位: 字体名}, {缩进属性: 值})
            for para_index in range(1, para_count + 1):
                try:
                    paragraph = doc.Paragraphs(para_index)
                    level_index = set_doc_format.get_paragraph_outline_level(paragraph)
                    if level_index < 1 or level_index > 5:
                        continue
                    idx = level_index - 1
                    need_size = idx in keep_size_levels
                    need_font = idx in keep_font_levels
                    need_indent = idx in keep_indent_levels
                    if not need_size and not need_font and not need_indent:
                        continue
                    size = get_effective_font_size(paragraph.Range) if need_size else None
                    names = get_effective_font_names(paragraph.Range) if need_font else None
                    indents = get_effective_indents(paragraph.Range) if need_indent else None
                    if size is not None or names or indents:
                        preserved_fonts[para_index] = (size, names, indents)
                except Exception:
                    continue
            
            for para_index in range(1, para_count + 1):
                try:
                    paragraph = doc.Paragraphs(para_index)
                    level_index = set_doc_format.get_paragraph_outline_level(paragraph)
                    
                    if level_index < 1 or level_index > 5:
                        continue
                    
                    if self.level_style[level_index - 1] is None:
                        font = None
                        font_size = None
                        number_style = None
                        indent = None
                        
                        controls = self.main_form.level_title_font_vars[level_index - 1]
                        if controls.get() != "不变更":
                            font = controls.get()
                        
                        controls = self.main_form.level_title_font_size_vars[level_index - 1]
                        if controls.get() != "不变更":
                            font_size = controls.get()
                        
                        controls = self.main_form.level_title_number_vars[level_index - 1]
                        if controls.get() != "不变更":
                            number_style = controls.get()
                        
                        controls = self.main_form.level_title_indent_vars[level_index - 1]
                        if controls.get() != "不变更":
                            indent = controls.get()
                        
                        if font is not None or font_size is not None or indent is not None or number_style is not None:
                            self.level_style[level_index - 1] = set_doc_format.set_title_styles(
                                level_index, font, font_size, number_style, indent)
                    
                    if self.level_style[level_index - 1] is not None:
                        set_range_style(paragraph.Range, self.level_style[level_index - 1])

                        # 写回处理前的字体/字号/缩进（本段落处理的最后一步）。
                        # 对"不变更"的级别，用处理前快照的值以直接格式覆盖回来：
                        # 套用"标题 N"样式与链接列表模板都会把列表模板级别的
                        # 字体与缩进位置（NumberPosition / TextPosition）带到
                        # 段落上，先写就会被冲掉。
                        saved_size, saved_font_names, saved_indents = \
                            preserved_fonts.get(para_index, (None, None, None))
                        apply_font_size(paragraph.Range, saved_size)
                        apply_font_names(paragraph.Range, saved_font_names)
                        if (level_index - 1) in keep_indent_levels:
                            # "不变更"：还原段落处理前的缩进
                            apply_indents(paragraph.Range, saved_indents)
                        else:
                            # 明确选了缩进方式：在列表模板之后强制生效，
                            # 否则会被列表级别的 NumberPosition 顶掉
                            apply_indent_style(
                                paragraph.Range.ParagraphFormat,
                                title_indent_styles.get(level_index - 1))
                except:
                    continue
        except:
            pass