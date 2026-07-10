import win32com.client
import word_constants as wc


class AddPageNumberErrorCodes:
    NoError = 0
    PageNumberAlreadyExists = 1
    DocumentIsEmpty = 2
    DocumentIsProtected = 3
    NoPageContent = 4


class PageNumberManager:
    add_page_number_error_code = AddPageNumberErrorCodes.NoError
    
    @staticmethod
    def add_page_numbers(word_doc,
                         page_number_position=wc.wdAlignPageNumberCenter,
                         number_format=wc.wdPageNumberStyleArabic):
        if not PageNumberManager.is_valid_document_for_page_numbering(word_doc):
            print("文档不符合添加页码的条件")
            return False
        
        try:
            document_sections = word_doc.Sections
            
            for document_section in document_sections:
                PageNumberManager._add_page_number_to_section(document_section, page_number_position, number_format)
            
            print("页码添加成功")
            return True
        except Exception as ex:
            print(f"添加页码时发生错误: {ex}")
            return False
    
    @staticmethod
    def add_page_numbers_custom(word_doc):
        for section in word_doc.Sections:
            ps = section.PageSetup
            ps.DifferentFirstPageHeaderFooter = 0
            ps.FooterDistance = 14.175
            footer = section.Footers(wc.wdHeaderFooterPrimary)
            
            footer.Range.Text = ""
            
            for idx in [wc.wdHeaderFooterPrimary, wc.wdHeaderFooterFirstPage, wc.wdHeaderFooterEvenPages]:
                header = section.Headers(idx)
                if header is not None:
                    header.Range.Text = ""
            
            footer.Range.ParagraphFormat.Alignment = wc.wdAlignParagraphCenter
            footer.Range.ParagraphFormat.SpaceAfter = 0
            footer.Range.ParagraphFormat.SpaceBefore = 0
            footer.Range.ParagraphFormat.LineSpacing = 12
            
            footer.Range.Text = "第页 / 共页"
            
            r1 = footer.Range
            r1.SetRange(footer.Range.Start + 1, footer.Range.Start + 1)
            r1.Fields.Add(r1, wc.wdFieldPage)
            
            r2 = footer.Range
            r2.SetRange(footer.Range.End - 2, footer.Range.End - 2)
            r2.Fields.Add(r2, wc.wdFieldNumPages)
        
        word_app = word_doc.Parent
        word_app.AutomationSecurity = wc.msoAutomationSecurityLow
        word_doc.Fields.Update()
        word_doc.Repaginate()
    
    @staticmethod
    def is_valid_document_for_page_numbering(target_document):
        PageNumberManager.add_page_number_error_code = AddPageNumberErrorCodes.NoError
        
        if PageNumberManager._has_existing_page_numbers(target_document):
            print("文档已存在页码")
            PageNumberManager.add_page_number_error_code = AddPageNumberErrorCodes.PageNumberAlreadyExists
            return False
        
        if target_document is None:
            print("目标文档为空")
            PageNumberManager.add_page_number_error_code = AddPageNumberErrorCodes.DocumentIsEmpty
            return False
        
        if target_document.ProtectionType != wc.wdNoProtection:
            print("文档受保护，无法添加页码")
            PageNumberManager.add_page_number_error_code = AddPageNumberErrorCodes.DocumentIsProtected
            return False
        
        if target_document.ComputeStatistics(wc.wdStatisticPages) == 0:
            print("文档没有页面内容")
            PageNumberManager.add_page_number_error_code = AddPageNumberErrorCodes.NoPageContent
            return False
        
        return True
    
    @staticmethod
    def _has_existing_page_numbers(target_document):
        try:
            for document_section in target_document.Sections:
                for header in document_section.Headers:
                    if PageNumberManager._contains_page_number_field(header.Range):
                        return True
                
                for footer in document_section.Footers:
                    if PageNumberManager._contains_page_number_field(footer.Range):
                        return True
            
            return False
        except Exception as ex:
            print(f"检查现有页码时发生错误: {ex}")
            return False
    
    @staticmethod
    def _contains_page_number_field(target_range):
        for field in target_range.Fields:
            if field.Type == wc.wdFieldPage:
                return True
        return False
    
    @staticmethod
    def _add_page_number_to_section(target_section, page_number_position, number_format):
        try:
            footer = target_section.Footers(wc.wdHeaderFooterPrimary)
            
            footer.Range.Delete()
            
            footer.Range.Fields.Add(
                Range=footer.Range,
                Type=wc.wdFieldPage)
            
            footer.PageNumbers.NumberStyle = number_format
            footer.PageNumbers.HeadingLevelForChapter = 0
            footer.PageNumbers.IncludeChapterNumber = False
            footer.PageNumbers.RestartNumberingAtSection = False
            footer.PageNumbers.StartingNumber = 1
            
            footer.PageNumbers.Add(page_number_position)
        except Exception as ex:
            print(f"为章节添加页码时发生错误: {ex}")
    
    @staticmethod
    def remove_page_numbers(target_document):
        if target_document is None:
            return False
        
        try:
            for document_section in target_document.Sections:
                for footer in document_section.Footers:
                    PageNumberManager._remove_page_number_fields(footer.Range)
                
                for header in document_section.Headers:
                    PageNumberManager._remove_page_number_fields(header.Range)
            
            print("页码移除成功")
            return True
        except Exception as ex:
            print(f"移除页码时发生错误: {ex}")
            return False
    
    @staticmethod
    def _remove_page_number_fields(target_range):
        fields_to_remove = []
        
        for field in target_range.Fields:
            if field.Type == wc.wdFieldPage:
                fields_to_remove.append(field)
        
        for field in fields_to_remove:
            try:
                field.Delete()
            except Exception as ex:
                print(f"删除页码字段时发生错误: {ex}")