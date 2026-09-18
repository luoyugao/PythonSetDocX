import os
import json
import sys
import tkinter as tk

DEFAULT_WORK_FOLDER = os.path.join(os.path.expanduser("~"), "Documents")
WORK_FOLDER = DEFAULT_WORK_FOLDER
MAX_FOLDER_DEPTH = 6

word_app = None
work_doc = None
content_select_handler = None

SETTINGS_FOLDER = os.path.join(os.environ.get('APPDATA', ''), 'AutoAdjustWordFormatV2')
SETTINGS_FILENAME = os.path.join(SETTINGS_FOLDER, 'ParaLib.json')

word_app_check_timer = None
is_timer_running = False

WordAppNullDetected = None

AutoMargin = True
TopMargin = 1.5
BottomMargin = 1.5
LeftMargin = 1.5
RightMargin = 1.5

ChangeContentFormat = True
ContentIndent = "首行缩进2字符"
ContentAlign = "两端对齐"
ContentFont = "不变更"
ContentFontSize = "不变更"
CodeFormat = "自动"
DeleteEmptyLine = True
SetStandardLineParagraphSpacing = True

ChangeMainTitleFormat = True
MainTitleFont = "黑体"
MainTitleFontSize = "二号"
MainTitleBold = True

ChangeTitleLevelFormat = True
LevelTitleFonts = ["不变更", "不变更", "不变更", "不变更", "不变更"]
LevelTitleFontSizes = ["三号", "四号", "不变更", "不变更", "不变更"]
LevelTitleNumbersStyle = ["不变更", "不变更", "不变更", "不变更", "不变更"]
LevelTitleIndents = ["不变更", "不变更", "不变更", "不变更", "不变更"]

ChangeOrientation = False
IsPortrait = True

AddPageNumberEnable = True

ChangeImageFormat = True
TableOnly = False
ImgAutoWidth = True
ImgAutoHeight = True
ImageNoIndent = True
ImageNoBorder = True
WrapAsInline = True

CatalogueLevel = "显示3层目录"
UseTitleAsFilename = True

WindowLeft = -1
WindowTop = -1
WindowWidth = -1
WindowHeight = -1
WindowState = 0

HistoryFolderPaths = []


def initialize_word_app_checker():
    global word_app_check_timer
    import threading
    def check_word_app_status():
        global word_app, is_timer_running
        while is_timer_running:
            from universal import Universal
            word_app = Universal().get_active_word_app()
            if word_app is not None:
                is_timer_running = False
                global work_doc
                try:
                    if word_app.Documents.Count > 0:
                        work_doc = word_app.ActiveDocument
                except:
                    pass
            import time
            time.sleep(2)
    
    is_timer_running = True
    thread = threading.Thread(target=check_word_app_status)
    thread.daemon = True
    thread.start()


def start_word_app_checking():
    global word_app, is_timer_running
    if word_app is None and not is_timer_running:
        initialize_word_app_checker()
        if WordAppNullDetected:
            WordAppNullDetected(None, None)


def load_settings():
    global AutoMargin, TopMargin, BottomMargin, LeftMargin, RightMargin
    global ChangeContentFormat, ContentIndent, ContentAlign, ContentFont
    global ContentFontSize, CodeFormat, DeleteEmptyLine, SetStandardLineParagraphSpacing
    global ChangeMainTitleFormat, MainTitleFont, MainTitleFontSize, MainTitleBold
    global ChangeTitleLevelFormat, LevelTitleFonts, LevelTitleFontSizes, LevelTitleNumbersStyle, LevelTitleIndents
    global ChangeImageFormat, ImageNoIndent, TableOnly
    global ChangeOrientation, IsPortrait
    global AddPageNumberEnable
    global CatalogueLevel, UseTitleAsFilename
    global HistoryFolderPaths
    global WindowLeft, WindowTop, WindowWidth, WindowHeight, WindowState
    global WORK_FOLDER
    
    json_content = None
    if os.path.exists(SETTINGS_FILENAME):
        try:
            with open(SETTINGS_FILENAME, 'r', encoding='utf-8') as f:
                json_content = f.read()
        except:
            pass
    else:
        legacy = os.path.join(os.path.dirname(sys.executable), 'ParaLib.json')
        if os.path.exists(legacy):
            try:
                with open(legacy, 'r', encoding='utf-8') as f:
                    json_content = f.read()
            except:
                pass
    
    if not json_content:
        return
    
    try:
        settings = json.loads(json_content)
        
        AutoMargin = settings.get('autoMargin', True)
        TopMargin = float(settings.get('topMargin', 1.5))
        BottomMargin = float(settings.get('bottomMargin', 1.5))
        LeftMargin = float(settings.get('leftMargin', 1.5))
        RightMargin = float(settings.get('rightMargin', 1.5))
        
        ChangeContentFormat = settings.get('changeContentFormat', True)
        ContentIndent = settings.get('contentIndent', "首行缩进2字符")
        ContentAlign = settings.get('contentAlign', "两端对齐")
        ContentFont = settings.get('contentFont', "不变更")
        ContentFontSize = settings.get('contentFontSize', "不变更")
        CodeFormat = settings.get('codeFormat', "自动")
        
        ChangeMainTitleFormat = settings.get('changeMainTitleFormat', True)
        MainTitleFont = settings.get('mainTitleFont', "黑体")
        MainTitleFontSize = settings.get('mainTitleFontSize', "二号")
        MainTitleBold = settings.get('mainTitleBold', True)
        
        ChangeTitleLevelFormat = settings.get('changeTitleLevelFormat', True)
        LevelTitleFonts = settings.get('levelTitleFonts', ["不变更", "不变更", "不变更", "不变更", "不变更"])
        LevelTitleNumbersStyle = settings.get('levelTitleNumbers', ["不变更", "不变更", "不变更", "不变更", "不变更"])
        LevelTitleIndents = settings.get('levelTitleIndents', ["不变更", "不变更", "不变更", "不变更", "不变更"])
        
        ChangeImageFormat = settings.get('changeImageFormat', True)
        ImageNoIndent = settings.get('imageNoIndent', True)
        TableOnly = settings.get('tableOnly', False)
        
        ChangeOrientation = settings.get('changeOrientation', False)
        IsPortrait = settings.get('isPortrait', True)
        
        AddPageNumberEnable = settings.get('addPageNumberEnable', True)
        
        CatalogueLevel = settings.get('catalogueLevel', "显示3层目录")
        UseTitleAsFilename = settings.get('useTitleAsFilename', True)
        
        HistoryFolderPaths = settings.get('historyFolderPaths', [])
        
        WindowLeft = int(settings.get('windowLeft', -1))
        WindowTop = int(settings.get('windowTop', -1))
        WindowWidth = int(settings.get('windowWidth', -1))
        WindowHeight = int(settings.get('windowHeight', -1))
        WindowState = int(settings.get('windowState', 0))
        
        work_folder = settings.get('workFolder', DEFAULT_WORK_FOLDER)
        if os.path.isdir(work_folder):
            WORK_FOLDER = work_folder
        
        print(f"Loaded window bounds: left={WindowLeft} top={WindowTop} w={WindowWidth} h={WindowHeight} state={WindowState}")
    except Exception as ex:
        print(f"加载设置文件失败: {ex}")


def save_settings():
    settings = {
        'autoMargin': AutoMargin,
        'topMargin': TopMargin,
        'bottomMargin': BottomMargin,
        'leftMargin': LeftMargin,
        'rightMargin': RightMargin,
        'changeContentFormat': ChangeContentFormat,
        'contentIndent': ContentIndent,
        'contentAlign': ContentAlign,
        'contentFont': ContentFont,
        'contentFontSize': ContentFontSize,
        'codeFormat': CodeFormat,
        'changeMainTitleFormat': ChangeMainTitleFormat,
        'mainTitleFont': MainTitleFont,
        'mainTitleFontSize': MainTitleFontSize,
        'mainTitleBold': MainTitleBold,
        'changeTitleLevelFormat': ChangeTitleLevelFormat,
        'levelTitleFonts': LevelTitleFonts,
        'levelTitleNumbers': LevelTitleNumbersStyle,
        'levelTitleIndents': LevelTitleIndents,
        'changeImageFormat': ChangeImageFormat,
        'imageNoIndent': ImageNoIndent,
        'tableOnly': TableOnly,
        'changeOrientation': ChangeOrientation,
        'isPortrait': IsPortrait,
        'addPageNumberEnable': AddPageNumberEnable,
        'catalogueLevel': CatalogueLevel,
        'useTitleAsFilename': UseTitleAsFilename,
        'historyFolderPaths': HistoryFolderPaths,
        'windowLeft': WindowLeft,
        'windowTop': WindowTop,
        'windowWidth': WindowWidth,
        'windowHeight': WindowHeight,
        'windowState': WindowState,
        'workFolder': WORK_FOLDER
    }
    
    json_content = json.dumps(settings, ensure_ascii=False, indent=2)
    
    try:
        os.makedirs(os.path.dirname(SETTINGS_FILENAME), exist_ok=True)
        with open(SETTINGS_FILENAME, 'w', encoding='utf-8') as f:
            f.write(json_content)
    except Exception as ex:
        print(f"写入主设置文件失败: {ex}")
    
    try:
        legacy = os.path.join(os.path.dirname(sys.executable), 'ParaLib.json')
        with open(legacy, 'w', encoding='utf-8') as f:
            f.write(json_content)
    except Exception as ex:
        print(f"写入备用设置文件失败: {ex}")


def write_init_settings_to_form(main_form):
    load_settings()
    write_to_form(main_form)


def read_from_form(main_form):
    global AutoMargin, TopMargin, BottomMargin, LeftMargin, RightMargin
    global ChangeContentFormat, ContentIndent, ContentAlign, ContentFont, ContentFontSize, CodeFormat
    global ChangeMainTitleFormat, MainTitleFont, MainTitleFontSize, MainTitleBold
    global ChangeTitleLevelFormat, LevelTitleFonts, LevelTitleNumbersStyle, LevelTitleIndents
    global ChangeImageFormat, ImageNoIndent, TableOnly
    global ChangeOrientation, IsPortrait
    global AddPageNumberEnable, CatalogueLevel, UseTitleAsFilename
    
    try:
        AutoMargin = main_form.chk_change_page_margin.get()
        TopMargin = float(main_form.top_margin.get())
        BottomMargin = float(main_form.bottom_margin.get())
        LeftMargin = float(main_form.left_margin.get())
        RightMargin = float(main_form.right_margin.get())
        
        ChangeContentFormat = main_form.chk_change_content_format.get()
        ContentIndent = main_form.cmb_content_indent.get()
        ContentAlign = main_form.cmb_content_align.get()
        ContentFont = main_form.cmb_content_font.get()
        ContentFontSize = main_form.cmb_content_font_size.get()
        CodeFormat = main_form.cmb_code_format.get()
        
        ChangeMainTitleFormat = main_form.chk_change_main_title_format.get()
        MainTitleFont = main_form.cmb_main_title_font.get()
        MainTitleFontSize = main_form.cmb_main_title_font_size.get()
        MainTitleBold = main_form.chk_main_title_bold.get()
        
        ChangeTitleLevelFormat = main_form.chk_change_level_title_format.get()
        for i in range(5):
            font_control = main_form.level_title_font_vars[i]
            number_control = main_form.level_title_number_vars[i]
            indent_control = main_form.level_title_indent_vars[i]
            
            if font_control is not None:
                LevelTitleFonts[i] = font_control.get()
            if number_control is not None:
                LevelTitleNumbersStyle[i] = number_control.get()
            if indent_control is not None:
                LevelTitleIndents[i] = indent_control.get()
        
        ChangeImageFormat = main_form.chk_change_image_and_table_format.get()
        ImageNoIndent = main_form.chk_image_no_indent.get()
        TableOnly = main_form.chk_table_only.get()
        
        ChangeOrientation = main_form.chk_change_page_orientation.get()
        IsPortrait = main_form.radio_page_portrait.get()
        
        AddPageNumberEnable = main_form.chk_add_page_num.get()
        
        CatalogueLevel = main_form.cmb_catalogue_level.get()
        UseTitleAsFilename = main_form.chk_title_as_filename.get()
    except Exception as ex:
        print(f"从窗体加载设置失败: {ex}")


def write_to_form(main_form):
    try:
        main_form.chk_change_page_margin.set(AutoMargin)
        
        main_form.top_margin.delete(0, tk.END)
        main_form.top_margin.insert(0, f"{TopMargin:.1f}")
        main_form.bottom_margin.delete(0, tk.END)
        main_form.bottom_margin.insert(0, f"{BottomMargin:.1f}")
        main_form.left_margin.delete(0, tk.END)
        main_form.left_margin.insert(0, f"{LeftMargin:.1f}")
        main_form.right_margin.delete(0, tk.END)
        main_form.right_margin.insert(0, f"{RightMargin:.1f}")
        
        main_form.chk_change_content_format.set(ChangeContentFormat)
        main_form.cmb_content_indent.set(ContentIndent)
        main_form.cmb_content_align.set(ContentAlign)
        main_form.cmb_content_font.set(ContentFont)
        main_form.cmb_content_font_size.set(ContentFontSize)
        main_form.cmb_code_format.set(CodeFormat)
        
        main_form.chk_change_main_title_format.set(ChangeMainTitleFormat)
        main_form.cmb_main_title_font.set(MainTitleFont)
        main_form.cmb_main_title_font_size.set(MainTitleFontSize)
        main_form.chk_main_title_bold.set(MainTitleBold)
        
        main_form.chk_change_level_title_format.set(ChangeTitleLevelFormat)
        for i in range(5):
            font_control = main_form.level_title_font_vars[i]
            number_control = main_form.level_title_number_vars[i]
            indent_control = main_form.level_title_indent_vars[i]
            
            if font_control is not None:
                font_control.set(LevelTitleFonts[i])
            if number_control is not None:
                number_control.set(LevelTitleNumbersStyle[i])
            if indent_control is not None:
                indent_control.set(LevelTitleIndents[i])
        
        main_form.chk_change_image_and_table_format.set(ChangeImageFormat)
        main_form.chk_image_no_indent.set(ImageNoIndent)
        main_form.chk_table_only.set(TableOnly)
        
        main_form.chk_change_page_orientation.set(ChangeOrientation)
        main_form.radio_page_portrait.set(IsPortrait)
        
        main_form.cmb_catalogue_level.set(CatalogueLevel)
        main_form.chk_title_as_filename.set(UseTitleAsFilename)
    except Exception as ex:
        print(f"保存设置到窗体失败: {ex}")


def add_history_folder_path(folder_path):
    if not folder_path:
        return
    
    if folder_path in HistoryFolderPaths:
        HistoryFolderPaths.remove(folder_path)
    
    HistoryFolderPaths.insert(0, folder_path)
    
    while len(HistoryFolderPaths) > 8:
        HistoryFolderPaths.pop()


def save_history_folder_paths():
    try:
        if os.path.exists(SETTINGS_FILENAME):
            with open(SETTINGS_FILENAME, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            settings['historyFolderPaths'] = HistoryFolderPaths
            with open(SETTINGS_FILENAME, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
    except:
        pass