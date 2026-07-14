import os
import sys
import subprocess
import threading
import tkinter as tk
import tkinter.font
from tkinter import ttk, filedialog, messagebox
import doc_parameters_manager as dpm
from universal import Universal
from word_constants import set_range_style, get_range_style
import set_document_format


class MainForm:
    """主窗体类，负责整个应用的界面创建和事件处理"""
    
    def __init__(self, root):
        """初始化主窗体"""
        self.root = root
        # 设置窗口标题为"文档格式设置PY1.0"
        self.root.title("文档格式设置PY程序1.0")
        # 设置窗口初始大小为738x836
        self.root.geometry("738x836")
        # 允许窗口调整大小（宽度和高度均可调整）
        self.root.resizable(True, True)
        
        # Word应用程序COM对象，用于与Word进行交互操作
        self.word_app = None
        # 当前正在处理的Word文档对象
        self.work_doc = None
        # 缓存文档第一段文本，用于检测内容变化
        self._last_first_para_text = None
        # 临时抑制文档监控标志，用于右键移动后防止监控重置UI
        self._suppress_monitor = False
        # 单击/双击防抖定时器
        self._click_timer = None
        
        # 创建所有界面控件，包括左侧设置面板、右侧目录树和底部操作栏
        self._create_widgets()
        
        # 从配置文件（ParaLib.json）加载之前保存的设置值到界面控件
        dpm.write_init_settings_to_form(self)
        
        # 初始化当前根路径为工作目录（已从配置文件加载）
        self.current_root_path = dpm.WORK_FOLDER
        # 从配置加载目录层级
        level_text = self.cmb_catalogue_level.get()
        levels = {"显示1层目录": 1, "显示2层目录": 2, "显示3层目录": 3, "显示4层目录": 4, "显示5层目录": 5}
        self.current_level = levels.get(level_text, 3)
        # 加载工作目录下的目录结构
        self._load_directory_tree(self.current_root_path, self.current_level)
        
        # 从配置中恢复上次关闭时的窗口位置
        self._apply_saved_window_bounds()
        
        # 初始化事件处理器（预留方法，当前为空实现）
        self._initialize_event_handlers()
        
        # 绑定窗口关闭事件，在关闭前保存配置
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 检查系统中是否有正在运行的Word应用，如有则连接并获取活动文档
        self._check_word_app()
        
        # 绑定窗口尺寸变化事件，更新状态栏显示当前尺寸
        self.root.bind('<Configure>', self._on_window_configure)
        

        
        # 启动对Word活动文档的监控（每1秒检查一次）
        self._monitor_active_document()
    
    def _create_widgets(self):
        """创建主界面布局结构"""
        # 创建主框架作为容器，设置3像素内边距，更紧凑
        main_frame = ttk.Frame(self.root, padding="3")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左侧设置面板框架，固定宽度560像素（更宽以展示更多内容），淡黄色背景
        left_frame = tk.Frame(main_frame, width=420, bg="#FFF8DC")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 2))
        left_frame.pack_propagate(False)  # 防止子控件改变框架大小
        
        # 创建右侧目录树面板框架，固定宽度200像素
        right_frame = ttk.Frame(main_frame, width=70)
        
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 0))
        right_frame.pack_propagate(False)  # 防止子控件改变框架大小
        
        # 在左侧面板中创建各种设置区域控件
        self._create_left_panel(left_frame)
        # 在右侧面板中创建目录树控件
        self._create_right_panel(right_frame)
        
        # 创建底部操作面板框架
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=(3, 3), pady=2)
        
        # 在底部面板中创建文件信息显示和操作按钮
        self._create_bottom_panel(bottom_frame)
    
    def _create_left_panel(self, parent):
        """创建左侧可滚动的设置面板"""
        style = ttk.Style()
        style.configure("Bold.TCheckbutton", font=('TkDefaultFont', 9, 'bold'), foreground='#1976D2')
        style.map("Bold.TCheckbutton", foreground=[('active', '#1976D2'), ('!disabled', '#1976D2')])
        canvas = tk.Canvas(parent, bg="#FFF8DC")
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        # 创建可滚动的框架，用于容纳所有设置区域控件，淡黄色背景
        scrollable_frame = tk.Frame(canvas, bg="#FFF8DC")
        
        # 绑定框架大小变化事件，当内容区域大小改变时更新画布的滚动区域范围
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # 将可滚动框架嵌入到画布中，左上角对齐
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        # 将画布的滚动控制与滚动条关联，实现同步滚动
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 将画布放置在左侧，填充父容器并允许拉伸
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # 将滚动条放置在右侧，垂直方向填充
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定鼠标滚轮事件，实现通过滚轮控制画布滚动
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # 创建各个设置区域（按从上到下顺序）
        self._create_page_margin_section(scrollable_frame)      # 页面边距设置区域
        self._create_content_format_section(scrollable_frame)   # 正文格式设置区域
        self._create_main_title_section(scrollable_frame)       # 文章总标题格式设置区域
        self._create_level_title_section(scrollable_frame)      # 章节标题格式设置区域
        self._create_page_orientation_section(scrollable_frame) # 纸张方向和页码设置区域
        self._create_image_table_section(scrollable_frame)      # 图片和表格格式设置区域
        self._create_cancel_button(scrollable_frame)            # 取消所有设置按钮
    
    def _create_page_margin_section(self, parent):
        """创建页面边距设置区域（参照C#界面紧凑布局）"""
        # group = ttk.LabelFrame(parent, text="调整页面边距")
        group = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="black", highlightcolor="black")
        group.pack(fill=tk.X, padx=3, pady=(3, 3))
        
        self.chk_change_page_margin = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="调整页面边距", variable=self.chk_change_page_margin, style="Bold.TCheckbutton").grid(
            row=0, column=0, columnspan=6, sticky=tk.W, padx=2, pady=3)
        
        # 第一行：上边距、下边距 + 全部1CM按钮
        ttk.Label(group, text="上边距:", anchor=tk.E).grid(row=1, column=0, sticky=tk.E, padx=(2,0), pady=3)
        self.top_margin = ttk.Entry(group, width=7)
        self.top_margin.grid(row=1, column=1, padx=(0,5), pady=3, sticky=tk.W)
        
        ttk.Label(group, text="下边距:", anchor=tk.E).grid(row=1, column=2, sticky=tk.E, padx=(2,0), pady=3)
        self.bottom_margin = ttk.Entry(group, width=7)
        self.bottom_margin.grid(row=1, column=3, padx=(0,5), pady=3, sticky=tk.W)
        
        ttk.Button(group, text="全部1CM", command=self._on_set_all_margins_to_1cm).grid(
            row=1, column=4, columnspan=2, padx=2, pady=3)
        
        # 第二行：左边距、右边距
        ttk.Label(group, text="左边距:", anchor=tk.E).grid(row=2, column=0, sticky=tk.E, padx=(2,0), pady=3)
        self.left_margin = ttk.Entry(group, width=7)
        self.left_margin.grid(row=2, column=1, padx=(0,5), pady=3, sticky=tk.W)
        
        ttk.Label(group, text="右边距:", anchor=tk.E).grid(row=2, column=2, sticky=tk.E, padx=(2,0), pady=3)
        self.right_margin = ttk.Entry(group, width=7)
        self.right_margin.grid(row=2, column=3, padx=(0,5), pady=3, sticky=tk.W)
    
    def _create_content_format_section(self, parent):
        """创建正文格式设置区域（参照C#紧凑布局）"""
        # group = ttk.LabelFrame(parent, text="变更正文格式")
        group = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="black", highlightcolor="black")
        group.pack(fill=tk.X, padx=3, pady=(3, 3))
        
        self.chk_change_content_format = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="变更正文格式", variable=self.chk_change_content_format, style="Bold.TCheckbutton").grid(
            row=0, column=0, columnspan=6, sticky=tk.W, padx=2, pady=3)
        
        # 第一行：缩进方式、对齐方式
        ttk.Label(group, text="缩进方式:", anchor=tk.E).grid(row=1, column=0, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_content_indent = ttk.Combobox(
            group, values=["首行缩进2字符", "无缩进", "悬挂缩进", "不变更"], width=14)
        self.cmb_content_indent.grid(row=1, column=1, padx=(0,3), pady=3, sticky=tk.W)
        
        ttk.Label(group, text="对齐方式:", anchor=tk.E).grid(row=1, column=2, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_content_align = ttk.Combobox(
            group, values=["左对齐", "右对齐", "居中对齐", "两端对齐", "分散对齐"], width=10)
        self.cmb_content_align.grid(row=1, column=3, padx=(0,3), pady=3, sticky=tk.W)
        
        # 第二行：字体、字号、代码格式
        ttk.Label(group, text="字体:", anchor=tk.E).grid(row=2, column=0, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_content_font = ttk.Combobox(
            group, values=["宋体", "黑体", "楷体", "仿宋", "不变更"], width=10)
        self.cmb_content_font.grid(row=2, column=1, padx=(0,3), pady=3, sticky=tk.W)
        
        ttk.Label(group, text="字号:", anchor=tk.E).grid(row=2, column=2, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_content_font_size = ttk.Combobox(
            group, values=["五号", "小四", "四号", "三号", "不变更"], width=10)
        self.cmb_content_font_size.grid(row=2, column=3, padx=(0,3), pady=3, sticky=tk.W)
        
        ttk.Label(group, text="代码格式:", anchor=tk.E).grid(row=0, column=2, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_code_format = ttk.Combobox(
            group, values=["自动", "不变更"], width=10)
        self.cmb_code_format.grid(row=0, column=3, padx=(0,2), pady=3, sticky=tk.W)
        
        # 第三行：删除空行、标准行段间距（短文本，同行排列）
        self.chk_delete_empty_lines = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="删除空行", variable=self.chk_delete_empty_lines).grid(
            row=3, column=0, columnspan=2, sticky=tk.W, padx=2, pady=3)
        
        self.chk_standard_line_spacing = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="标准行段间距", variable=self.chk_standard_line_spacing).grid(
            row=3, column=2, columnspan=2, sticky=tk.W, padx=2, pady=3)
    
    def _create_main_title_section(self, parent):
        """创建文章首行总标题设置区域"""
        group = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="black", highlightcolor="black") # , text="变更文章首行总标题"
        group.pack(fill=tk.X, padx=3, pady=(3, 3))
        
        self.chk_change_main_title_format = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="变更文章首行总标题", variable=self.chk_change_main_title_format, style="Bold.TCheckbutton").grid(
            row=0, column=0, columnspan=5, sticky=tk.W, padx=2, pady=3)
        
        # 字体、字号、加粗并列
        ttk.Label(group, text="字体:", anchor=tk.E).grid(row=1, column=0, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_main_title_font = ttk.Combobox(
            group, values=["宋体", "黑体", "楷体", "仿宋"], width=10)
        self.cmb_main_title_font.grid(row=1, column=1, padx=(0,3), pady=3, sticky=tk.W)
        
        ttk.Label(group, text="字号:", anchor=tk.E).grid(row=1, column=2, sticky=tk.E, padx=(2,0), pady=3)
        self.cmb_main_title_font_size = ttk.Combobox(
            group, values=["二号", "三号", "四号"], width=10)
        self.cmb_main_title_font_size.grid(row=1, column=3, padx=(0,3), pady=3, sticky=tk.W)
        
        self.chk_main_title_bold = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="加粗", variable=self.chk_main_title_bold).grid(
            row=1, column=4, sticky=tk.W, padx=2, pady=3)
    
    def _create_level_title_section(self, parent):
        """创建各章节标题格式设置区域（参照C#紧凑布局）"""
        group = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="black", highlightcolor="black") # , text="变更各章节标题格式"
        group.pack(fill=tk.X, padx=3, pady=(3, 3))
        
        self.chk_change_level_title_format = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="变更各章节标题格式", variable=self.chk_change_level_title_format, style="Bold.TCheckbutton").grid(
            row=0, column=0, columnspan=5, sticky=tk.W, padx=2, pady=3)
        
        # 升级/降级按钮行（短标签）
        btn_frame = ttk.Frame(group)
        btn_frame.grid(row=0, column=0, columnspan=5, pady=3, sticky=tk.E)
        self.btn_title_level_up = ttk.Button(btn_frame, text="各级标题升一级", command=self._on_title_level_up)
        self.btn_title_level_up.pack(side=tk.LEFT, padx=2)
        self.btn_title_level_down = ttk.Button(btn_frame, text="各级标题降一级", command=self._on_title_level_down)
        self.btn_title_level_down.pack(side=tk.LEFT, padx=2)
        
        # 表头行
        headers = ["标题级别", "字体", "字号", "序号样式", "缩进方式"]
        for i, header in enumerate(headers):
            ttk.Label(group, text=header).grid(row=2, column=i, padx=2, pady=3)
        
        self.level_title_font_vars = []
        self.level_title_font_size_vars = []
        self.level_title_number_vars = []
        self.level_title_indent_vars = []
        
        fonts = ["不变更", "宋体", "黑体", "楷体", "仿宋"]
        font_sizes = ["不变更", "二号", "小二", "三号", "四号", "小四", "五号"]
        number_styles = ["不变更", "资料1. ", "一.", "一）", "1.", "1)", "①"]
        indents = ["不变更", "无缩进", "首行缩进2字符"]
        
        for i in range(5):
            ttk.Label(group, text=f"{i+1}级").grid(row=i+3, column=0, padx=2, pady=1)
            
            font_var = ttk.Combobox(group, values=fonts, width=8)
            font_var.grid(row=i+3, column=1, padx=1, pady=1)
            self.level_title_font_vars.append(font_var)
            
            font_size_var = ttk.Combobox(group, values=font_sizes, width=8)
            font_size_var.grid(row=i+3, column=2, padx=1, pady=1)
            self.level_title_font_size_vars.append(font_size_var)
            
            number_var = ttk.Combobox(group, values=number_styles, width=8)
            number_var.grid(row=i+3, column=3, padx=1, pady=1)
            self.level_title_number_vars.append(number_var)
            
            indent_var = ttk.Combobox(group, values=indents, width=10)
            indent_var.grid(row=i+3, column=4, padx=1, pady=1)
            self.level_title_indent_vars.append(indent_var)
        
        self.level_title_font_vars[0].set("黑体")
        self.level_title_font_vars[1].set("黑体")
        
        self.level_title_font_size_vars[0].set("四号")
        self.level_title_font_size_vars[1].set("小四")
        self.level_title_font_size_vars[2].set("不变更")
        self.level_title_font_size_vars[3].set("不变更")
        self.level_title_font_size_vars[4].set("不变更")
        
        self.level_title_number_vars[0].set("不变更")
        self.level_title_number_vars[1].set("不变更")
    
    def _create_page_orientation_section(self, parent):
        """创建纸张方向和页码设置区域"""
        group = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="black", highlightcolor="black") # , text="变更纸张方向")
        group.pack(fill=tk.X, padx=3, pady=(3, 3))
        
        self.chk_change_page_orientation = tk.BooleanVar(value=False)
        ttk.Checkbutton(group, text="变更纸张方向", variable=self.chk_change_page_orientation, style="Bold.TCheckbutton").grid(
            row=0, column=0, sticky=tk.W, padx=2, pady=3)
        
        self.radio_page_portrait = tk.BooleanVar(value=True)
        ttk.Radiobutton(group, text="纵向", variable=self.radio_page_portrait, value=True).grid(
            row=0, column=1, padx=2, pady=3)
        ttk.Radiobutton(group, text="横向", variable=self.radio_page_portrait, value=False).grid(
            row=0, column=2, padx=2, pady=3)
        
        self.chk_add_page_num = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="添加页码", variable=self.chk_add_page_num).grid(
            row=0, column=3, sticky=tk.W, padx=(5,2), pady=3)
    
    def _create_image_table_section(self, parent):
        """创建图片和表格格式设置区域（参照C#紧凑排列）"""
        group = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground="black", highlightcolor="black") # , text="变更图片和表格格式")
        group.pack(fill=tk.X, padx=3, pady=(3, 3))
        
        self.chk_change_image_and_table_format = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="变更图片与表格的格式", variable=self.chk_change_image_and_table_format, style="Bold.TCheckbutton").grid(
            row=0, column=0, columnspan=5, sticky=tk.W, padx=2, pady=3)
        
        # 复选框行：不缩进、最大宽度
        self.chk_image_no_indent = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="不缩进", variable=self.chk_image_no_indent).grid(
            row=1, column=0, sticky=tk.W, padx=2, pady=3)
        
        self.chk_max_width = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="最大宽度", variable=self.chk_max_width).grid(
            row=1, column=1, sticky=tk.W, padx=2, pady=3)
        
        self.btn_cancel_wrap_as_inline = ttk.Button(group, text="取消嵌入型环绕", command=self._on_cancel_wrap_as_inline)
        self.btn_cancel_wrap_as_inline.grid(row=1, column=2, sticky=tk.W, padx=(44, 2), pady=3)

        group.grid_columnconfigure(3, weight=1)

        # 删除所有图片按钮（与上一行控件同行，靠最右排列）
        self.btn_delete_all_pic = ttk.Button(group, text="删除所有图片", command=self._on_delete_all_pictures)
        self.btn_delete_all_pic.grid(row=1, column=4, sticky=tk.E, padx=2, pady=3)
    
    def _create_cancel_button(self, parent):
        """创建取消所有设置按钮和使能所有设置按钮"""
        btn_frame = tk.Frame(parent, bg="#FFF8DC", highlightthickness=0)
        btn_frame.pack(fill=tk.X, padx=3, pady=(16, 3))
        center_frame = ttk.Frame(btn_frame)
        center_frame.pack(expand=True)
        self.btn_cancel_settings = ttk.Button(center_frame, text="取消所有设置", command=self._on_cancel_settings, width=15)
        self.btn_cancel_settings.pack(side=tk.LEFT, padx=(0, 11))
        self.btn_enable_settings = ttk.Button(center_frame, text="使能所有设置", command=self._on_enable_settings, width=15)
        self.btn_enable_settings.pack(side=tk.LEFT, padx=0)
    
    def _create_right_panel(self, parent):
        """创建右侧目录树面板（去除水平滚动条，设置列宽度确保完整显示）"""
        style = ttk.Style()
        style.configure("FileInfo.TLabelframe", font=('TkDefaultFont', 9, 'bold'), foreground='#1976D2')
        style.configure("FileInfo.TLabelframe.Label", font=('TkDefaultFont', 9, 'bold'), foreground='#1976D2')
        style.configure("SaveDir.TLabelframe", font=('TkDefaultFont', 11, 'bold'), foreground='#1976D2')
        style.configure("SaveDir.TLabelframe.Label", font=('TkDefaultFont', 11, 'bold'), foreground='#1976D2')
        # 创建目录树
        title_frame = ttk.LabelFrame(parent, text="选择保存目录", style="SaveDir.TLabelframe")
        
        title_frame.pack(fill=tk.X, padx=0, pady=3, ipady=8)
        
        # 选择目录按钮
        btn_select_dir = ttk.Button(title_frame, text="选择目录", command=self._on_select_directory)
        btn_select_dir.pack(side=tk.LEFT, padx=1)
        
        # 目录级数下拉框
        ttk.Label(title_frame, text="显示级数:").pack(side=tk.LEFT, padx=1)
        self.cmb_catalogue_level = ttk.Combobox(
            title_frame,
            values=["显示1层目录", "显示2层目录", "显示3层目录", "显示4层目录", "显示5层目录"],
            width=15)
        self.cmb_catalogue_level.pack(side=tk.LEFT, padx=1)
        self.cmb_catalogue_level.bind("<<ComboboxSelected>>", self._on_catalogue_level_changed)
        self.cmb_catalogue_level.set("显示5层目录")
        
        # 提示标签（上：鼠标操作说明，下：文件夹图标说明）
        hint_label1 = ttk.Label(
            parent,
            text="左键：展开/折叠目录，右键：移动文件但不关闭",
            foreground="dark red",
            justify=tk.LEFT)
        hint_label1.pack(fill=tk.X, padx=2, pady=(0, 1))

        hint_label2 = ttk.Label(
            parent,
            text="双击：可用资源管理器打开文件夹",
            foreground="dark red",
            justify=tk.LEFT)
        hint_label2.pack(fill=tk.X, padx=2, pady=(0, 2))
        
        # 目录树容器（无水平边距，充分利用宽度）
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=0)
        
        self.tree = ttk.Treeview(tree_frame, show="tree")
        # 设置树列宽度为330px，确保常见目录名完整显示

        self.tree.column("#0", width=625, minwidth=625, stretch=True)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        # 移除水平滚动条
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        self.tree.bind("<Double-1>", self._on_tree_item_double_click)
        self.tree.bind("<Button-1>", self._on_tree_item_left_click)
        self.tree.bind("<Button-3>", self._on_tree_item_right_click)

        # 为"移动（重命名）到其它目录"特殊节点配置醒目样式
        self.tree.tag_configure(
            "move_action",
            foreground="dark blue",
            font=('TkDefaultFont', 9, 'bold'))

        # 加载文件夹图标，用于目录树每个节点前显示
        if getattr(sys, 'frozen', False):
            script_dir = sys._MEIPASS
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "folder.png")
        try:
            self.folder_icon = tk.PhotoImage(file=icon_path)
        except Exception:
            self.folder_icon = None

        # 节点ID到文件系统路径的映射字典（特殊节点不加入此映射）
        self.tree_path_map = {}
        # "移动（重命名）到其它目录"特殊节点的ID，每次重建目录树时更新
        self._move_other_node_id = None
        # 记录"移动到其它目录"对话框上次所选目录，作为下次对话框初值
        self._last_move_other_dir = None
    
    def _create_bottom_panel(self, parent):
        """创建底部操作面板（参照C#紧凑布局）"""
        # 定义文件信息区域样式（与"调整页面边距"控件相同的颜色与字体）
        style = ttk.Style()
        style.configure("FileInfo.TLabelframe", font=('TkDefaultFont', 9, 'bold'), foreground='#1976D2')
        style.configure("FileInfo.TLabelframe.Label", font=('TkDefaultFont', 9, 'bold'), foreground='#1976D2')
        style.configure("Bold.TLabel", font=('TkDefaultFont', 9, 'bold'))
        style.configure("NoBorder.TLabelframe", borderwidth=0, padding=0)
        style.configure("NoBorder.TLabelframe.Label", foreground="#FFF8DC")

        # 文件信息区域
        file_frame = ttk.LabelFrame(parent, text="文件信息", style="FileInfo.TLabelframe")
        file_frame.pack(fill=tk.X, padx=(2, 0), pady=2)

        # 第一行：当前文档路径
        ttk.Label(file_frame, text="当前文档:", style="Bold.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(2,5))
        self.lab_full_filename = ttk.Frame(file_frame)
        self.lab_full_filename.grid(row=0, column=1, columnspan=2, sticky=tk.EW, padx=(0,3), pady=3)
        self._lab_full_filename_label = ttk.Label(self.lab_full_filename, text="", anchor=tk.W)
        self._lab_full_filename_label.pack(fill=tk.X, padx=2, pady=0)
        file_frame.columnconfigure(1, weight=1)

        # 第二行：新文件名 + 使用标题作为文件名复选框 + 新文件名输入框
        ttk.Label(file_frame, text="新文件名:", style="Bold.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(2,5))
        self.chk_title_as_filename = tk.BooleanVar(value=True)
        ttk.Checkbutton(file_frame, text="使用标题作为文件名", variable=self.chk_title_as_filename).grid(
            row=1, column=1, sticky=tk.W, padx=0, pady=3)
        self.txt_new_filename = ttk.Entry(file_frame, width=20)
        self.txt_new_filename.grid(row=1, column=1, columnspan=3, sticky=tk.EW, padx=(2,3), pady=3)
        file_frame.columnconfigure(2, weight=1)
        
        # 按钮区域
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, padx=2, pady=(16, 16))
        
        btn_style = ttk.Style()
        btn_style.configure("BoldBtn.TButton", font=('TkDefaultFont', 11, 'bold'))
        
        # 主要操作按钮 - 使用grid布局实现水平分散分布
        self.btn_format_adjust = ttk.Button(btn_frame, text="调整文档格式", command=self._on_format_adjust, style="BoldBtn.TButton")
        self.btn_format_adjust.grid(row=0, column=0, padx=4, pady=2, ipady=3, sticky=tk.EW)
        
        self.btn_select_active_doc = ttk.Button(btn_frame, text="选择活动文档", command=self._on_select_active_doc, style="BoldBtn.TButton")
        self.btn_select_active_doc.grid(row=0, column=1, padx=4, pady=2, ipady=3, sticky=tk.EW)
        
        self.btn_save_close = ttk.Button(btn_frame, text="保存并关闭", command=self._on_save_close, style="BoldBtn.TButton")
        self.btn_save_close.grid(row=0, column=2, padx=4, pady=2, ipady=3, sticky=tk.EW)
        
        self.btn_rename = ttk.Button(btn_frame, text="重命名", command=self._on_rename, style="BoldBtn.TButton")
        self.btn_rename.grid(row=0, column=3, padx=4, pady=2, ipady=3, sticky=tk.EW)
        
        self.btn_filename_as_title = ttk.Button(btn_frame, text="文件名为标题", command=self._on_filename_as_title, style="BoldBtn.TButton")
        self.btn_filename_as_title.grid(row=0, column=4, padx=4, pady=2, ipady=3, sticky=tk.EW)
        
        self.btn_delete = ttk.Button(btn_frame, text="删除文件", command=self._on_delete, style="BoldBtn.TButton")
        self.btn_delete.grid(row=0, column=5, padx=4, pady=2, ipady=3, sticky=tk.EW)
        
        # 设置6列等宽分布
        for i in range(6):
            btn_frame.columnconfigure(i, weight=1)
        
        # 状态栏
        self.status_bar = ttk.Label(parent, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, padx=2, pady=(0,2))

    def _set_full_filename(self, text):
        """更新当前文档路径显示"""
        self._lab_full_filename_label.config(text=text)

    def _get_full_filename(self):
        """获取当前文档路径"""
        return self._lab_full_filename_label.cget("text")

    def _load_directory_tree(self, root_path, levels):
        """加载目录树结构"""
        # 清空Treeview中现有的所有节点
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 清空节点ID到文件系统路径的映射字典
        self.tree_path_map.clear()

        # 在最顶部插入特殊的"移动（重命名）到其它目录"节点
        # 该节点不对应任何实际目录，点击时会弹出目录选择对话框，
        # 便于将当前文档移动到目录树之外的任意目录
        move_kw = dict(text="移动（重命名）到其它目录", open=False, tags=("move_action",))
        if self.folder_icon:
            move_kw["image"] = self.folder_icon
        self._move_other_node_id = self.tree.insert("", 0, **move_kw)

        # 检查根路径是否存在，不存在则直接返回（特殊节点仍保留显示）
        if not os.path.exists(root_path):
            return
        
        # 定义递归添加目录项的内部函数
        def add_items(parent, path, depth):
            # 如果当前深度超过指定的层级限制，则停止递归
            if depth > levels:
                return
            
            # 获取当前目录下的所有文件和文件夹名称，按字母排序
            try:
                items = sorted(os.listdir(path))
            except PermissionError:
                # 遇到权限问题时跳过该目录
                return
            
            # 只收集文件夹，不显示文件
            folders = []
            for item in items:
                full_path = os.path.join(path, item)
                if os.path.isdir(full_path):
                    folders.append(item)
            
            # 遍历每个文件夹，添加到目录树中
            for folder in folders:
                full_path = os.path.join(path, folder)
                # 根据当前深度决定是否自动展开节点（未达到最大层级时自动展开）
                # 插入节点时显示文件夹图标
                kw = dict(text=folder, open=(depth < levels))
                if self.folder_icon:
                    kw["image"] = self.folder_icon
                node = self.tree.insert(parent, tk.END, **kw)
                # 将节点ID与完整路径关联存储
                self.tree_path_map[node] = full_path
                # 递归添加当前文件夹的子目录，深度加1
                add_items(node, full_path, depth + 1)
        
        # 创建根节点，显示根目录的名称，并默认展开
        # 创建根节点，显示根目录的名称，并默认展开（带文件夹图标）
        kw = dict(text=os.path.basename(root_path), open=True)
        if self.folder_icon:
            kw["image"] = self.folder_icon
        root_node = self.tree.insert("", tk.END, **kw)
        # 将根节点ID与根路径关联存储
        self.tree_path_map[root_node] = root_path
        # 从深度1开始递归添加根目录下的所有子目录和文件
        add_items(root_node, root_path, 1)
    
    def _on_catalogue_level_changed(self, event):
        """目录级数选择变更事件处理"""
        # 获取用户在下拉框中选择的目录级数文本
        level_text = self.cmb_catalogue_level.get()
        # 创建文本到数字的映射字典
        levels = {"显示1层目录": 1, "显示2层目录": 2, "显示3层目录": 3, "显示4层目录": 4, "显示5层目录": 5}
        # 更新当前级数并重新加载目录树
        self.current_level = levels.get(level_text, 5)
        self._load_directory_tree(self.current_root_path, self.current_level)
    
    def _on_select_directory(self):
        """选择目录按钮点击事件：打开文件夹选择对话框，更新根路径并加载目录树"""
        selected_path = filedialog.askdirectory(
            title="选择要显示的目录",
            initialdir=self.current_root_path)
        if selected_path:
            self.current_root_path = selected_path
            self._load_directory_tree(self.current_root_path, self.current_level)

    def _save_current_with_rename(self):
        """将当前打开的文档以新文件名保存（如果新文件名不为空则重命名），然后关闭当前文档"""
        if self.work_doc is None:
            return
        try:
            new_name = self.txt_new_filename.get().strip()
            if new_name:
                old_path = self.work_doc.FullName
                old_name = os.path.splitext(os.path.basename(old_path))[0]
                ext = os.path.splitext(old_path)[1]
                if new_name != old_name:
                    new_path = os.path.join(os.path.dirname(old_path), new_name + ext)
                    self.work_doc.SaveAs(new_path)
                    # 删除原文件
                    try:
                        os.remove(old_path)
                    except:
                        pass
                    # 重新打开新文件以保持文档引用
                    self.work_doc = self.word_app.Documents.Open(new_path)
                    self._set_full_filename(new_path)
                else:
                    self.work_doc.Save()
            else:
                self.work_doc.Save()
            # 关闭当前文档
            self.work_doc.Close(0)
            self.work_doc = None
            # 重置缓存
            self._last_first_para_text = None
            # 清空文件名输入框
            self._set_full_filename("")
            self.txt_new_filename.delete(0, tk.END)
        except Exception as ex:
            messagebox.showerror("错误", f"保存文档失败: {ex}")

    def _move_current_document(self, target_dir, close_after=True):
        """将当前文档移动（可同时重命名）到指定目录。
        close_after=True 时移动后关闭文档；False 时保持文档打开。
        此方法供目录树普通节点点击与"移动到其它目录"特殊节点共用。"""
        if self.work_doc is None:
            self.status_bar.config(text="请先选择一个文档")
            return
        try:
            old_path = self.work_doc.FullName
            new_name = self.txt_new_filename.get().strip()
            if not new_name:
                new_name = os.path.splitext(os.path.basename(old_path))[0]

            ext = os.path.splitext(old_path)[1]
            new_path = os.path.join(target_dir, new_name + ext)

            if new_path.lower() != old_path.lower():
                overwrite = False
                if os.path.exists(new_path):
                    overwrite = True

                self.work_doc.SaveAs(new_path)
                if close_after:
                    self.work_doc.Close(0)

                try:
                    os.remove(old_path)
                except Exception as delete_ex:
                    self.status_bar.config(text=f"新文件已保存，但删除原文件失败: {delete_ex}")
                    return

                if close_after:
                    self.work_doc = None
                    self._last_first_para_text = None
                    self._set_full_filename("")
                    self.txt_new_filename.delete(0, tk.END)
                    if overwrite:
                        self.status_bar.config(text=f"文件《{new_name + ext}》已移动并覆盖同名文件，文档已关闭")
                    else:
                        self.status_bar.config(text=f"文件《{new_name + ext}》已移动，文档已关闭")
                else:
                    self._set_full_filename(new_path)
                    dpm.word_app = self.word_app
                    dpm.work_doc = self.work_doc
                    self._suppress_monitor = True
                    self.root.after(3000, lambda: setattr(self, '_suppress_monitor', False))
                    if overwrite:
                        self.status_bar.config(text=f"文件《{new_name + ext}》已移动并覆盖同名文件，文档保持打开")
                    else:
                        self.status_bar.config(text=f"文件《{new_name + ext}》已移动，文档保持打开")
            else:
                # 目标路径与原路径相同，仅保存
                self.work_doc.Save()
                if close_after:
                    self.work_doc.Close(0)
                    self.work_doc = None
                    self._last_first_para_text = None
                    self._set_full_filename("")
                    self.txt_new_filename.delete(0, tk.END)
                    self.status_bar.config(text="文件已保存并关闭")
                else:
                    self.status_bar.config(text="文件已保存")
        except Exception as ex:
            self.status_bar.config(text=f"移动文档失败: {ex}")

    def _move_to_other_directory(self):
        """点击"移动（重命名）到其它目录"特殊节点时调用：
        弹出目录选择对话框（初值为上次调用此对话框所选目录），
        按右击目录树节点的相同逻辑（保持文档打开）将当前文档移动到选定目录。
        若当前没有正在编辑的文档则不响应点击。"""
        # 当前没有正在编辑的文档时不响应点击
        if self.work_doc is None:
            return
        # 初值优先使用上次调用此对话框所选目录，否则回退到当前根目录，再回退到用户主目录
        initial_dir = (self._last_move_other_dir
                       or self.current_root_path
                       or os.path.expanduser("~"))
        target_dir = filedialog.askdirectory(
            title="选择要移动到的目标目录",
            initialdir=initial_dir)
        if not target_dir:
            return
        # 记住本次所选目录，作为下次对话框的初值
        self._last_move_other_dir = target_dir
        # 与右击目录树节点相同：移动后保持文档打开
        self._move_current_document(target_dir, close_after=False)

    def _on_tree_item_left_click(self, event):
        """左键单击事件：延迟执行文档移动，以便与双击区分"""
        element = self.tree.identify_element(event.x, event.y)
        if element == 'indicator':
            return

        item = self.tree.identify('item', event.x, event.y)
        if not item:
            return

        if item == self._move_other_node_id:
            self._move_to_other_directory()
            return

        if item not in self.tree_path_map:
            return

        path = self.tree_path_map[item]
        if not os.path.isdir(path):
            return

        if self._click_timer:
            self.root.after_cancel(self._click_timer)

        def _delayed_move():
            self._move_current_document(path, close_after=True)

        self._click_timer = self.root.after(250, _delayed_move)

    def _on_tree_item_double_click(self, event):
        """目录树双击事件处理：在资源管理器中打开对应文件夹，保持文档打开状态"""
        if self._click_timer:
            self.root.after_cancel(self._click_timer)
            self._click_timer = None

        item = self.tree.identify("item", event.x, event.y)
        if not item or item not in self.tree_path_map:
            return

        # 目录树节点均为目录，使用os.path.normpath确保Windows反斜杠路径
        target = os.path.normpath(self.tree_path_map[item])

        def _wait_and_reload():
            try:
                # 使用explorer.exe完整路径，不使用shell=True，确保路径解析正确
                proc = subprocess.Popen(['explorer.exe', target])
                proc.wait()
                self.root.after(0, lambda: self._load_directory_tree(self.current_root_path, self.current_level))
            except Exception as ex:
                self.root.after(0, lambda: messagebox.showerror("错误", f"打开文件夹失败: {ex}"))

        t = threading.Thread(target=_wait_and_reload, daemon=True)
        t.start()

    def _on_tree_item_right_click(self, event):
        """右键单击事件：将当前文档移动到所选文件夹但不关闭"""
        item = self.tree.identify('item', event.x, event.y)
        if not item:
            return

        # 命中"移动（重命名）到其它目录"特殊节点时，弹出目录选择对话框（保持文档打开）
        if item == self._move_other_node_id:
            self._move_to_other_directory()
            return

        if item not in self.tree_path_map:
            return

        path = self.tree_path_map[item]
        if not os.path.isdir(path):
            return

        self._move_current_document(path, close_after=False)
    
    def _open_document(self, path, close_current=True):
        """打开指定路径的Word文档"""
        try:
            # 如果Word应用对象尚未连接，则尝试获取系统中正在运行的Word应用
            if self.word_app is None:
                self.word_app = Universal().get_active_word_app()
            
            # 如果未找到Word应用，显示警告对话框并返回
            if self.word_app is None:
                messagebox.showwarning("警告", "未找到正在运行的Word应用")
                return
            
            # 如果需要关闭当前打开的文档（左键双击时）
            if close_current and self.work_doc is not None:
                try:
                    # 关闭当前文档，参数0表示不保存
                    self.work_doc.Close(0)
                except:
                    # 忽略关闭异常
                    pass
            
            # 使用Word应用打开指定路径的文档
            self.work_doc = self.word_app.Documents.Open(path)
            # 更新界面上"当前文档"输入框的显示内容
            self._set_full_filename(self.work_doc.FullName)
            
            # 更新全局变量中的Word应用和文档对象，供其他模块使用
            dpm.word_app = self.word_app
            dpm.work_doc = self.work_doc
            
            # 根据文档第一段落内容自动更新新文件名输入框
            self._update_filename_from_first_paragraph()
        except Exception as ex:
            # 捕获并显示打开文档过程中的错误
            messagebox.showerror("错误", f"打开文档失败: {ex}")
    
    def _on_set_all_margins_to_1cm(self):
        """将所有页面边距设置为1厘米"""
        # 将上边距输入框设置为1.0厘米
        self.top_margin.delete(0, tk.END)
        self.top_margin.insert(0, "1.0")
        # 将下边距输入框设置为1.0厘米
        self.bottom_margin.delete(0, tk.END)
        self.bottom_margin.insert(0, "1.0")
        # 将左边距输入框设置为1.0厘米
        self.left_margin.delete(0, tk.END)
        self.left_margin.insert(0, "1.0")
        # 将右边距输入框设置为1.0厘米
        self.right_margin.delete(0, tk.END)
        self.right_margin.insert(0, "1.0")
    
    def _on_rename(self):
        """重命名当前打开的文档（保持在原目录）"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        old_path = self.work_doc.FullName
        target_dir = os.path.dirname(old_path)
        self._move_current_document(target_dir, close_after=False)
    
    def _on_filename_as_title(self):
        """将文件名主名作为标题插入文档第一段前，并按总标题格式设置"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        try:
            filename = os.path.splitext(os.path.basename(self.work_doc.FullName))[0]
            
            self.work_doc.Paragraphs(1).Range.InsertBefore(filename + '\r')
            
            set_doc_format = set_document_format.SetDocumentFormat(doc=self.work_doc)
            set_doc_format.set_main_title_format(
                self.cmb_main_title_font.get(),
                self.cmb_main_title_font_size.get(),
                self.chk_main_title_bold.get())
        except Exception as ex:
            messagebox.showerror("错误", f"操作失败: {ex}")
    
    def _apply_saved_window_bounds(self):
        """应用上次保存的窗口位置（不恢复尺寸）"""
        # 仅恢复位置，不恢复尺寸（尺寸固定为738x836）
        if dpm.WindowWidth > 0 and dpm.WindowHeight > 0:
            self.root.geometry(f"738x836+{dpm.WindowLeft}+{dpm.WindowTop}")
    
    def _initialize_event_handlers(self):
        """初始化事件处理器（预留方法，当前为空实现）"""
        pass
    
    def _check_word_app(self):
        """检查系统中是否有运行的Word应用，并连接获取活动文档"""
        # 获取系统中正在运行的Word应用对象
        self.word_app = Universal().get_active_word_app()
        if self.word_app is not None:
            try:
                # 检查Word中是否有打开的文档
                if self.word_app.Documents.Count > 0:
                    # 将活动文档设置为当前工作文档
                    self.work_doc = self.word_app.ActiveDocument
                    # 更新界面上"当前文档"输入框的显示
                    self._set_full_filename(self.work_doc.FullName)
                    # 更新全局变量中的Word应用和文档对象
                    dpm.word_app = self.word_app
                    dpm.work_doc = self.work_doc
                    # 根据文档第一段落自动更新新文件名输入框
                    self._update_filename_from_first_paragraph()
            except:
                # 忽略连接过程中的异常
                pass
    
    def _on_format_adjust(self):
        """调整文档格式：以Word当前激活的文档为目标，按界面参数调整格式"""
        # 确保Word应用已连接
        if self.word_app is None:
            self.word_app = Universal().get_active_word_app()
        if self.word_app is None:
            messagebox.showwarning("警告", "未找到正在运行的Word应用")
            return
        
        try:
            self.status_bar.config(text="")
            
            active_doc = self.word_app.ActiveDocument
            if active_doc is None:
                messagebox.showwarning("警告", "Word中没有打开的文档")
                return
            
            self.work_doc = active_doc
            self._set_full_filename(self.work_doc.FullName)
            dpm.word_app = self.word_app
            dpm.work_doc = self.work_doc
            self._update_filename_from_first_paragraph()
            
            # 执行格式调整
            from set_document_format import SetDocumentFormat
            from page_number_manager import PageNumberManager
            
            set_doc_format = SetDocumentFormat(doc=self.work_doc)
            
            if self.chk_change_page_margin.get():
                set_doc_format.set_page_margins(
                    float(self.top_margin.get()),
                    float(self.bottom_margin.get()),
                    float(self.left_margin.get()),
                    float(self.right_margin.get()))
            
            if self.chk_add_page_num.get():
                PageNumberManager.add_page_numbers_custom(self.work_doc)
            
            if self.chk_change_main_title_format.get():
                set_doc_format.set_main_title_format(
                    self.cmb_main_title_font.get(),
                    self.cmb_main_title_font_size.get(),
                    self.chk_main_title_bold.get())
            
            if self.chk_change_image_and_table_format.get():
                set_doc_format.set_images_and_tables()
            
            if self.chk_change_content_format.get():
                indent_style = self.cmb_content_indent.get()
                alignment = self.cmb_content_align.get()
                font = self.cmb_content_font.get()
                font_size = self.cmb_content_font_size.get()
                if font == "不变更":
                    font = None
                if indent_style == "不变更":
                    indent_style = None
                if alignment == "不变更":
                    alignment = None
                set_doc_format.set_content_format(
                    indent_style, alignment, font, font_size,
                    None,
                    self.chk_delete_empty_lines.get(),
                    self.chk_standard_line_spacing.get())
            
            if self.chk_change_level_title_format.get():
                self._set_level_title_styles(set_doc_format)
            
            filename = self.txt_new_filename.get().strip()
            if not filename:
                filename = os.path.basename(self.work_doc.FullName)
            self.status_bar.config(text=f"《{filename}》调整文档格式完成")
            
        except Exception as ex:
            messagebox.showerror("错误", f"调整格式失败: {ex}")
    
    def _set_level_title_styles(self, set_doc_format):
        """设置章节标题样式，遍历文档所有段落并应用对应级别的标题样式"""
        # 初始化存储1-5级标题样式对象的列表，初始值为None表示尚未创建
        level_styles = [None, None, None, None, None]
        
        try:
            # 获取文档对象，优先使用格式设置对象中的文档，其次使用当前工作文档
            doc = set_doc_format.work_doc
            if doc is None:
                doc = self.work_doc
            
            # 如果文档对象仍为None，则直接返回
            if doc is None:
                return
            
            # 获取文档中的段落总数
            para_count = doc.Paragraphs.Count
            
            # 遍历文档中的每一个段落（从1开始，Word对象索引从1开始）
            for para_index in range(1, para_count + 1):
                try:
                    # 获取当前段落对象
                    paragraph = doc.Paragraphs(para_index)
                    # 获取段落的大纲级别（1-9级，0表示正文）
                    level_index = set_doc_format.get_paragraph_outline_level(paragraph)
                    
                    # 只处理1-5级标题，忽略正文和其他级别
                    if level_index < 1 or level_index > 5:
                        continue
                    
                    # 如果该级别的样式尚未创建
                    if level_styles[level_index - 1] is None:
                        # 获取该级别标题的各项格式设置
                        font = self.level_title_font_vars[level_index - 1].get()
                        font_size = self.level_title_font_size_vars[level_index - 1].get()
                        number_style = self.level_title_number_vars[level_index - 1].get()
                        indent = self.level_title_indent_vars[level_index - 1].get()
                        
                        # 将"不变更"选项转换为None，表示保持原有设置
                        if font == "不变更":
                            font = None
                        if font_size == "不变更":
                            font_size = None
                        if number_style == "不变更":
                            number_style = None
                        if indent == "不变更":
                            indent = None
                        
                        # 如果有至少一项属性需要修改，则创建该级别的标题样式
                        if font is not None or font_size is not None or indent is not None or number_style is not None:
                            level_styles[level_index - 1] = set_doc_format.set_title_styles(
                                level_index, font, font_size, number_style, indent)
                    
                    if level_styles[level_index - 1] is not None:
                        set_range_style(paragraph.Range, level_styles[level_index - 1])
                except:
                    # 忽略单个段落处理时的异常，继续处理下一个段落
                    continue
        except:
            # 忽略整体处理过程中的异常
            pass
    
    def _on_select_active_doc(self):
        """选择当前活动的Word文档"""
        try:
            # 获取活动Word应用
            self.word_app = Universal().get_active_word_app()
            if self.word_app is None:
                messagebox.showwarning("警告", "未找到正在运行的Word应用")
                return
            
            # 检查是否有打开的文档
            if self.word_app.Documents.Count < 1:
                messagebox.showwarning("警告", "请先打开一个文档")
                return
            
            # 设置当前活动文档为工作文档
            self.work_doc = self.word_app.ActiveDocument
            self._set_full_filename(self.work_doc.FullName)
            
            # 更新全局变量
            dpm.word_app = self.word_app
            dpm.work_doc = self.work_doc
            
            # 根据文档第一段落更新文件名
            self._update_filename_from_first_paragraph()
        except Exception as ex:
            messagebox.showerror("错误", f"选择文档失败: {ex}")
    
    def _update_filename_from_first_paragraph(self):
        """根据文档第一段落更新新文件名输入框"""
        if self.work_doc is None:
            return
        
        try:
            # 获取第一段内容
            first_para = self.work_doc.Paragraphs(1)
            title = first_para.Range.Text.strip()
            
            if title:
                # 更新缓存
                self._last_first_para_text = title
                # 移除非法文件名字符
                invalid_chars = '<>:"/\\|?*'
                title = ''.join(c for c in title if c not in invalid_chars)
                # 更新新文件名输入框
                self.txt_new_filename.delete(0, tk.END)
                self.txt_new_filename.insert(0, title)
        except:
            pass
    
    def _on_save_close(self):
        """保存文档并关闭"""
        if self.work_doc is None:
            self.status_bar.config(text="请先选择一个文档")
            return
        
        try:
            old_path = self.work_doc.FullName
            new_name = self.txt_new_filename.get().strip()
            need_delete_old = False
            
            if new_name:
                old_name = os.path.splitext(os.path.basename(old_path))[0]
                ext = os.path.splitext(old_path)[1]
                new_path = os.path.join(os.path.dirname(old_path), new_name + ext)
                
                if new_name.lower() != old_name.lower():
                    self.work_doc.SaveAs(new_path)
                    need_delete_old = True
                else:
                    self.work_doc.Save()
            else:
                self.work_doc.Save()
            
            self.work_doc.Close(0)
            self.work_doc = None
            
            delete_success = True
            if need_delete_old:
                try:
                    os.remove(old_path)
                except Exception as delete_ex:
                    delete_success = False
                    self.status_bar.config(text=f"新文件已保存，但删除原文件失败: {delete_ex}")
            
            self._last_first_para_text = None
            self._set_full_filename("")
            self.txt_new_filename.delete(0, tk.END)
            if delete_success:
                self.status_bar.config(text="保存成功")
        except Exception as ex:
            self.status_bar.config(text=f"保存失败: {ex}")
    
    def _on_delete(self):
        """删除当前文档"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        filename = self.work_doc.FullName
        
        try:
            self.work_doc.Close(0)
            os.remove(filename)
            self._set_full_filename("未选择文件")
            self.work_doc = None
            self._last_first_para_text = None
            self.status_bar.config(text="删除成功")
        except Exception as ex:
            self.status_bar.config(text=f"删除失败: {ex}")
    
    def _on_cancel_settings(self):
        """取消所有设置"""
        self.chk_change_page_margin.set(False)
        self.chk_change_content_format.set(False)
        self.chk_change_main_title_format.set(False)
        self.chk_change_level_title_format.set(False)
        self.chk_change_page_orientation.set(False)
        self.chk_change_image_and_table_format.set(False)
    
    def _on_enable_settings(self):
        """使能所有设置（变更纸张方向除外）"""
        self.chk_change_page_margin.set(True)
        self.chk_change_content_format.set(True)
        self.chk_change_main_title_format.set(True)
        self.chk_change_level_title_format.set(True)
        self.chk_change_page_orientation.set(False)
        self.chk_change_image_and_table_format.set(True)
    
    def _on_title_level_up(self):
        """所有标题升一级"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        try:
            from win32com.client import constants as wc

            # 预检：若最高一级标题已为第1级，则不再执行升一级逻辑并提示
            for i in range(1, self.work_doc.Paragraphs.Count + 1):
                para = self.work_doc.Paragraphs(i)
                style = get_range_style(para)
                if style.NameLocal == "标题 1":
                    messagebox.showinfo("提示", "最高一级弹窗已升到第1级，不能再升了！")
                    return

            # 遍历所有段落
            for i in range(1, self.work_doc.Paragraphs.Count + 1):
                para = self.work_doc.Paragraphs(i)
                style = get_range_style(para)
                style_name = style.NameLocal

                # 根据当前样式名确定新样式
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

            self.status_bar.config(text="所有标题已升一级")
        except Exception as ex:
            messagebox.showerror("错误", f"操作失败: {ex}")

    def _on_title_level_down(self):
        """所有标题降一级"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        try:
            from win32com.client import constants as wc
            
            # 遍历所有段落
            for i in range(1, self.work_doc.Paragraphs.Count + 1):
                para = self.work_doc.Paragraphs(i)
                style = get_range_style(para)
                style_name = style.NameLocal
                
                # 根据当前样式名确定新样式
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
            
            self.status_bar.config(text="所有标题已降一级")
        except Exception as ex:
            messagebox.showerror("错误", f"操作失败: {ex}")
    
    def _on_delete_all_pictures(self):
        """删除文档中所有图片"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        # 确认删除
        if messagebox.askyesno("确认", "确定要删除所有图片吗？"):
            try:
                from set_document_format import SetDocumentFormat
                set_doc_format = SetDocumentFormat(doc=self.work_doc)
                set_doc_format.delete_all_pictures()
                messagebox.showinfo("完成", "所有图片已删除")
            except Exception as ex:
                messagebox.showerror("错误", f"删除失败: {ex}")
    
    def _on_cancel_wrap_as_inline(self):
        """取消图片和表格的嵌入型环绕方式"""
        if self.work_doc is None:
            messagebox.showwarning("警告", "请先选择一个文档")
            return
        
        try:
            from set_document_format import SetDocumentFormat
            set_doc_format = SetDocumentFormat(doc=self.work_doc)
            set_doc_format.cancel_wrap_as_inline()
            messagebox.showinfo("完成", "已取消嵌入型环绕")
        except Exception as ex:
            messagebox.showerror("错误", f"操作失败: {ex}")

    def _on_window_configure(self, event):
        """窗口尺寸变化时更新状态栏显示当前宽度和高度"""
        # 只响应根窗口的尺寸变化，避免子控件事件干扰
        if event.widget == self.root:
            current_text = self.status_bar.cget("text")
            if current_text and any(keyword in current_text for keyword in ["移动", "覆盖", "关闭", "保存", "失败", "成功"]):
                return
            width = event.width
            height = event.height
            self.status_bar.config(text=f"宽度: {width} x 高度: {height}")

    def _monitor_active_document(self):
        """定时检查 Word 的活动文档是否改变，如改变则更新界面显示"""
        try:
            if self._suppress_monitor:
                return
            
            # 如果 word_app 存在，检查其活动文档
            if self.word_app is not None:
                try:
                    active_doc = self.word_app.ActiveDocument
                    if active_doc is not None:
                        # 如果当前 work_doc 与活动文档不同，则更新
                        if self.work_doc is None or self.work_doc.FullName != active_doc.FullName:
                            self.work_doc = active_doc
                            # 更新当前文档路径
                            self._set_full_filename(self.work_doc.FullName)
                            # 更新全局变量
                            dpm.word_app = self.word_app
                            dpm.work_doc = self.work_doc
                            # 重置缓存，强制更新文件名
                            self._last_first_para_text = None
                            # 更新新文件名
                            self._update_filename_from_first_paragraph()
                        else:
                            # 文档不变，检查第一段内容是否变化
                            if self.chk_title_as_filename.get():
                                try:
                                    current_text = self.work_doc.Paragraphs(1).Range.Text.strip()
                                    if current_text != self._last_first_para_text:
                                        self._last_first_para_text = current_text
                                        self._update_filename_from_first_paragraph()
                                except Exception:
                                    pass
                except Exception:
                    # 如果无法获取活动文档（如 Word 关闭），忽略
                    pass
        except Exception:
            pass
        finally:
            # 每 1000 毫秒（1秒）再次检查
            self.root.after(1000, self._monitor_active_document)

    def _on_closing(self):
        """窗口关闭事件处理"""
        # 保存窗口位置和大小
        dpm.WindowWidth = self.root.winfo_width()
        dpm.WindowHeight = self.root.winfo_height()
        dpm.WindowLeft = self.root.winfo_x()
        dpm.WindowTop = self.root.winfo_y()
        # 保存当前根目录到配置
        dpm.WORK_FOLDER = self.current_root_path
        # 读取当前窗体设置到全局变量
        dpm.read_from_form(self)
        # 保存配置文件
        dpm.save_settings()
        
        # 销毁窗口
        self.root.destroy()
