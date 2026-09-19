wdHeaderFooterPrimary = 1
wdHeaderFooterFirstPage = 2
wdHeaderFooterEvenPages = 3

wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdAlignParagraphRight = 2
wdAlignParagraphJustify = 3
wdAlignParagraphDistribute = 4

wdAlignPageNumberLeft = 1
wdAlignPageNumberCenter = 2
wdAlignPageNumberRight = 3

wdPageNumberStyleArabic = 1

wdFieldPage = 33
wdFieldNumPages = 26

wdNoProtection = 0

wdStatisticPages = 2

wdLineSpaceSingle = 0
wdLineSpace1pt5 = 1
wdLineSpaceDouble = 2
wdLineSpaceAtLeast = 3
wdLineSpaceExactly = 4
wdLineSpaceMultiple = 5

wdInlineShapePicture = 3

wdStyleTypeParagraph = 1

wdOutlineNumberGallery = 2
wdListNumberStyleArabic = 0

wdWrapSquare = 0
wdWrapTopBottom = 3
wdWrapInline = 7

wdAutoFitFixed = 0
wdAutoFitContent = 1
wdAutoFitWindow = 2

# 单元格垂直对齐方式
wdCellAlignVerticalTop = 0
wdCellAlignVerticalCenter = 1
wdCellAlignVerticalBottom = 3

msoAutomationSecurityLow = 1

wdFirstCharacterLineNumber = 10
wdWithInTable = 12
wdRelativeHorizontalPositionPage = 0
wdShapeCenter = -999994

wdLineStyleSingle = 1
wdLineWidth025pt = 2
wdLineWidth075pt = 8
wdColorBlack = 0

wdBorderTop = -1
wdBorderLeft = -2
wdBorderBottom = -3
wdBorderRight = -4

wdListNumberStyleSimpChinNum = 37
wdListNumberStyleNumberInCircle = 18
wdListNumberStyleNone = 255
# 注：wdListNumberStyleArabic 已在文件上方（第36行）定义为 0

# ============================================================
# 编号样式（WdListNumberStyle）——"转自动序号"保持原形式所需
# 转换后的自动编号应与转换前的手动序号形式一致，因此各序号体系
# 必须映射到各自的 NumberStyle，而不是统一成阿拉伯数字。
# ============================================================
wdListNumberStyleUppercaseRoman = 1    # I, II, III
wdListNumberStyleLowercaseRoman = 2    # i, ii, iii
wdListNumberStyleUppercaseLetter = 3   # A, B, C
wdListNumberStyleLowercaseLetter = 4   # a, b, c
wdListNumberStyleArabicFullWidth = 14  # １, ２, ３（全角阿拉伯数字）
wdListNumberStyleUppercaseGreek = 61   # Α, Β, Γ
wdListNumberStyleLowercaseGreek = 60   # α, β, γ
# 天干 / 地支序号（甲、乙、丙… / 子、丑、寅…）
wdListNumberStyleZodiac1 = 30
wdListNumberStyleZodiac2 = 31
wdListNumberStyleZodiac3 = 32
# 简体中文数字：1=一二三(37)，2=壹贰叁(38)
wdListNumberStyleSimpChinNum1 = 37
wdListNumberStyleSimpChinNum2 = 38

wdTrailingTab = 0
wdTrailingSpace = 1
wdTrailingNone = 2


def set_range_style(range_obj, style):
    try:
        range_obj.set_Style(style)
    except AttributeError:
        range_obj.Style = style


def get_range_information(range_obj, info_type):
    try:
        return range_obj.Information[info_type]
    except TypeError:
        return range_obj.Information(info_type)


def get_range_style(range_obj):
    try:
        return range_obj.get_Style()
    except AttributeError:
        return range_obj.Style


def get_list_gallery(app_obj, gallery_type):
    try:
        return app_obj.ListGalleries[gallery_type]
    except TypeError:
        return app_obj.ListGalleries(gallery_type)


# ============================================================
# 字体属性读取 / 写回辅助
# 用于"不变更"的字体、字号保护：把段落改套"标题 N"样式时，
# 若段落本身没有直接格式，有效字体与字号会随样式一起改变
# （例如段落原用"标题 1"，改套"标题 2"后字号随之变化）。
# 做法是改样式前读取段落的有效值，改样式后再以直接格式写回——
# 直接格式优先于样式，因此处理前的字体与字号得以保留。
# ============================================================

def get_effective_font_size(range_obj):
    """读取 Range 的有效字号（磅）。

    区域内字号不一致时 Word 返回 wdUndefined(9999999)，对象不支持该属性
    时取值会抛异常；两种情况都返回 None，由调用方跳过写回，避免用无效值
    覆盖段落字号。
    """
    try:
        size = float(range_obj.Font.Size)
    except Exception:
        return None
    if size <= 0 or size >= 1000:
        return None
    return size


def get_effective_font_names(range_obj):
    """读取 Range 的有效字体名，返回 {属性名: 字体名}。

    Word 把字体分槽位保存：Name/NameAscii 为西文，NameFarEast 为中文，
    NameOther 为其它文种。全部读取以便原样写回；取不到、为空或不支持的
    槽位直接跳过（兼容 WPS）。
    """
    names = {}
    for attr in ('Name', 'NameFarEast', 'NameAscii', 'NameOther'):
        try:
            value = getattr(range_obj.Font, attr)
        except Exception:
            continue
        if isinstance(value, str) and value:
            names[attr] = value
    return names


def apply_font_size(range_obj, size):
    """以直接格式把字号写回 Range；size 为 None 时不做任何操作。"""
    if size is None:
        return
    try:
        range_obj.Font.Size = size
    except Exception:
        pass


def apply_font_names(range_obj, names):
    """以直接格式把字体名写回 Range；names 为空时不做任何操作。"""
    if not names:
        return
    for attr, value in names.items():
        try:
            setattr(range_obj.Font, attr, value)
        except Exception:
            pass


# ============================================================
# 段落缩进读取 / 写回辅助
# 与字体同理：应用自动编号列表模板时，Word 会把列表级别的
# NumberPosition/TextPosition 作为段落缩进写到段落上，从而改掉
# 标题原有的缩进方式（如首行缩进2字符被清零）。
# 做法同样是先读后写：处理前记录，全部操作结束后以直接格式写回。
# ============================================================

# 缩进属性：点值与字符单位值并存（"首行缩进2字符"存的是字符单位值），
# 必须成对保存与还原，否则会出现"点值为0、字符单位仍为2"的矛盾状态。
_INDENT_ATTRS = ('LeftIndent', 'RightIndent', 'FirstLineIndent',
                 'CharacterUnitLeftIndent', 'CharacterUnitRightIndent',
                 'CharacterUnitFirstLineIndent')


def get_effective_indents(range_obj):
    """读取 Range 的有效缩进属性，返回 {属性名: 值}。

    值为 wdUndefined(9999999，表示继承或未设置) 或取值失败时跳过该项，
    避免把无效值写回段落。
    """
    indents = {}
    for attr in _INDENT_ATTRS:
        try:
            value = float(getattr(range_obj.ParagraphFormat, attr))
        except Exception:
            continue
        if abs(value) >= 1000:      # wdUndefined 等无效值
            continue
        indents[attr] = value
    return indents


def apply_indents(range_obj, indents):
    """把缩进属性以直接格式写回 Range；indents 为空时不做任何操作。

    先写点值再写字符单位值，保证"首行缩进N字符"这类设置最终生效。
    """
    if not indents:
        return
    for attr in _INDENT_ATTRS:
        if attr not in indents:
            continue
        try:
            setattr(range_obj.ParagraphFormat, attr, indents[attr])
        except Exception:
            pass


def clear_indents(range_obj):
    """把 Range 所在段落的缩进属性全部清零（直接格式）。

    用于"图片段落必须不缩进"：只把 FirstLineIndent 置 0 是不够的——
      - "悬挂缩进"在 Word 里是 LeftIndent>0 且 FirstLineIndent<0，
        只清首行负值会剩下一段左缩进，图片看上去依然被顶偏；
      - 中文文档的"缩进2字符"存在字符单位属性里，点值清零并不能
        保证字符单位值同步归零（二者是独立的两个属性）。
    因此 6 个缩进属性成对清零，先点值后字符单位值。

    与 apply_indents 相反，这里是"无条件写 0"，不做取值有效性判断。
    """
    for attr in _INDENT_ATTRS:
        try:
            setattr(range_obj.ParagraphFormat, attr, 0)
        except Exception:
            pass


# 每字符对应的磅值：沿用"变更正文格式"原有算法（0.35厘米 × 28.35），
# 保证标题与正文在同样选项下缩进一致。
_PT_PER_CHAR = 0.35 * 28.35

# "缩进方式" → (左缩进磅值, 左缩进字符数, 首行缩进磅值, 首行缩进字符数)
_INDENT_STYLE_MAP = {
    "首行缩进2字符": (0.0, 0, 2 * _PT_PER_CHAR, 2),
    "无缩进":        (0.0, 0, 0.0, 0),
    "悬挂缩进":      (2 * _PT_PER_CHAR, 2, -2 * _PT_PER_CHAR, -2),
}


def apply_indent_style(paragraph_format, indent_style):
    """按界面的"缩进方式"设置段落缩进。

    点值与字符单位值同时写入：字符单位值让中文 Word 显示为"N字符"，
    点值作为对象不支持字符单位属性时的兜底。左缩进一并设置——带自动
    编号的标题会被列表级别的 NumberPosition 顶出左缩进，不归零的话
    选"无缩进"看起来等于没生效。

    Args:
        paragraph_format: 段落的 ParagraphFormat（可直接传 Range.ParagraphFormat）
        indent_style: 界面上的缩进方式文本

    Returns:
        bool: True 表示已按该选项设置；False 表示取值未识别，未做任何修改。
    """
    if indent_style not in _INDENT_STYLE_MAP:
        return False

    left_pt, left_ch, first_pt, first_ch = _INDENT_STYLE_MAP[indent_style]

    # 先写点值（兜底），再写字符单位值（中文 Word 下最终生效）
    for attr, value in (('LeftIndent', left_pt),
                        ('FirstLineIndent', first_pt),
                        ('CharacterUnitLeftIndent', left_ch),
                        ('CharacterUnitFirstLineIndent', first_ch)):
        try:
            setattr(paragraph_format, attr, value)
        except Exception:
            pass
    return True