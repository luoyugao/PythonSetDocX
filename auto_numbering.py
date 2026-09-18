"""
auto_numbering.py
标题手动序号转Word自动序号模块

功能：
  遍历文档中大纲级别为1~9的标题段落，检测人工输入的手动序号文本
  （如"1."、"一、"、"（一）"、"①"、"A."、"α."、"Ⅰ."、"甲、"等），
  将其从段落文本中移除，并链接Word自动多级编号列表模板，实现
  标题自动编号。

支持的序号体系：
  - 阿拉伯数字（含全角）：1, 2, 3… / １, ２, ３…
  - 中文小写数字：一, 二, 三…十二, 二十, 一百…
  - 中文大写数字：壹, 贰, 叁…（暂不处理）
  - 天干：甲, 乙, 丙, 丁, 戊, 己, 庚, 辛, 壬, 癸
  - 地支：子, 丑, 寅, 卯, 辰, 巳, 午, 未, 申, 酉, 戌, 亥
  - 英文大写字母（含全角）：A, B, C… / Ａ, Ｂ, Ｃ…
  - 英文小写字母（含全角）：a, b, c… / ａ, ｂ, ｃ…
  - 罗马数字（大写/小写）：Ⅰ, Ⅱ, Ⅲ… / ⅰ, ⅱ, ⅲ…
  - 希腊字母（大写/小写）：Α, Β, Γ… / α, β, γ…
  - 圈号数字：①, ②, ③…⑳, ㉑…㊿
  - 圈号字母：Ⓐ, Ⓑ… / ⓐ, ⓑ…
  - 资料1.  资料2.  等特殊前缀
  - 步骤1.  步骤2.  等特殊前缀
  - 章序号：第1章 第一章 等
  - 多级序号：3.1  3.1.2  １．２  等（各级计数器由Word级联生成）

支持的分隔符（含全角形式）：
  - 点号：. ．
  - 逗号：, ，
  - 顿号：、
  - 右括号：) ）
  - 括号包围：（一） (1) 等
"""

import collections
import re
import word_constants as wc

# ============================================================
# 全角 → 半角 字符映射表
# ============================================================
_FULL_TO_HALF_MAP = str.maketrans(
    '０１２３４５６７８９'
    'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
    'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
    '．，））（'
    '。',
    '0123456789'
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    'abcdefghijklmnopqrstuvwxyz'
    '.,))('
    '.'
)

# ============================================================
# 数字字符集（用于判断字符类别）
# ============================================================
_ARABIC_HALF = set('0123456789')
_ARABIC_FULL_SET = set('０１２３４５６７８９')
_CHINESE_NUMS = set('零〇一二三四五六七八九十百千万亿两')
_HEAVENLY_STEMS = set('甲乙丙丁戊己庚辛壬癸')
_EARTHLY_BRANCHES = set('子丑寅卯辰巳午未申酉戌亥')
_UPPER_ALPHA_HALF = set('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
_UPPER_ALPHA_FULL_SET = set('ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ')
_LOWER_ALPHA_HALF = set('abcdefghijklmnopqrstuvwxyz')
_LOWER_ALPHA_FULL_SET = set('ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ')
_ROMAN_UPPER = set('ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫⅬⅭⅮⅯ')
_ROMAN_LOWER = set('ⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅺⅻⅼⅽⅾⅿ')
_GREEK_UPPER = set('ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ')
_GREEK_LOWER = set('αβγδεζηθικλμνξοπρστυφχψω')
_CIRCLED_NUMS = set('①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿')
_CIRCLED_LOWER_ALPHA = set('ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝⓞⓟⓠⓡⓢⓣⓤⓥⓦⓧⓨⓩ')
_CIRCLED_UPPER_ALPHA = set('ⒶⒷⒸⒹⒺⒻⒼⒽⒾⒿⓀⓁⓂⓃⓄⓅⓆⓇⓈⓉⓊⓋⓌⓍⓎⓏ')

# 分隔符（半角 + 全角）
_SEP_DOT    = '.．。'        # 点号（含全角句号作点号用）
_SEP_COMMA  = ',，'          # 逗号
_SEP_DUNHAO = '、'           # 中文顿号
_SEP_PAREN_R = ')）'         # 右括号
_SEP_PAREN_L = '(（'         # 左括号

# 成对括号哨兵：_detect_paren_form 匹配到"(1)"/"（一）"这类左右成对的
# 形式时传入，与"1)"这类仅右括号收尾的形式区分开。
_SEP_PAREN_PAIRED = '_PAIRED_'

# 所有分隔符合集
_ALL_SEPS = _SEP_DOT + _SEP_COMMA + _SEP_DUNHAO + _SEP_PAREN_R

# convert_all 第一遍扫描得到的"带手动序号的标题"信息：
#   index     — 段落序号（1-based）
#   level     — 大纲级别 (1~9)
#   prefix    — 检测到的手动序号前缀文本
#   style     — 对应的自动编号样式名
#   values    — 各级序号数值（"3.1" → [3, 1]）
_NumberedHeading = collections.namedtuple(
    '_NumberedHeading', ['index', 'level', 'prefix', 'style', 'values'])

# 多级序号（"3.1"、"3.1.2"）内部各数字组之间的分隔符，仅点号类
_MULTI_LEVEL_SEPS = '.．'

# 多级序号的样式标识：各级标题必须链接到同一个多级列表模板，
# 由Word级联生成"%1.%2"这样的编号，故独立于单级样式单独标识。
MULTI_LEVEL_STYLE = '1.1'


def _to_half(ch):
    """将全角字符转换为半角，非全角字符原样返回"""
    return ch.translate(_FULL_TO_HALF_MAP)


def _classify(ch):
    """返回单个字符的类别标识，无法识别返回 None"""
    # 优先匹配中文特有字符（不经过全角→半角转换）
    if ch in _CHINESE_NUMS:
        return 'chinese_num'
    if ch in _HEAVENLY_STEMS:
        return 'heavenly_stem'
    if ch in _EARTHLY_BRANCHES:
        return 'earthly_branch'
    # 圈号类
    if ch in _CIRCLED_NUMS:
        return 'circled_num'
    if ch in _CIRCLED_UPPER_ALPHA:
        return 'circled_upper'
    if ch in _CIRCLED_LOWER_ALPHA:
        return 'circled_lower'
    # 罗马数字
    if ch in _ROMAN_UPPER:
        return 'roman_upper'
    if ch in _ROMAN_LOWER:
        return 'roman_lower'
    # 希腊字母
    if ch in _GREEK_UPPER:
        return 'greek_upper'
    if ch in _GREEK_LOWER:
        return 'greek_lower'
    # 经过全角→半角归一化后再判断
    nh = _to_half(ch)
    if nh in _ARABIC_HALF:
        return 'arabic'
    if nh in _UPPER_ALPHA_HALF:
        return 'upper_alpha'
    if nh in _LOWER_ALPHA_HALF:
        return 'lower_alpha'
    return None


def _is_fullwidth_digit(ch):
    """判断字符是否为全角阿拉伯数字（１２３…）"""
    return ch in _ARABIC_FULL_SET


def _detect_style_base(prefix, char_type):
    """在字符类别基础上细化序号体系，保留全角等显示形式。

    _classify() 会把全角数字归一化为 'arabic'（同类别便于比对），但
    "转自动序号"要求转换后形式一致，因此全角序号必须映射到全角编号
    样式。此处在类别之上补一层形式判定。

    Args:
        prefix: 检测到的序号前缀文本
        char_type: _classify() 返回的字符类别

    Returns:
        str: 细化后的体系名（如 'arabic_full'），无细化时原样返回 char_type
    """
    if char_type != 'arabic' or not prefix:
        return char_type

    # 取前缀中的第一个数字字符判断宽度：多级序号（"３．１"）各级
    # 通常同为全角，按首级判定即可
    for ch in prefix:
        if ch in _ARABIC_HALF:
            return 'arabic'
        if ch in _ARABIC_FULL_SET:
            return 'arabic_full'
    return char_type


def _is_same_class(ch, cls):
    """判断字符 ch 是否属于 class 类别（同宗类别如 roman_upper/roman_lower 视为同类别）"""
    cc = _classify(ch)
    if cc is None or cls is None:
        return False
    if cc == cls:
        return True
    # 合并同类
    same_groups = [
        {'roman_upper', 'roman_lower'},
        {'greek_upper', 'greek_lower'},
        {'upper_alpha', 'lower_alpha'},
        {'arabic', 'chinese_num'},
        {'circled_num', 'circled_upper', 'circled_lower'},
    ]
    for grp in same_groups:
        if cc in grp and cls in grp:
            return True
    return False


def _is_separator(ch):
    """判断字符是否为分隔符"""
    return ch in _ALL_SEPS


def _is_blank(ch):
    """判断字符是否为空白字符（空格、制表符等）"""
    return ch in ' \t\r\n'


def detect_number_prefix(text):
    """检测文本开头的序号前缀。

    对段落文本进行分析，识别开头的手动序号（含各种数字体系与分隔符组合），
    返回序号前缀文本及其对应的Word自动编号样式名称。

    Args:
        text: 段落文本（可含末尾 \\r）

    Returns:
        (prefix, style_name, char_type) 三元组：
        - prefix: str — 检测到的序号前缀文本
        - style_name: str — Word自动编号样式名
          （"1." / "一." / "①" / "资料1. " / "步骤1. " / "第1章" / "1.1"）
        - char_type: str — 序号字符类别标识
        若未检测到序号则返回 (None, None, None)
    """
    if not text:
        return None, None, None

    # 去除末尾段落标记（\r、\x07 等Word内部字符）
    clean = text.rstrip('\r\n\x07')
    if not clean:
        return None, None, None

    # 去除开头的空白，记录空白长度
    stripped = clean.lstrip()
    leading_ws_len = len(clean) - len(stripped)

    if not stripped:
        return None, None, None

    result = None

    # ---- 模式1：括号包围形式  (1) （一） (a) （Ａ） 等 ----
    result = _detect_paren_form(stripped)
    if result:
        inner_prefix, style, ctype = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype

    # ---- 模式2：资料N. 特殊前缀 ----
    result = _detect_ziliao_form(stripped)
    if result:
        inner_prefix, style, ctype = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype

    # ---- 模式3：步骤N. 特殊前缀 ----
    result = _detect_step_form(stripped)
    if result:
        inner_prefix, style, ctype = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype

    # ---- 模式4：第N章 章序号（第1章、第一章、第１章）----
    result = _detect_chapter_form(stripped)
    if result:
        inner_prefix, style, ctype = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype

    # ---- 模式5：圈号单字符序号（①②③…、ⓐⓑⓒ…） ----
    c0 = _classify(stripped[0])
    if c0 in ('circled_num', 'circled_upper', 'circled_lower'):
        prefix = clean[:leading_ws_len] + stripped[0]
        return prefix, '①', c0

    # ---- 模式6：多级序号（3.1、3.1.2）----
    result = _detect_multilevel_form(stripped)
    if result:
        inner_prefix, style, ctype = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype

    # ---- 模式7：数字/字母序列 + 可选分隔符 ----
    result = _detect_num_sep_form(stripped)
    if result:
        inner_prefix, style, ctype = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype

    return None, None, None


def _detect_paren_form(stripped):
    """检测括号包围形式的序号： (1) （一） (a) （Ａ） 等"""
    m = re.match(r'^[\(\（]\s*(\S+?)\s*[\)\）]', stripped)
    if not m:
        return None

    inner = m.group(1)
    if not inner:
        return None

    # 确定括号内字符的类型
    inner_cls = _classify(inner[0])
    if inner_cls is None:
        return None

    # 验证 inner 所有字符类型一致
    for ch in inner[1:]:
        if not _is_same_class(ch, inner_cls):
            return None

    full_prefix = m.group()
    # 成对括号：(1) （一）
    style = _classify_to_style(inner_cls, sep='_PAIRED_')
    return full_prefix, style, inner_cls


def _detect_ziliao_form(stripped):
    """检测"资料N. "形式：资料 + 阿拉伯数字 + . ＋ 可选空格"""
    m = re.match(r'^资料\s*(\d+)\s*[.．]\s*', stripped)
    if not m:
        return None

    full_prefix = m.group()
    return full_prefix, '资料1. ', 'arabic'


def _detect_step_form(stripped):
    """检测"步骤N. "形式：步骤 + 阿拉伯数字 + . ＋ 可选空格

    与"资料N."同为纯文字前缀序号，"步骤"二字整体属于序号的一部分
    （删除时连同"步骤"一起删除）。样式名统一返回 "步骤1. "，
    实际渲染的编号由 _link_list_template 按目标级别生成
    （如一级标题 → "步骤1. "，二级标题 → "步骤2. "）。
    """
    m = re.match(r'^步骤\s*(\d+)\s*[.．]\s*', stripped)
    if not m:
        return None

    full_prefix = m.group()
    return full_prefix, '步骤1. ', 'arabic'


def _detect_chapter_form(stripped):
    """检测"第N章"形式的章序号：第1章 / 第一章 / 第１章 ＋ 可选空格。

    与"资料N."类似，"第"与"章"整体属于序号的一部分，因此整个匹配文本
    都作为前缀返回（删除时连同"第""章"一起删除）。
    样式名统一返回 "第1章"，实际渲染的章号由 _link_list_template 按
    目标级别生成（如一级标题 → "第1章 "）。

    Args:
        stripped: 已去除开头空白的段落文本

    Returns:
        (prefix, '第1章', char_type) 三元组，未匹配返回 None
    """
    m = re.match(r'^第\s*(\d+|[０-９]+|[零〇一二三四五六七八九十百千]+)\s*章\s*', stripped)
    if not m:
        return None

    # 只接受阿拉伯数字（含全角）与中文数字两类章号
    ctype = _classify(m.group(1)[0])
    if ctype not in ('arabic', 'chinese_num'):
        return None

    matched = m.group()
    # "章"后没有空白时，若紧跟"节/条/款"等字，则可能是
    # "第三章节能设计"这类普通词语而非章序号，二者无法可靠区分；
    # 由于识别结果会连同"第""章"一起从标题中删除，此处宁可不识别，
    # 以免误删标题文字。带空白的"第1章 节能设计"不受影响。
    if (not matched[-1].isspace()
            and m.end() < len(stripped)
            and stripped[m.end()] in '节節条款章篇'):
        return None

    return matched, '第1章', ctype


def _detect_multilevel_form(stripped):
    """检测"3.1"、"3.1.2"形式的多级序号。

    多级序号由若干阿拉伯数字组（含全角）用点号连接而成。整串数字都是
    序号的一部分，必须整体识别：否则 _detect_num_sep_form 只会把开头的
    "3."当作单级序号删掉，把剩下的"1 项目背景"留在标题里。

    Args:
        stripped: 已去除开头空白的段落文本

    Returns:
        (prefix, MULTI_LEVEL_STYLE, 'arabic') 三元组，未匹配返回 None
    """
    m = re.match(
        r'^[0-9０-９]+(?:[%s][0-9０-９]+)+' % re.escape(_MULTI_LEVEL_SEPS),
        stripped)
    if not m:
        return None

    end = m.end()
    # 序号后可以跟分隔符（"3.1. 项目背景"）、空白（"3.1 项目背景"），
    # 也可以直接接标题文字（"3.1项目背景"）。
    if end < len(stripped) and _is_separator(stripped[end]):
        end += 1
    while end < len(stripped) and _is_blank(stripped[end]):
        end += 1

    return stripped[:end], MULTI_LEVEL_STYLE, 'arabic'


def _detect_num_sep_form(stripped):
    """检测"数字序列 + 可选分隔符"形式的序号。

    收集连续同类型字符作为数字部分，然后检查后续字符是否为有效分隔符。
    对于中文数字、天干、地支，允许省略分隔符。
    """
    first_cls = _classify(stripped[0])
    if first_cls is None:
        return None

    # 收集连续同类数字字符
    num_end = 1
    for i in range(1, len(stripped)):
        if _is_same_class(stripped[i], first_cls):
            num_end = i + 1
        else:
            break

    num_part = stripped[:num_end]

    # 检查分隔符
    sep = ''
    sep_offset = 0
    if num_end < len(stripped):
        ch = stripped[num_end]
        if _is_separator(ch):
            sep = ch
            sep_offset = 1
            # 跳过分隔符后的空白
            while num_end + sep_offset < len(stripped) and _is_blank(stripped[num_end + sep_offset]):
                sep_offset += 1

    # ---- 判断是否接受此序号 ----
    # 阿拉伯数字、字母、罗马数字、希腊字母等必须有分隔符
    needs_sep = first_cls not in ('chinese_num', 'heavenly_stem', 'earthly_branch')

    if needs_sep and not sep:
        return None

    full_prefix = stripped[:num_end + sep_offset]
    # 细化体系（全角数字 → 'arabic_full'），使转换后保持全角形式
    style_base = _detect_style_base(num_part, first_cls)
    style = _classify_to_style(style_base, sep if sep else '、')
    return full_prefix, style, first_cls
def _classify_to_style(char_type, sep):
    """根据序号字符类型和分隔符，返回Word自动编号样式名（内部标识）。

    该样式名是模块内部的"序号体系标识"，用于在检测与建列表之间传递
    **完整的序号形式信息**（字符体系 + 分隔符），使转换后的自动编号
    与转换前的手动序号形式一致：
      "一、"  → 中文小写数字 + 顿号
      "1. "   → 阿拉伯数字 + 点号
      "（一）" → 中文小写数字 + 括号包围
      "甲、"  → 天干 + 顿号
      "A."    → 大写字母 + 点号
      "Ⅰ."   → 罗马数字（大写） + 点号

    Args:
        char_type: 序号字符类别标识（_classify 的返回值）
        sep: 序号后的分隔符字符（无分隔符时传'、'等占位）

    Returns:
        str: 序号体系标识，形如 "arabic:sep" / "chinese_num:paren"。
             ":" 之后的字符即分隔符类别，用于还原转换前的手动序号形式。
    """
    # 圈号类：字符本身自带编号，不需要分隔符
    if char_type in ('circled_num', 'circled_upper', 'circled_lower'):
        return '①'

    # 成对括号形式：(1) （一） (a) （Ａ）
    if sep == _SEP_PAREN_PAIRED:
        return '%s:paren' % char_type

    # 仅右括号收尾形式：1) a) 三)——转换后只保留右括号
    if sep in _SEP_PAREN_R:
        return '%s:rparen' % char_type

    # 点号 / 逗号 / 顿号后的空白：统一归一化为"分隔符+空格"以便还原
    return '%s:sep' % char_type


# ============================================================
# 序号体系标识 → Word 编号样式（WdListNumberStyle）映射
#
# "转自动序号"的核心要求：转换后的自动编号必须与转换前的手动序号
# 形式完全一致。因此每个序号字符体系都映射到 Word 对应的编号样式，
# 而不是一律套用阿拉伯数字。
#
# 值为 (NumberStyle, 是否可与阿拉伯数字级联显示)。
# "是否可级联"用于多级序号：能级联的体系可与 %1.%2 组合显示，
# 不能级联的体系（中文数字、圈号等把上级序号也当作自身字体的样式）
# 在多级场景下需要降级处理，详见 _ensure_multilevel_template。
# ============================================================
_STYLE_MAP = {
    # ---- 阿拉伯数字 ----
    'arabic':          (wc.wdListNumberStyleArabic, True),
    'arabic_full':     (wc.wdListNumberStyleArabicFullWidth, True),
    # ---- 中文数字 ----
    'chinese_num':     (wc.wdListNumberStyleSimpChinNum1, False),
    'chinese_num_upper': (wc.wdListNumberStyleSimpChinNum2, False),
    # ---- 天干 / 地支 ----
    'heavenly_stem':   (wc.wdListNumberStyleZodiac1, False),
    'earthly_branch':  (wc.wdListNumberStyleZodiac2, False),
    # ---- 英文大小写字母 ----
    'upper_alpha':     (wc.wdListNumberStyleUppercaseLetter, True),
    'lower_alpha':     (wc.wdListNumberStyleLowercaseLetter, True),
    # ---- 罗马数字 ----
    'roman_upper':     (wc.wdListNumberStyleUppercaseRoman, True),
    'roman_lower':     (wc.wdListNumberStyleLowercaseRoman, True),
    # ---- 希腊字母 ----
    'greek_upper':     (wc.wdListNumberStyleUppercaseGreek, True),
    'greek_lower':     (wc.wdListNumberStyleLowercaseGreek, True),
    # ---- 圈号（数字 / 大小写字母）----
    'circled_num':     (wc.wdListNumberStyleNumberInCircle, False),
    'circled_upper':   (wc.wdListNumberStyleNumberInCircle, False),
    'circled_lower':   (wc.wdListNumberStyleNumberInCircle, False),
}

# 全角字符体系 → 半角体系（供 NumberStyle 复用半角样式时取用）
_FULL_TO_BASE_CLASS = {
    'arabic_full': 'arabic',
}


def _resolve_style(char_type, sep=None):
    """把序号体系标识解析为 Word 编号样式。

    Args:
        char_type: 序号体系标识，形如 'chinese_num:sep'、'arabic:paren'；
                   也接受 '①'（圈号哨兵）或不带后缀的原始类别名（如 'arabic'）
        sep: 兼容早期调用而保留，当前未使用

    Returns:
        tuple: (base_class, number_style, can_cascade)
        - base_class: 归一化后的字符体系名（如 'arabic'）
        - number_style: WdListNumberStyle 数值
        - can_cascade: 该体系能否与阿拉伯数字级联显示
        无法识别时返回 (None, None, False)
    """
    if not char_type:
        return None, None, False

    # 圈号哨兵：字符本身即序号，无分隔符
    if char_type == '①':
        return ('circled_num', wc.wdListNumberStyleNumberInCircle, False)

    # 去掉":sep" / ":paren"后缀，得到字符体系名
    base = char_type.split(':', 1)[0]

    entry = _STYLE_MAP.get(base)
    if entry is None:
        # 未知体系：退回阿拉伯数字，保证至少能建立自动编号
        return 'arabic', wc.wdListNumberStyleArabic, True

    number_style, can_cascade = entry
    # 全角体系复用半角样式名（Word 无独立的全角字母/罗马数字编号样式）
    return _FULL_TO_BASE_CLASS.get(base, base), number_style, can_cascade


# Word 编号样式中"本身已含括号"的样式不存在，括号由 NumberFormat 的前缀
# 字符实现；这些字符就是手动序号里的括号，转换后照原样保留。
_PAREN_LEFT = '('
_PAREN_RIGHT = ')'

# 圈号类体系：字符集里只有 ①~⑳ / Ⓐ~ⓩ，没有 0-9 的数字形，
# 因此无法在多级序号里作为 %N 参与级联显示，需要降级为数字。
_UNRESOLVABLE_CASCADE = {'circled_num', 'circled_upper', 'circled_lower'}


def _style_cannot_cascade(number_style):
    """判断某序号体系能否在多级序号中作为 %N 级联显示。

    多级序号（如"3.1"）要求上级序号以数字形出现在本级编号里。圈号类
    字符集没有 0-9 数字形，无法级联；其余体系（阿拉伯数字、全角数字、
    中文数字、天干地支、字母、罗马/希腊字母）都能在 %N 位置渲染。

    Args:
        number_style: 序号体系标识（形如 "arabic:sep"）

    Returns:
        bool: True 表示无法级联、需要降级为数字
    """
    base = (number_style or '').split(':', 1)[0]
    return base in _UNRESOLVABLE_CASCADE


def _number_format_for_style(number_style, level):
    """把序号体系标识转换为指定级别的Word列表级别格式。

    转换后的编号形式与转换前的手动序号形式保持一致：
      "一、"    → 中文数字 + 顿号     → ("%N、", SimpChinNum1)
      "（一）"  → 中文数字 + 括号     → ("(%N)", SimpChinNum1)
      "1. "     → 阿拉伯数字 + 点号   → ("%N. ", Arabic)
      "１．"    → 全角阿拉伯数字 + 点号 → ("%N. ", ArabicFullWidth)
      "甲、"    → 天干 + 顿号         → ("%N、", Zodiac1)
      "A."      → 大写字母 + 点号     → ("%N. ", UppercaseLetter)
      "Ⅰ."     → 罗马数字（大写）     → ("%N. ", UppercaseRoman)
      "①"      → 圈号数字             → ("%N", NumberInCircle)

    Args:
        number_style: 序号体系标识
            （形如 "arabic:sep" / "chinese_num:paren"，也兼容
              "1." / "一." / "①" / MULTI_LEVEL_STYLE / "资料1. " /
              "步骤1. " / "第1章" 等早期样式名）
        level: 目标标题级别 (1~9)

    Returns:
        (NumberFormat, NumberStyle) 二元组；无法识别时返回 (None, None)
    """
    # ---- 特殊文字前缀序号：保持原有文字前缀形式 ----
    if number_style == "资料1. ":
        return f"资料%{level}. ", wc.wdListNumberStyleArabic
    if number_style == "步骤1. ":
        return f"步骤%{level}. ", wc.wdListNumberStyleArabic
    if number_style == "第1章":
        return f"第%{level}章 ", wc.wdListNumberStyleArabic

    # ---- 多级序号：级联格式（一级 "%1. "、二级 "%1.%2. "……）----
    if number_style == MULTI_LEVEL_STYLE:
        return ".".join(f"%{i}" for i in range(1, level + 1)) + ". ", wc.wdListNumberStyleArabic

    # ---- 圈号：编号字符本身即序号，不加分隔符 ----
    if number_style == '①' or number_style.split(':', 1)[0] in (
            'circled_num', 'circled_upper', 'circled_lower'):
        _, ns, _ = _resolve_style(number_style)
        return f"%{level}", ns

    # ---- 兼容早期样式名（"1." / "一."）----
    if number_style == "一.":
        return f"%{level}、", wc.wdListNumberStyleSimpChinNum1
    if number_style == "1.":
        return f"%{level}. ", wc.wdListNumberStyleArabic

    # ---- 常规体系：按原序号形式还原分隔符 ----
    base, ns, _ = _resolve_style(number_style)
    if ns is None:
        return None, None

    marker = f"%{level}"
    if number_style.endswith(':paren'):
        # 成对括号形式：(1) （一）
        return f"{_PAREN_LEFT}{marker}{_PAREN_RIGHT}", ns
    if number_style.endswith(':rparen'):
        # 仅右括号收尾形式：1) a)——保持原样，不加左括号
        return f"{marker}{_PAREN_RIGHT}", ns

    # 点号分隔（形如 "1. "、"一. "、"甲. "）
    return f"{marker}. ", ns


# ============================================================
# 序号前缀 → 数值 转换
# ============================================================

# 有序字符映射（用于将序号字符转换为1-based位置值）
_CIRCLED_NUM_ORDER = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿'
_CIRCLED_UPPER_ORDER = 'ⒶⒷⒸⒹⒺⒻⒼⒽⒾⒿⓀⓁⓂⓃⓄⓅⓆⓇⓈⓉⓊⓋⓌⓍⓎⓏ'
_CIRCLED_LOWER_ORDER = 'ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝⓞⓟⓠⓡⓢⓣⓤⓥⓦⓧⓨⓩ'
_HEAVENLY_STEMS_ORDER = '甲乙丙丁戊己庚辛壬癸'
_EARTHLY_BRANCHES_ORDER = '子丑寅卯辰巳午未申酉戌亥'
_ROMAN_MAP = {
    'Ⅰ': 1, 'Ⅱ': 2, 'Ⅲ': 3, 'Ⅳ': 4, 'Ⅴ': 5,
    'Ⅵ': 6, 'Ⅶ': 7, 'Ⅷ': 8, 'Ⅸ': 9, 'Ⅹ': 10,
    'Ⅺ': 11, 'Ⅻ': 12, 'Ⅼ': 50, 'Ⅽ': 100, 'Ⅾ': 500, 'Ⅿ': 1000,
    'ⅰ': 1, 'ⅱ': 2, 'ⅲ': 3, 'ⅳ': 4, 'ⅴ': 5,
    'ⅵ': 6, 'ⅶ': 7, 'ⅷ': 8, 'ⅸ': 9, 'ⅹ': 10,
    'ⅺ': 11, 'ⅻ': 12, 'ⅼ': 50, 'ⅽ': 100, 'ⅾ': 500, 'ⅿ': 1000,
}

# 中文数字 → 数值映射
_CHINESE_DIGIT_MAP = {
    '零': 0, '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4,
    '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
    '两': 2,
}

# 希腊字母顺序（大写 + 小写各24个）
_GREEK_UPPER_ORDER = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ'
_GREEK_LOWER_ORDER = 'αβγδεζηθικλμνξοπρστυφχψω'


def _extract_prefix_core(prefix):
    """从序号前缀文本中提取核心序号部分。

    去除前导空白、括号包围、结尾分隔符，以及"资料/步骤/第…章"等
    属于序号但需要单独剥离的文字前缀。

    Args:
        prefix: detect_number_prefix() 返回的序号前缀文本

    Returns:
        str: 核心序号文本（如"（一）"→"一"、"资料3. "→"3"、"3.1 "→"3.1"）
    """
    # 去除前导空白
    core = prefix.lstrip()
    # 去除括号包围
    core = re.sub(r'^[\(\（]\s*', '', core)
    core = re.sub(r'\s*[\)\）]$', '', core)
    # 去除结尾分隔符及后续空白
    core = re.sub(r'[.．,，、)）]\s*$', '', core)
    # 去除"资料"前缀（资料N. 形式）
    core = re.sub(r'^资料\s*', '', core)
    # 去除"步骤"前缀（步骤N. 形式）
    core = re.sub(r'^步骤\s*', '', core)
    # 去除"第N章"的"第"与"章"（章序号形式）
    core = re.sub(r'^第\s*', '', core)
    core = re.sub(r'\s*章\s*$', '', core)
    return core


def parse_prefix_values(prefix, char_type):
    """解析序号前缀，返回各级序号的数值列表（1-based）。

    单级序号返回单元素列表（"3. "→[3]、"（一）"→[1]、"①"→[1]）；
    多级序号按点号拆分为各级数值（"3.1"→[3, 1]、"3.1.2"→[3, 1, 2]）。
    拆出的上级数值用于多级编号时设置各级计数器的起始值。

    Args:
        prefix: detect_number_prefix() 返回的序号前缀文本
                 （如"（一）"、"3. "、"3.1"、"①"、"IV."等）
        char_type: 序号字符类别标识（如 'chinese_num', 'arabic', 'circled_num' 等）

    Returns:
        list: 各级序号数值（从1开始），无法解析返回空列表
    """
    if not prefix or char_type is None:
        return []

    core = _extract_prefix_core(prefix)
    if not core:
        return []

    try:
        # ---- 阿拉伯数字（含全角），可能是"3.1"这类多级序号 ----
        if char_type == 'arabic':
            half = core.translate(_FULL_TO_HALF_MAP)
            values = []
            for part in re.split(r'[.．。]', half):
                part = part.strip()
                if part:
                    values.append(int(part))
            return values

        value = _parse_single_prefix_value(core, char_type)
        return [value] if value > 0 else []

    except (ValueError, IndexError):
        return []


def _parse_prefix_value(prefix, char_type):
    """将序号前缀文本转换为对应的序号数值（1-based）。

    多级序号取最后一级的数值（"3.1"→1、"3.1.2"→2），因为标题自身
    只占用最后一级的序号。

    Args:
        prefix: detect_number_prefix() 返回的序号前缀文本
        char_type: 序号字符类别标识

    Returns:
        int: 序号数值（从1开始），无法转换返回 0
    """
    values = parse_prefix_values(prefix, char_type)
    return values[-1] if values else 0


def _parse_single_prefix_value(core, char_type):
    """将单级序号的核心文本转换为数值（1-based），无法转换返回 0。

    Args:
        core: _extract_prefix_core() 得到的核心序号文本
        char_type: 序号字符类别标识

    Returns:
        int: 序号数值（从1开始），无法转换返回 0
    """
    if not core:
        return 0

    try:
        # ---- 圈号数字 ----
        if char_type == 'circled_num':
            idx = _CIRCLED_NUM_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0

        # ---- 圈号字母（大写/小写）----
        if char_type == 'circled_upper':
            idx = _CIRCLED_UPPER_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0
        if char_type == 'circled_lower':
            idx = _CIRCLED_LOWER_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0

        # ---- 中文数字 ----
        if char_type == 'chinese_num':
            return _parse_chinese_num(core)

        # ---- 天干 ----
        if char_type == 'heavenly_stem':
            idx = _HEAVENLY_STEMS_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0

        # ---- 地支 ----
        if char_type == 'earthly_branch':
            idx = _EARTHLY_BRANCHES_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0

        # ---- 罗马数字（大写/小写）----
        if char_type in ('roman_upper', 'roman_lower'):
            return _parse_roman_num(core)

        # ---- 大写字母（含全角）----
        if char_type == 'upper_alpha':
            half = core[0].translate(_FULL_TO_HALF_MAP)
            if 'A' <= half <= 'Z':
                return ord(half) - ord('A') + 1
            return 0

        # ---- 小写字母（含全角）----
        if char_type == 'lower_alpha':
            half = core[0].translate(_FULL_TO_HALF_MAP)
            if 'a' <= half <= 'z':
                return ord(half) - ord('a') + 1
            return 0

        # ---- 希腊字母（大写/小写）----
        if char_type == 'greek_upper':
            idx = _GREEK_UPPER_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0
        if char_type == 'greek_lower':
            idx = _GREEK_LOWER_ORDER.find(core[0])
            if idx >= 0:
                return idx + 1
            return 0

    except (ValueError, IndexError):
        pass

    return 0


def _parse_chinese_num(s):
    """将中文数字字符串解析为整数值。

    支持格式：
      - 一 ~ 九          → 1~9
      - 十 ~ 十九        → 10~19
      - 二十 ~ 九十九    → 20~99
      - 一百             → 100
      - 零 / 〇          → 0

    Args:
        s: 中文数字字符串（如"十二"、"三十"等）

    Returns:
        int: 对应的整数值
    """
    if not s:
        return 0

    result = 0
    current = 0  # 当前累积的数字（用于处理"三百"等形式）

    for ch in s:
        if ch in _CHINESE_DIGIT_MAP:
            current = _CHINESE_DIGIT_MAP[ch]
        elif ch == '十':
            if current == 0:
                current = 1  # "十" → 10
            result += current * 10
            current = 0
        elif ch == '百':
            if current == 0:
                current = 1  # "百" → 100
            result += current * 100
            current = 0
        elif ch == '千':
            if current == 0:
                current = 1
            result += current * 1000
            current = 0
        elif ch == '万':
            if current == 0:
                current = 1
            result = (result + current) * 10000
            current = 0
        elif ch == '亿':
            if current == 0:
                current = 1
            result = (result + current) * 100000000
            current = 0

    result += current
    return result


def _parse_roman_num(s):
    """将罗马数字字符串解析为整数值。

    支持 Unicode 罗马数字字符（Ⅰ~Ⅿ, ⅰ~ⅿ）。
    采用累加法：从左到右，若当前值小于下一个值则减去，否则加上。

    Args:
        s: 罗马数字字符串（如"Ⅲ"、"Ⅳ"、"Ⅷ"等）

    Returns:
        int: 对应的整数值
    """
    values = []
    for ch in s:
        v = _ROMAN_MAP.get(ch, 0)
        if v == 0:
            # 遇到无法识别的字符则终止解析
            break
        values.append(v)

    if not values:
        return 0

    total = 0
    for i, v in enumerate(values):
        if i + 1 < len(values) and v < values[i + 1]:
            total -= v
        else:
            total += v

    return total


# ============================================================
# 主类：与Word文档交互
# ============================================================

class AutoNumbering:
    """标题手动序号 → Word自动编号 转换器

    用法:
        converter = AutoNumbering(work_doc)
        count = converter.convert_all()
        print(f"已转换 {count} 个标题")
    """

    def __init__(self, doc):
        """初始化转换器。

        Args:
            doc: Word Document COM对象
        """
        self.work_doc = doc
        # 记录已为各级别设置的编号样式，避免重复创建列表模板
        self._level_styles_set = {}
        # 多级序号（"3.1"）共用的列表模板，首次需要时创建
        self._multilevel_template = None

    def convert_all(self):
        """遍历文档所有标题段落（大纲级别1~9），将手动序号转为自动编号。

        对每个标题段落：
          1. 调用 detect_number_prefix() 检测手动序号
          2. 使用Range.Delete删除序号前缀文本
          3. 为该级别标题样式链接对应的自动编号列表模板

        多级序号（"3.1"、"3.1.2"）必须让各级标题共用同一个多级列表模板，
        由Word级联生成各级编号，因此扫描与修改分成两遍：先在只读的第一遍
        里探明文档是否使用多级序号，再决定各级标题的列表模板。

        Returns:
            int: 成功转换的段落数量
        """
        if self.work_doc is None:
            return 0

        # ---- 第一遍：扫描标题段落，收集手动序号信息 ----
        items = []
        para_count = self.work_doc.Paragraphs.Count

        for i in range(1, para_count + 1):
            try:
                para = self.work_doc.Paragraphs(i)

                # 只处理大纲级别1~9的段落，无论其样式名是否为"标题X"
                level = para.OutlineLevel
                if level < 1 or level > 9:
                    continue

                text = para.Range.Text
                if not text or not text.strip():
                    continue

                # 检测手动序号前缀
                prefix, style_name, char_type = detect_number_prefix(text)
                if prefix is None:
                    continue

                items.append(_NumberedHeading(
                    index=i, level=level, prefix=prefix, style=style_name,
                    values=parse_prefix_values(prefix, char_type)))

            except Exception:
                # 忽略单个段落的异常，继续处理后续段落
                continue

        if not items:
            return 0

        # ---- 多级序号：先把各级标题接入同一个多级列表模板 ----
        # 必须在删除任何序号文本之前完成：多级编号要求所有相关级别共用
        # 一个列表模板，先改后建会让已处理的级别停留在单级模板上。
        if any(item.style == MULTI_LEVEL_STYLE for item in items):
            self._setup_multilevel_numbering(items)
        multilevel = self._multilevel_template is not None

        # ---- 第二遍：删除手动序号文本并建立自动编号 ----
        converted = 0

        for item in items:
            try:
                para = self.work_doc.Paragraphs(item.index)
                level, style_name = item.level, item.style

                if multilevel:
                    # 多级编号：模板与样式链接已在第一遍之后统一建立，
                    # 这里逐段套用同一列表模板即可（详见 _apply_multilevel_list）。
                    # 注意多级序号不能像单级序号那样逐段强制起始值：
                    # ApplyListTemplate(ContinuePreviousList=False) 会连带
                    # 重置上级计数器，把"2.5"变成"1.5"。
                    self._apply_multilevel_list(para, level)
                else:
                    # ---- 对比并调整自动序号值 ----
                    # 提取前缀对应的序号数值（如 "（三）" → 3）
                    prefix_value = item.values[-1] if item.values else 0
                    if prefix_value > 0:
                        # 先为该级别链接列表模板以激活自动编号
                        if level not in self._level_styles_set:
                            self._link_list_template(level, style_name)
                            self._level_styles_set[level] = style_name
                        # 获取当前自动序号值，与期望值对比
                        try:
                            current_value = para.Range.ListFormat.ListValue
                            if current_value != prefix_value:
                                self._restart_list_numbering(para, level, prefix_value)
                        except Exception:
                            pass

                # ---- 删除序号前缀文本 ----
                # 通过复制段落Range并调整End位置来精确定位前缀区域，
                # 然后删除该区域。相比直接设置Range.Text更安全，
                # 不会意外影响段落标记或后续段落。
                #
                # 删除长度以**当前**文本检测到的前缀为准：套用列表模板时
                # Word可能已把手动序号换成了列表编号（此时段落里已经没有
                # 序号文本），若沿用第一遍扫描的长度会误删标题文字。
                current_text = para.Range.Text
                current_prefix, _, _ = detect_number_prefix(current_text)
                if current_prefix:
                    rng = para.Range.Duplicate
                    prefix_len_in_range = len(current_prefix)
                    # 防御：确保不超过段落文本长度（保留末尾段落标记）
                    max_delete = max(1, len(current_text) - 1)  # 至少保留 \r
                    if prefix_len_in_range > max_delete:
                        prefix_len_in_range = max_delete
                    rng.End = rng.Start + prefix_len_in_range
                    rng.Delete()

                # ---- 链接自动编号列表模板 ----
                # 确保每个级别只链接一次，避免重复创建列表模板
                # （多级序号模式下样式已统一链接到多级列表模板，此处跳过）
                if not multilevel and level not in self._level_styles_set:
                    self._link_list_template(level, style_name)
                    self._level_styles_set[level] = style_name

                converted += 1

            except Exception:
                # 忽略单个段落的异常，继续处理后续段落
                continue

        return converted

    def _get_heading_style(self, level):
        """获取"标题 N"样式对象，文档中不存在该样式时返回 None。

        Args:
            level: 标题级别 (1~9)

        Returns:
            Style COM对象或 None
        """
        for sn in (f"标题 {level}", f"标题{level}"):
            try:
                return self.work_doc.Styles(sn)
            except Exception:
                continue
        return None

    def _link_style_to_template(self, level, list_template):
        """把"标题 N"样式链接到列表模板的第 N 级。

        使用 LinkToListTemplate 显式覆盖样式已有的列表模板链接。

        Args:
            level: 标题级别 (1~9)
            list_template: ListTemplate COM对象
        """
        style = self._get_heading_style(level)
        if style is None:
            return
        try:
            style.LinkToListTemplate(
                ListTemplate=list_template, ListLevelNumber=level)
        except Exception:
            try:
                list_template.ListLevels(level).LinkedStyle = style
            except Exception:
                pass

    def _link_list_template(self, level, number_style):
        """为指定级别的"标题 N"样式链接自动编号列表模板（单级序号）。

        先清空全部9级，再设置目标级别和第1级（安全兜底），
        使用 LinkToListTemplate 显式覆盖样式已有的列表模板链接。

        Args:
            level: 标题级别 (1~9)
            number_style: 编号样式名
                ("1." / "一." / "①" / "资料1. " / "步骤1. " / "第1章")
        """
        try:
            style = self._get_heading_style(level)
            if style is None:
                return

            list_template = self.work_doc.ListTemplates.Add(True)

            # --- 先清空全部9级 ---
            for lvl in range(1, 10):
                try:
                    level_obj = list_template.ListLevels(lvl)
                    level_obj.NumberFormat = ""
                    level_obj.NumberStyle = 255  # wdListNumberStyleNone
                except:
                    pass

            # --- 确定 NumberFormat 与 NumberStyle ---
            # 统一使用 \". \"（点号+空格）作为序号分隔符，
            # 将空格纳入 NumberFormat 字符串，配合 TrailingCharacter=wdTrailingNone
            # 使分隔符后的间距为0字符。
            fmt, ns = _number_format_for_style(number_style, level)
            if fmt is None:
                return

            # --- 同时设置目标级别和第1级作为安全兜底 ---
            for lvl in (1, level):
                target_level = list_template.ListLevels(lvl)
                target_level.NumberFormat = fmt
                target_level.NumberStyle = ns
                try:
                    list_gallery = wc.get_list_gallery(
                        self.work_doc.Application, wc.wdOutlineNumberGallery)
                    ref_template = list_gallery.ListTemplates(1)
                    ref_level = ref_template.ListLevels(1)
                    for attr in ('NumberPosition', 'Alignment',
                                 'ResetOnHigher', 'StartAt'):
                        try:
                            setattr(target_level, attr, getattr(ref_level, attr))
                        except Exception:
                            pass
                except Exception:
                    pass
                # 强制设置 TrailingCharacter 为无尾随字符，
                # 使序号后分隔符（\". \"）与段落文字之间的间距为0字符。
                try:
                    target_level.TrailingCharacter = wc.wdTrailingNone
                except Exception:
                    pass
                # 设置 TabPosition 为 0，消除序号与文字间的制表符间距。
                try:
                    target_level.TabPosition = 0
                except Exception:
                    pass

            # 使用 LinkToListTemplate 显式覆盖样式已有的列表模板链接
            self._link_style_to_template(level, list_template)

        except Exception:
            import traceback
            traceback.print_exc()

    def _setup_multilevel_numbering(self, items):
        """建立多级序号所需的共用列表模板，并设置各级起始序号。

        仅当文档中存在"3.1"这类多级手动序号时调用。Word 只有在各级标题
        样式链接到**同一个**多级列表模板时才会级联显示各级计数器
        （"3.1"、"3.1.2"…）；各级各自一个单级模板时，上级计数器没有
        来源，只能显示固定的起始值。

        Args:
            items: convert_all 第一遍扫描得到的标题信息（_NumberedHeading 列表）
        """
        # 每个级别采用的编号格式：取该级别第一个被识别标题的序号样式。
        # 多级序号（"3.1"）所在级别使用级联格式（%1.%2…），其余级别沿用
        # 各自的单级格式（"%1. "、"第%1章 "、"一. " 等），但它们共用同一个
        # 列表模板，序号计数器才能级联。
        level_styles = {}
        for item in items:
            level_styles.setdefault(item.level, item.style)

        list_template = self._ensure_multilevel_template(level_styles)
        if list_template is None:
            return

        # 各级起始序号：让自动序号从原文的起始值开始编号
        for level, value in self._compute_multilevel_start_values(items).items():
            try:
                list_template.ListLevels(level).StartAt = value
            except Exception:
                continue

        # 把各级"标题 N"样式链接到同一模板的对应级别，
        # 使没有手动序号、未被逐段处理的标题段落也沿用同一列表
        for level in sorted(level_styles):
            self._link_style_to_template(level, list_template)

    def _ensure_multilevel_template(self, level_styles):
        """创建（并缓存）多级序号共用的列表模板。

        模板的1~9级默认为级联的十进制格式（第1级 "%1. "、第2级
        "%1.%2. "、第3级 "%1.%2.%3. "…），再按 level_styles 覆盖已识别
        级别的格式，例如一级标题用章序号时保持"第%1章 "。

        Args:
            level_styles: {级别: 编号样式名}

        Returns:
            ListTemplate COM对象；创建失败返回 None
        """
        if self._multilevel_template is not None:
            return self._multilevel_template

        try:
            list_template = self.work_doc.ListTemplates.Add(True)
        except Exception:
            return None

        # 参与级联显示的级别：多级序号"3.1"会把上级计数器（%1）一起显示，
        # 因此 1~N 级都在同一串序号里出现。%1 的显示形式由上级级别自身的
        # 编号样式决定，上级若用中文数字，原文的"3.1"就会被渲染成"三.1"，
        # 所以这些级别必须统一为阿拉伯数字。
        # 多级序号本身（"3.1"）隐含要求上级计数器以阿拉伯数字显示，
        # 但若上级标题自己用的是中文/字母等形式（如"一.1"），则应以
        # 上级**自身检测到的形式**为准 —— Word 在渲染 %1 时会采用被
        # 引用级别自己的 NumberStyle，无需在此强制改写。
        #
        # cascade_levels 只记录"作为上级被多级序号引用"的级别，
        # 用于判断其是否为未识别级别补一个可计数的默认格式。
        cascade_levels = set()
        for lvl, style_name in level_styles.items():
            if style_name == MULTI_LEVEL_STYLE:
                cascade_levels.update(range(1, lvl + 1))

        # 内置多级列表模板的版式（缩进、对齐）作为参照，只取一次
        try:
            list_gallery = wc.get_list_gallery(
                self.work_doc.Application, wc.wdOutlineNumberGallery)
            ref_level = list_gallery.ListTemplates(1).ListLevels(1)
        except Exception:
            ref_level = None

        for lvl in range(1, 10):
            # 未识别到标题的级别也按级联格式定义：它们可能只作为上级
            # 计数器参与显示（如文档只有"3.1"这类二级标题时，%1 取
            # 一级计数器的起始值3），必须保留可计数的编号样式。
            style_name = level_styles.get(lvl, MULTI_LEVEL_STYLE)
            fmt, ns = _number_format_for_style(style_name, lvl)
            if fmt is None:
                continue
            if style_name == MULTI_LEVEL_STYLE:
                # 多级序号所在级别：用级联格式（%1.%2…）并保持阿拉伯数字，
                # 与手动序号"3.1"的显示形式一致。
                # 括号包围形式的多级序号（如"（3.1）"）保留括号。
                pass
            elif lvl in cascade_levels and _style_cannot_cascade(style_name):
                # 该级别既被上级计数器引用、自身又用了 Word 无法与
                # 数字级联显示的体系（圈号等）——这类体系的字符集里
                # 没有 0-9 的数字形，%N 无法解析。此时保留其编号体系
                # 但需要一个数字兜底：改用同级阿拉伯数字格式，
                # 避免整个多级列表建立失败。
                fmt, ns = f"%{lvl}. ", wc.wdListNumberStyleArabic
            try:
                level_obj = list_template.ListLevels(lvl)
                level_obj.NumberFormat = fmt
                level_obj.NumberStyle = ns
                if lvl > 1:
                    # 出现上级标题时重新开始本级编号
                    try:
                        level_obj.ResetOnHigher = lvl - 1
                    except Exception:
                        pass
                # 与单级编号保持一致：沿用内置多级列表模板的版式，
                # 并让分隔符后不留尾随字符、不产生制表符间距
                if ref_level is not None:
                    for attr in ('NumberPosition', 'Alignment'):
                        try:
                            setattr(level_obj, attr, getattr(ref_level, attr))
                        except Exception:
                            pass
                try:
                    level_obj.TrailingCharacter = wc.wdTrailingNone
                except Exception:
                    pass
                try:
                    level_obj.TabPosition = 0
                except Exception:
                    pass
            except Exception:
                continue

        self._multilevel_template = list_template
        return list_template

    @staticmethod
    def _compute_multilevel_start_values(items):
        """计算多级编号时各级标题的起始序号值。

        取该级别**第一个**标题的序号值作为起始值（例如全文从"3."开始
        编号时，一级标题起始值为3）；若该级别自己没有标题，而多级序号
        中蕴含了它的数值（如全文只有"3.1"这类二级标题，一级序号为3），
        则取蕴含值，使 %1 显示正确的上级序号。

        Args:
            items: convert_all 第一遍扫描得到的标题信息（_NumberedHeading 列表）

        Returns:
            dict: {级别: 起始值}
        """
        own_first = {}   # 级别 → 该级别首个标题的序号值
        preceded = {}    # 级别 → 该级别首个标题之前是否出现过更高级别标题
        implied = {}     # 级别 → 由多级序号的上级数字推断出的序号值

        for item in items:
            level, values = item.level, item.values
            if values and values[-1] > 0 and level not in own_first:
                own_first[level] = values[-1]
                preceded[level] = any(lv < level for lv in own_first)
            # "3.1"中的"3"即一级标题的序号
            for depth, value in enumerate(values[:-1], start=1):
                if value > 0 and depth not in implied:
                    implied[depth] = value

        start_values = {}
        for level, value in own_first.items():
            # 前面出现过更高级别标题时，本级别序号由上级标题触发重新
            # 计数，起始值应当是"重新开始"的值，不能取首个标题的序号值，
            # 否则每次重新计数都会从该值开始。
            start_values[level] = 1 if (level > 1 and preceded.get(level)) else value
        for level, value in implied.items():
            start_values.setdefault(level, value)
        return start_values

    def _apply_multilevel_list(self, para, level):
        """把共用的多级列表模板套用到指定段落。

        ContinuePreviousList=True 表示续接同一条列表，Word 会按文档顺序
        级联各级计数器，并在出现上级标题时自动重置下级序号。

        Args:
            para: 段落 COM 对象
            level: 标题级别 (1~9)，作为列表级别显式指定，
                   使样式未链接到模板（自建大纲样式）的标题也能落到正确级别
        """
        if self._multilevel_template is None:
            return
        try:
            # ApplyTo=0 → wdListApplyToWholeList
            para.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=self._multilevel_template,
                ContinuePreviousList=True,
                ApplyTo=0,
                ApplyLevel=level)
        except Exception:
            # 老版本 Word 没有 ApplyListTemplateWithLevel，退回普通应用
            try:
                para.Range.ListFormat.ApplyListTemplate(
                    self._multilevel_template, True, 0)
            except Exception:
                pass

    def _restart_list_numbering(self, para, level, start_value):
        """重新开始段落自动编号，并将起始值设为指定数值。

        通过临时修改列表模板的 StartAt 属性，然后对该段落调用
        ApplyListTemplate（ContinuePreviousList=False），
        使该段落的自动编号从 start_value 重新开始。

        Args:
            para: 段落 COM 对象
            level: 标题级别 (1~9)
            start_value: 目标起始编号值
        """
        try:
            lf = para.Range.ListFormat
            lt = lf.ListTemplate
            if lt is None:
                return

            old_start = None
            try:
                old_start = lt.ListLevels(level).StartAt
            except Exception:
                pass

            try:
                lt.ListLevels(level).StartAt = start_value
                # ContinuePreviousList=False → 重新开始编号
                # ApplyTo=2 → wdListApplyToThisPointForward
                lf.ApplyListTemplate(lt, False, 2)
            finally:
                if old_start is not None:
                    try:
                        lt.ListLevels(level).StartAt = old_start
                    except Exception:
                        pass

        except Exception:
            pass

# ============================================================
# 便捷函数：供外部直接调用
# ============================================================

def force_heading_auto_numbering(work_doc):
    """将文档中所有标题的手动序号转换为Word自动编号（便捷函数）。

    Args:
        work_doc: Word Document COM对象

    Returns:
        int: 成功转换的段落数量
    """
    # 创建一个AutoNumbering类的实例，传入Word文档COM对象
    converter = AutoNumbering(work_doc)
    # 调用converter实例的convert_all方法，转换所有标题并返回成功转换的段落数量
    return converter.convert_all()
