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

支持的分隔符（含全角形式）：
  - 点号：. ．
  - 逗号：, ，
  - 顿号：、
  - 右括号：) ）
  - 括号包围：（一） (1) 等
"""

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

# 所有分隔符合集
_ALL_SEPS = _SEP_DOT + _SEP_COMMA + _SEP_DUNHAO + _SEP_PAREN_R


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
        - style_name: str — Word自动编号样式名（"1." / "一." / "①"）
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

    # ---- 模式3：圈号单字符序号（①②③…、ⓐⓑⓒ…） ----
    c0 = _classify(stripped[0])
    if c0 in ('circled_num', 'circled_upper', 'circled_lower'):
        prefix = clean[:leading_ws_len] + stripped[0]
        return prefix, '①', c0

    # ---- 模式4：数字/字母序列 + 可选分隔符 ----
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
    style = _classify_to_style(inner_cls, ')')
    return full_prefix, style, inner_cls


def _detect_ziliao_form(stripped):
    """检测"资料N. "形式：资料 + 阿拉伯数字 + . ＋ 可选空格"""
    m = re.match(r'^资料\s*(\d+)\s*[.．]\s*', stripped)
    if not m:
        return None

    full_prefix = m.group()
    return full_prefix, '资料1. ', 'arabic'


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
    style = _classify_to_style(first_cls, sep if sep else '、')
    return full_prefix, style, first_cls


def _classify_to_style(char_type, sep):
    """根据序号字符类型和分隔符，返回Word自动编号样式名。

    统一使用 \". \"（点号+空格）作为分隔符。

    可用样式：
      - "1."   → 阿拉伯数字 + 点号
      - "一."  → 中文数字 + 点号
      - "①"   → 圈号数字
    """
    # 圈号类
    if char_type in ('circled_num', 'circled_upper', 'circled_lower'):
        return '①'

    # 中文数字 / 天干 / 地支 → 归入中文风格
    if char_type in ('chinese_num', 'heavenly_stem', 'earthly_branch'):
        return '一.'

    # 其它所有类型（阿拉伯数字、字母、罗马数字、希腊字母等）统一用点号分隔
    return '1.'


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


def _parse_prefix_value(prefix, char_type):
    """将序号前缀文本转换为对应的序号数值（1-based）。

    从检测到的序号前缀中提取核心数字部分，根据字符类别转换为整数值。

    Args:
        prefix: detect_number_prefix() 返回的序号前缀文本
                 （如"（一）"、"3. "、"①"、"IV."等）
        char_type: 序号字符类别标识（如 'chinese_num', 'arabic', 'circled_num' 等）

    Returns:
        int: 序号数值（从1开始），无法转换返回 0
    """
    if not prefix or char_type is None:
        return 0

    # ---- 提取核心数字部分 ----
    # 去除前导空白
    core = prefix.lstrip()
    # 去除括号包围
    core = re.sub(r'^[\(\（]\s*', '', core)
    core = re.sub(r'\s*[\)\）]$', '', core)
    # 去除结尾分隔符及后续空白
    core = re.sub(r'[.．,，、)）]\s*$', '', core)
    # 去除"资料"前缀（资料N. 形式）
    core = re.sub(r'^资料\s*', '', core)

    if not core:
        return 0

    try:
        # ---- 阿拉伯数字（含全角）----
        if char_type == 'arabic':
            half = core.translate(_FULL_TO_HALF_MAP)
            return int(half)

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

    def convert_all(self):
        """遍历文档所有标题段落（大纲级别1~9），将手动序号转为自动编号。

        对每个标题段落：
          1. 调用 detect_number_prefix() 检测手动序号
          2. 使用Range.Delete删除序号前缀文本
          3. 为该级别标题样式链接对应的自动编号列表模板

        Returns:
            int: 成功转换的段落数量
        """
        if self.work_doc is None:
            return 0

        para_count = self.work_doc.Paragraphs.Count
        converted = 0

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

                # ---- 对比并调整自动序号值 ----
                # 提取前缀对应的序号数值（如 "（三）" → 3）
                prefix_value = _parse_prefix_value(prefix, char_type)
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
                rng = para.Range.Duplicate
                prefix_len_in_range = len(prefix)
                # 防御：确保不超过段落文本长度（保留末尾段落标记）
                max_delete = max(1, len(text) - 1)  # 至少保留 \r
                if prefix_len_in_range > max_delete:
                    prefix_len_in_range = max_delete
                rng.End = rng.Start + prefix_len_in_range
                rng.Delete()

                # ---- 链接自动编号列表模板 ----
                # 确保每个级别只链接一次，避免重复创建列表模板
                if level not in self._level_styles_set:
                    self._link_list_template(level, style_name)
                    self._level_styles_set[level] = style_name

                converted += 1

            except Exception:
                # 忽略单个段落的异常，继续处理后续段落
                continue

        return converted

    def _link_list_template(self, level, number_style):
        """为指定级别的"标题 N"样式链接自动编号列表模板。

        先清空全部9级，再设置目标级别和第1级（安全兜底），
        使用 LinkToListTemplate 显式覆盖样式已有的列表模板链接。

        Args:
            level: 标题级别 (1~9)
            number_style: 编号样式名 ("1." / "一." / "①")
        """
        try:
            style_names = [f"标题 {level}", f"标题{level}"]
            style = None
            for sn in style_names:
                try:
                    style = self.work_doc.Styles(sn)
                    break
                except Exception:
                    continue

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
            if number_style == "一.":
                fmt = f"%{level}. "
                ns = 37  # wdListNumberStyleSimpChinNum
            elif number_style == "1.":
                fmt = f"%{level}. "
                ns = 0  # wdListNumberStyleArabic
            elif number_style == "①":
                fmt = f"%{level}"
                ns = 18  # wdListNumberStyleNumberInCircle
            elif number_style == "资料1. ":
                fmt = f"资料%{level}. "
                ns = 0
            else:
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
            try:
                style.LinkToListTemplate(
                    ListTemplate=list_template, ListLevelNumber=level)
            except Exception:
                list_template.ListLevels(level).LinkedStyle = style

        except Exception as ex:
            import traceback
            traceback.print_exc()

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
