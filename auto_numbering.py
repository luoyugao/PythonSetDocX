"""
auto_numbering.py
标题手动序号转Word自动序号模块

功能：
  遍历文档中大纲级别为1~5的标题段落，检测人工输入的手动序号文本
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
        """遍历文档所有标题段落（大纲级别1~5），将手动序号转为自动编号。

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

                # 只处理大纲级别1~5的标题段落，同时兼容通过样式名识别的标题
                level = para.OutlineLevel
                if level < 1 or level > 5:
                    # 回退：通过样式名判断（部分文档标题可能未正确设置大纲级别）
                    try:
                        from word_constants import get_range_style
                        style_obj = get_range_style(para)
                        style_name = style_obj.NameLocal
                        if '标题' not in style_name:
                            continue
                        # 提取标题级别数字，如 "标题 1" → 1
                        import re
                        m = re.search(r'(\d+)', style_name)
                        if m:
                            level = int(m.group(1))
                            if level < 1 or level > 5:
                                continue
                        else:
                            continue
                    except Exception:
                        continue

                text = para.Range.Text
                if not text or not text.strip():
                    continue

                # 检测手动序号前缀
                prefix, style_name, _ = detect_number_prefix(text)
                if prefix is None:
                    continue

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
            level: 标题级别 (1~5)
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
            # 统一使用 \". \"（点号+空格）作为序号分隔符
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
                    for attr in ('NumberPosition', 'Alignment', 'TrailingCharacter',
                                 'TabPosition', 'ResetOnHigher', 'StartAt'):
                        try:
                            setattr(target_level, attr, getattr(ref_level, attr))
                        except Exception:
                            pass
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
    converter = AutoNumbering(work_doc)
    return converter.convert_all()
