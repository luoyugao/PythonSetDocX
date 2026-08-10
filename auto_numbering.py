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
        (prefix, style_name, char_type, is_first) 四元组：
        - prefix: str — 检测到的序号前缀文本
        - style_name: str — Word自动编号样式名（"1." / "一." / "①"）
        - char_type: str — 序号字符类别标识
        - is_first: bool — 是否为该序号体系中的第一个元素
        若未检测到序号则返回 (None, None, None, None)
    """
    if not text:
        return None, None, None, None

    # 去除末尾段落标记（\r、\x07 等Word内部字符）
    clean = text.rstrip('\r\n\x07')
    if not clean:
        return None, None, None, None

    # 去除开头的空白，记录空白长度
    stripped = clean.lstrip()
    leading_ws_len = len(clean) - len(stripped)

    if not stripped:
        return None, None, None, None

    result = None

    # ---- 模式1：括号包围形式  (1) （一） (a) （Ａ） 等 ----
    result = _detect_paren_form(stripped)
    if result:
        inner_prefix, style, ctype, is_first = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype, is_first

    # ---- 模式2：资料N. 特殊前缀 ----
    result = _detect_ziliao_form(stripped)
    if result:
        inner_prefix, style, ctype, is_first = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype, is_first

    # ---- 模式3：圈号单字符序号（①②③…、ⓐⓑⓒ…） ----
    c0 = _classify(stripped[0])
    if c0 in ('circled_num', 'circled_upper', 'circled_lower'):
        prefix = clean[:leading_ws_len] + stripped[0]
        is_first = _is_first_in_sequence(stripped[0], c0)
        return prefix, '①', c0, is_first

    # ---- 模式4：数字/字母序列 + 可选分隔符 ----
    result = _detect_num_sep_form(stripped)
    if result:
        inner_prefix, style, ctype, is_first = result
        prefix = clean[:leading_ws_len] + inner_prefix
        return prefix, style, ctype, is_first

    return None, None, None, None


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
    is_first = _is_first_in_sequence(inner, inner_cls)
    return full_prefix, style, inner_cls, is_first


def _detect_ziliao_form(stripped):
    """检测"资料N. "形式：资料 + 阿拉伯数字 + . ＋ 可选空格"""
    m = re.match(r'^资料\s*(\d+)\s*[.．]\s*', stripped)
    if not m:
        return None

    digits = m.group(1)
    full_prefix = m.group()
    is_first = _is_first_in_sequence(digits, 'arabic')
    return full_prefix, '资料1. ', 'arabic', is_first


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
    is_first = _is_first_in_sequence(num_part, first_cls)
    return full_prefix, style, first_cls, is_first


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


_FIRST_IN_SEQUENCE = {
    'arabic':        '1',
    'chinese_num':   '一',
    'heavenly_stem': '甲',
    'earthly_branch':'子',
    'upper_alpha':   'A',
    'lower_alpha':   'a',
    'roman_upper':   'Ⅰ',
    'roman_lower':   'ⅰ',
    'greek_upper':   'Α',
    'greek_lower':   'α',
    'circled_num':   '①',
    'circled_upper': 'Ⓐ',
    'circled_lower': 'ⓐ',
}


def _is_first_in_sequence(num_text, char_type):
    """判断序号文本是否为其体系中的第一个元素。

    对于字母类和阿拉伯数字，必须归一化为半角后再比较。
    """
    if not num_text or char_type is None:
        return False
    first = _FIRST_IN_SEQUENCE.get(char_type)
    if first is None:
        return False
    # 字母/数字类：归一化后比较；中文/罗马等：直接比较
    if char_type in ('arabic', 'upper_alpha', 'lower_alpha'):
        return _to_half(num_text) == _to_half(first)
    return num_text == first


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
        # 记录已链接过模板的级别，避免重复 link
        self._level_linked = set()
        # 共享的多级列表模板（所有标题级别共用，使 ResetOnHigher 在同一列表实例内生效）
        self._master_template = None

    def _parse_level(self, para):
        """从段落中提取标题级别，失败返回 None。

        优先用 OutlineLevel；若未设置则通过样式名（"标题 N"）推导。
        """
        level = para.OutlineLevel
        if 1 <= level <= 5:
            return level
        try:
            from word_constants import get_range_style
            sn = get_range_style(para).NameLocal
            if '标题' not in sn:
                return None
            import re
            m = re.search(r'(\d+)', sn)
            if m:
                level = int(m.group(1))
                return level if 1 <= level <= 5 else None
        except Exception:
            pass
        return None

    def convert_all(self):
        """遍历文档所有标题段落（大纲级别1~5），将手动序号转为自动编号。

        策略（用户建议方案）：
          1. 第一遍：删除手动序号前缀 + 为各级标题样式链接自动编号列表模板
                    链接后 Word 自动为所有同样式标题生成连续递增的自动编号
          2. 第二遍：将原文序号为"1"（体系首元素）的段落，通过 ListValue=1
                     强制重新开始编号，断开与前列段落的连续关系

        Returns:
            int: 成功转换的段落数量
        """
        if self.work_doc is None:
            return 0

        para_count = self.work_doc.Paragraphs.Count
        # 记录需要重新开始编号的段落：(paragraph_index, level)
        restart_paragraphs = []
        para_converted = 0

        # ---- 第一遍：删除前缀 + 链接样式模板 ----
        for i in range(1, para_count + 1):
            try:
                para = self.work_doc.Paragraphs(i)
                level = self._parse_level(para)
                if level is None:
                    continue
                text = para.Range.Text
                if not text or not text.strip():
                    continue
                prefix, style_name, _, is_first = detect_number_prefix(text)
                if prefix is None:
                    continue

                # ---- 删除序号前缀文本 ----
                rng = para.Range.Duplicate
                p_len = min(len(prefix), max(1, len(text) - 1))
                rng.End = rng.Start + p_len
                rng.Delete()

                # ---- 链接样式到列表模板（每级别仅执行一次） ----
                if level not in self._level_linked:
                    self._link_list_template_level(level, style_name)
                    self._level_linked.add(level)

                # ---- 记录原文为首元素的段落，稍后重启编号 ----
                if is_first:
                    restart_paragraphs.append((i, level))

                para_converted += 1

            except Exception:
                continue

        if para_converted == 0:
            return 0

        # ---- 第二遍：对原文为首元素的段落，强制重新开始编号 ----
        # style.LinkToListTemplate 建立后，Word 已为所有同样式标题生成
        # 连续递增的自动编号。此处通过设置 ListValue=1 将指定段落强制
        # 设为编号起点，等效于 Word 界面中的"重新开始编号"右键命令。
        if restart_paragraphs:
            for para_idx, _level in restart_paragraphs:
                try:
                    para = self.work_doc.Paragraphs(para_idx)
                    # 防御：仅当段落确实处于编号列表中时才设置 ListValue
                    if para.Range.ListFormat.ListType != 0:  # 0 = wdListNoNumbering
                        para.Range.ListFormat.ListValue = 1
                except Exception:
                    pass

        return para_converted

    def _link_list_template_level(self, level, number_style):
        """配置共享列表模板的指定级别，并将标题样式链接至该模板。

        所有标题级别共用同一个 ListTemplate，使 ResetOnHigher 能在
        上级标题出现时自动重置子级编号。
        首次调用时创建模板并清空全部9级，后续调用仅配置目标级别。
        """
        try:
            # 获取标题样式
            style = None
            for sn in (f"标题 {level}", f"标题{level}"):
                try:
                    style = self.work_doc.Styles(sn)
                    break
                except Exception:
                    continue
            if style is None:
                return

            # 首次调用时创建共享模板
            if self._master_template is None:
                self._master_template = self.work_doc.ListTemplates.Add(True)
                for lvl in range(1, 10):
                    try:
                        lo = self._master_template.ListLevels(lvl)
                        lo.NumberFormat = ""
                        lo.NumberStyle = 255  # wdListNumberStyleNone
                    except:
                        pass

            # --- 确定 NumberFormat 与 NumberStyle ---
            # 统一使用 ". "（点号+空格）作为序号分隔符
            if number_style == "一.":
                fmt = f"%{level}."
                ns = 37  # wdListNumberStyleSimpChinNum
            elif number_style == "1.":
                fmt = f"%{level}."
                ns = 0   # wdListNumberStyleArabic
            elif number_style == "①":
                fmt = f"%{level}"
                ns = 18  # wdListNumberStyleNumberInCircle
            elif number_style == "资料1. ":
                fmt = f"资料%{level}."
                ns = 0
            else:
                return

            target_level = self._master_template.ListLevels(level)
            target_level.NumberFormat = fmt
            target_level.NumberStyle = ns
            # 从图库模板获取通用布局参数
            try:
                list_gallery = wc.get_list_gallery(
                    self.work_doc.Application, wc.wdOutlineNumberGallery)
                ref_level = list_gallery.ListTemplates(1).ListLevels(1)
                for attr in ('NumberPosition', 'Alignment', 'TrailingCharacter',
                             'TabPosition', 'ResetOnHigher', 'StartAt'):
                    try:
                        setattr(target_level, attr, getattr(ref_level, attr))
                    except Exception:
                        pass
            except Exception:
                pass
            # 强制覆写关键属性
            try:
                target_level.TrailingCharacter = wc.wdTrailingSpace
            except Exception:
                pass
            try:
                target_level.StartAt = 1
            except Exception:
                pass
            try:
                target_level.ResetOnHigher = True
            except Exception:
                pass

            # 链接样式到共享模板的对应级别
            try:
                style.LinkToListTemplate(
                    ListTemplate=self._master_template, ListLevelNumber=level)
            except Exception:
                self._master_template.ListLevels(level).LinkedStyle = style

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
