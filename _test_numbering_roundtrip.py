# -*- coding: utf-8 -*-
"""往返一致性测试：检测手动序号 → 生成自动编号格式 → 校验形式是否保持。

运行： python _test_numbering_roundtrip.py
"""
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import auto_numbering as an
import word_constants as wc

# 每个用例：(标题文本, 期望的体系标识前缀, 期望的 NumberFormat, 期望的 NumberStyle, 期望数值)
# NumberFormat / NumberStyle 按 level=1 校验。
CASES = [
    # ---- 阿拉伯数字（半角）----
    ("1. 项目背景",        'arabic:sep',   "%1. ",  wc.wdListNumberStyleArabic, 1),
    ("12、项目背景",       'arabic:sep',   "%1. ",  wc.wdListNumberStyleArabic, 12),
    ("3) 项目背景",        'arabic:rparen', "%1)",  wc.wdListNumberStyleArabic, 3),
    # ---- 全角阿拉伯数字 ----
    ("１．项目背景",        'arabic_full:sep', "%1. ",  wc.wdListNumberStyleArabicFullWidth, 1),
    # ---- 中文小写数字 ----
    ("一、项目背景",        'chinese_num:sep', "%1. ",  wc.wdListNumberStyleSimpChinNum1, 1),
    ("十二、项目背景",      'chinese_num:sep', "%1. ",  wc.wdListNumberStyleSimpChinNum1, 12),
    ("一. 项目背景",        'chinese_num:sep', "%1. ",  wc.wdListNumberStyleSimpChinNum1, 1),
    # ---- 括号包围 ----
    ("（一）项目背景",      'chinese_num:paren', "(%1)", wc.wdListNumberStyleSimpChinNum1, 1),
    ("(1) 项目背景",        'arabic:paren', "(%1)",  wc.wdListNumberStyleArabic, 1),
    ("(a) 项目背景",        'lower_alpha:paren', "(%1)", wc.wdListNumberStyleLowercaseLetter, 1),
    # ---- 英文大小写字母 ----
    ("A. 项目背景",         'upper_alpha:sep', "%1. ",  wc.wdListNumberStyleUppercaseLetter, 1),
    ("c) 项目背景",         'lower_alpha:rparen', "%1)", wc.wdListNumberStyleLowercaseLetter, 3),
    # ---- 全角字母 ----
    ("Ａ．项目背景",        'upper_alpha:sep', "%1. ",  wc.wdListNumberStyleUppercaseLetter, 1),
    # ---- 罗马数字 ----
    ("Ⅰ. 项目背景",         'roman_upper:sep', "%1. ",  wc.wdListNumberStyleUppercaseRoman, 1),
    ("Ⅳ、项目背景",         'roman_upper:sep', "%1. ",  wc.wdListNumberStyleUppercaseRoman, 4),
    ("ⅲ. 项目背景",         'roman_lower:sep', "%1. ",  wc.wdListNumberStyleLowercaseRoman, 3),
    # ---- 希腊字母 ----
    ("α. 项目背景",         'greek_lower:sep', "%1. ",  wc.wdListNumberStyleLowercaseGreek, 1),
    ("Β、项目背景",         'greek_upper:sep', "%1. ",  wc.wdListNumberStyleUppercaseGreek, 2),
    # ---- 天干 / 地支 ----
    ("甲、项目背景",        'heavenly_stem:sep', "%1. ", wc.wdListNumberStyleZodiac1, 1),
    ("丙、项目背景",        'heavenly_stem:sep', "%1. ", wc.wdListNumberStyleZodiac1, 3),
    ("子、项目背景",        'earthly_branch:sep', "%1. ", wc.wdListNumberStyleZodiac2, 1),
    # ---- 圈号 ----
    ("①项目背景",           'circled_num',  "%1",    wc.wdListNumberStyleNumberInCircle, 1),
    ("⑤项目背景",           'circled_num',  "%1",    wc.wdListNumberStyleNumberInCircle, 5),
    ("ⓐ项目背景",           'circled_num', "%1",   wc.wdListNumberStyleNumberInCircle, 1),
    # ---- 章序号 / 文字前缀 ----
    ("第1章 项目背景",      '第1章',        "第%1章 ", wc.wdListNumberStyleArabic, 1),
    ("第一章 项目背景",     '第1章',        "第%1章 ", wc.wdListNumberStyleArabic, 1),
    ("资料1. 项目背景",     '资料1. ',      "资料%1. ", wc.wdListNumberStyleArabic, 1),
    ("步骤2. 项目背景",     '步骤1. ',      "步骤%1. ", wc.wdListNumberStyleArabic, 2),
]


def main():
    passed = failed = 0
    print("=" * 96)
    print("往返一致性测试（level=1）")
    print("=" * 96)
    for text, want_base, want_fmt, want_ns, want_val in CASES:
        prefix, style, ctype = an.detect_number_prefix(text)
        if prefix is None:
            print("[FAIL] 未检测到序号      : %r" % text)
            failed += 1
            continue

        got_fmt, got_ns = an._number_format_for_style(style, 1)
        values = an.parse_prefix_values(prefix, ctype)
        got_val = values[-1] if values else 0

        # 期望值写的是完整体系标识（含 :paren / :rparen / :sep 后缀），
        # 与检测结果整体比对
        got_base = style or ''

        ok = True
        reasons = []
        # 圈号哨兵 '①' 归一化为 "circled_num" 语义比对
        want_cmp = want_base
        got_cmp = 'circled_num' if got_base == '①' else got_base
        if want_cmp == 'circled_num' and got_base == '①':
            got_cmp = 'circled_num'
        if got_cmp != want_cmp:
            ok = False
            reasons.append("标识 %r != %r" % (got_cmp, want_cmp))
        if got_fmt != want_fmt:
            ok = False
            reasons.append("格式 %r != %r" % (got_fmt, want_fmt))
        if got_ns != want_ns:
            ok = False
            reasons.append("样式 %r != %r" % (got_ns, want_ns))
        if got_val != want_val:
            ok = False
            reasons.append("数值 %r != %r" % (got_val, want_val))

        if ok:
            passed += 1
            print("[ OK ] %-14r → prefix=%-10r fmt=%-9r ns=%-4s val=%s"
                  % (text, prefix, got_fmt, got_ns, got_val))
        else:
            failed += 1
            print("[FAIL] %-14r → %s" % (text, "; ".join(reasons)))
            print("        prefix=%r style=%r ctype=%r fmt=%r ns=%r val=%r"
                  % (prefix, style, ctype, got_fmt, got_ns, got_val))

    # ---- 多级序号：验证各级形式保持 ----
    print()
    print("=" * 96)
    print("多级序号级联测试")
    print("=" * 96)
    ml_cases = [
        ("3.1 项目背景", 2),
        ("１．２ 项目背景", 2),
    ]
    for text, lvl in ml_cases:
        prefix, style, ctype = an.detect_number_prefix(text)
        fmt, ns = an._number_format_for_style(style, lvl)
        values = an.parse_prefix_values(prefix, ctype)
        ok = style == an.MULTI_LEVEL_STYLE and ns == wc.wdListNumberStyleArabic
        print("[%s] %-16r → style=%r fmt=%r ns=%r values=%r"
              % (" OK " if ok else "FAIL", text, style, fmt, ns, values))
        if ok:
            passed += 1
        else:
            failed += 1

    # ---- 级联可行性判定 ----
    print()
    print("=" * 96)
    print("级联可行性判定（多级序号中的上级体系）")
    print("=" * 96)
    cascade_cases = [
        ('arabic:sep', False),
        ('chinese_num:sep', False),
        ('upper_alpha:sep', False),
        ('roman_upper:sep', False),
        ('heavenly_stem:sep', False),
        ('circled_num', True),      # 圈号无法级联
        ('circled_lower', True),
    ]
    for style, want_cannot in cascade_cases:
        got = an._style_cannot_cascade(style)
        ok = got == want_cannot
        print("[%s] %-20r cannot_cascade=%s (期望 %s)"
              % (" OK " if ok else "FAIL", style, got, want_cannot))
        if ok:
            passed += 1
        else:
            failed += 1

    print()
    print("=" * 96)
    print("结果：%d 通过 / %d 失败" % (passed, failed))
    print("=" * 96)
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
