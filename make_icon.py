# -*- coding: utf-8 -*-
"""生成程序图标 app.ico（含多种尺寸）及预览图 app_icon_preview.png

用法：
    python make_icon.py

设计说明：
    底色沿用界面主色 #1976D2（main_form.py 中所有分组标题的蓝色），
    主体为白色文档纸 + 加粗字母 A + 文本行，表达"文档格式设置"。
    ico 内置 16/24/32/48/64/128/256 七种尺寸，供资源管理器、
    任务栏、窗口标题栏在不同 DPI 下取用。

    小尺寸（<= 32px）单独走简化渲染：去掉投影、放大字母 A、减少文本行，
    避免大图直接缩小时细节糊成一团。
"""
import os

from PIL import Image, ImageDraw, ImageFilter, ImageFont

# 与界面配色保持一致
BRAND_LIGHT = (33, 150, 243)    # #2196F3 渐变上端
BRAND_DARK = (21, 101, 192)     # #1565C0 渐变下端
BRAND = (25, 118, 210)          # #1976D2 字母 A
SHEET_COLOR = (255, 255, 255)
TEXT_LINE_COLOR = (176, 190, 197)   # #B0BEC5 文本行

ICO_SIZES = [256, 128, 64, 48, 32, 24, 16]
# 小于等于该尺寸时改用简化渲染
SMALL_ICON_THRESHOLD = 32
# 母版最大绘制边长，再降采样到目标尺寸以保证边缘平滑
MAX_RENDER_SIZE = 1024

# 绘制用字体，取系统里的粗体无衬线字体；都取不到时退回几何图形
_FONT_CANDIDATES = [
    r'C:\Windows\Fonts\arialbd.ttf',
    r'C:\Windows\Fonts\seguisb.ttf',
    r'C:\Windows\Fonts\segoeuib.ttf',
]

# 布局参数：(普通尺寸, 小尺寸简化版)
_LAYOUT = {
    'normal': {
        'sheet_w': 0.50, 'sheet_h': 0.64, 'sheet_y': 0.18,
        'shadow': True,
        'a_left': 0.10, 'a_top': 0.10, 'a_right': 0.90, 'a_bottom': 0.60,
        'a_font_ratio': 0.62,
        'line_ratios': (1.0, 1.0, 0.55), 'line_h': 0.024, 'line_gap': 0.045,
        'line_top': 0.68,
    },
    'small': {
        'sheet_w': 0.62, 'sheet_h': 0.72, 'sheet_y': 0.14,
        'shadow': False,
        'a_left': 0.08, 'a_top': 0.08, 'a_right': 0.92, 'a_bottom': 0.70,
        'a_font_ratio': 0.86,
        'line_ratios': (1.0,), 'line_h': 0.062, 'line_gap': 0.0,
        'line_top': 0.78,
    },
}


def _vertical_gradient(size, top_color, bottom_color):
    """生成自上而下的线性渐变图"""
    gradient = Image.new('RGB', (1, size))
    for y in range(size):
        ratio = y / max(1, size - 1)
        gradient.putpixel((0, y), tuple(
            round(top_color[i] + (bottom_color[i] - top_color[i]) * ratio)
            for i in range(3)))
    return gradient.resize((size, size), Image.BILINEAR)


def _load_font(size):
    """加载粗体字体，失败返回 None"""
    for path in _FONT_CANDIDATES:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return None


def _draw_letter_a(draw, box, font):
    """在 box=(left, top, right, bottom) 区域内居中绘制字母 A"""
    left, top, right, bottom = box
    cx = (left + right) / 2
    if font is not None:
        bbox = draw.textbbox((0, 0), 'A', font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        draw.text((cx - w / 2 - bbox[0], top + (bottom - top - h) / 2 - bbox[1]),
                  'A', font=font, fill=BRAND)
        return

    # 取不到字体时用几何图形拼一个 A：外三角 - 内三角 + 横杠
    w = right - left
    h = bottom - top
    draw.polygon([(cx, top), (right, bottom), (left, bottom)], fill=BRAND)
    inset = w * 0.22
    draw.polygon([(cx, top + h * 0.34), (right - inset, bottom), (left + inset, bottom)],
                 fill=SHEET_COLOR)
    bar_top = top + h * 0.62
    draw.rectangle([left + w * 0.16, bar_top, right - w * 0.16, bar_top + h * 0.13],
                   fill=BRAND)


def build_icon(render_size, simplified=False):
    """绘制指定边长的图标母版（RGBA），所有几何量按边长等比换算"""
    size = render_size
    layout = _LAYOUT['small' if simplified else 'normal']
    icon = Image.new('RGBA', (size, size), (0, 0, 0, 0))

    # 1) 圆角方形底色
    margin = round(size * 0.055)
    bg_mask = Image.new('L', (size, size), 0)
    ImageDraw.Draw(bg_mask).rounded_rectangle(
        [margin, margin, size - margin, size - margin],
        radius=round(size * 0.22), fill=255)
    icon.paste(_vertical_gradient(size, BRAND_LIGHT, BRAND_DARK), (0, 0), bg_mask)

    # 2) 文档纸：先画投影，再加白纸
    sheet_w = round(size * layout['sheet_w'])
    sheet_h = round(size * layout['sheet_h'])
    sheet_x = (size - sheet_w) // 2
    sheet_y = round(size * layout['sheet_y'])
    sheet_radius = round(size * 0.032)

    if layout['shadow']:
        shadow_offset = round(size * 0.016)
        shadow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        ImageDraw.Draw(shadow).rounded_rectangle(
            [sheet_x, sheet_y + shadow_offset,
             sheet_x + sheet_w, sheet_y + sheet_h + shadow_offset],
            radius=sheet_radius, fill=(0, 0, 0, 70))
        shadow = shadow.filter(ImageFilter.GaussianBlur(round(size * 0.018)))
        icon = Image.alpha_composite(icon, shadow)

    sheet = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(sheet).rounded_rectangle(
        [sheet_x, sheet_y, sheet_x + sheet_w, sheet_y + sheet_h],
        radius=sheet_radius, fill=SHEET_COLOR + (255,))
    icon = Image.alpha_composite(icon, sheet)

    # 3) 字母 A 与文本行
    draw = ImageDraw.Draw(icon)
    _draw_letter_a(
        draw,
        (sheet_x + sheet_w * layout['a_left'], sheet_y + sheet_h * layout['a_top'],
         sheet_x + sheet_w * layout['a_right'], sheet_y + sheet_h * layout['a_bottom']),
        _load_font(round(sheet_w * layout['a_font_ratio'])))

    line_h = max(1, round(size * layout['line_h']))
    line_gap = max(line_h, round(size * layout['line_gap']))
    line_x = sheet_x + sheet_w * 0.14
    line_right = sheet_x + sheet_w * 0.86
    line_y = sheet_y + sheet_h * layout['line_top']
    for ratio in layout['line_ratios']:
        draw.rounded_rectangle(
            [line_x, line_y, line_x + (line_right - line_x) * ratio, line_y + line_h],
            radius=line_h / 2, fill=TEXT_LINE_COLOR)
        line_y += line_gap

    return icon


def render_size(px):
    """渲染目标边长的成品图，内部先放大再降采样以抗锯齿"""
    simplified = px <= SMALL_ICON_THRESHOLD
    # 小尺寸本身很小，需要足够高的倍率才能把圆角和字母画平滑
    supersample = 4 if simplified else min(4, MAX_RENDER_SIZE // px)
    master = build_icon(px * supersample, simplified)
    return master.resize((px, px), Image.LANCZOS)


def save_outputs(out_dir):
    """输出 app.ico 与预览图，返回 (ico路径, 预览图路径)"""
    # 每个尺寸单独渲染：Pillow 的 ICO 保存会优先取同尺寸的 append_images
    frames = {px: render_size(px) for px in ICO_SIZES}
    base = frames[ICO_SIZES[0]]
    ico_path = os.path.join(out_dir, 'app.ico')
    base.save(ico_path, format='ICO', sizes=[(px, px) for px in ICO_SIZES],
              append_images=[frames[px] for px in ICO_SIZES[1:]])

    # 预览图：浅色与深色背景各一行，便于判断小尺寸下的辨识度
    padding = 16
    cell_w = 256 + padding
    width = cell_w * len(ICO_SIZES) + padding
    height = 256 * 2 + padding * 3
    preview = Image.new('RGB', (width, height), (245, 245, 245))
    dark_top = 256 + padding * 2
    ImageDraw.Draw(preview).rectangle(
        [0, dark_top - padding // 2, width, height], fill=(45, 45, 48))

    for index, px in enumerate(ICO_SIZES):
        small = frames[px]
        cell_x = padding + index * cell_w + (256 - px) // 2
        preview.paste(small, (cell_x, padding + (256 - px) // 2), small)
        preview.paste(small, (cell_x, dark_top + (256 - px) // 2), small)

    preview_path = os.path.join(out_dir, 'app_icon_preview.png')
    preview.save(preview_path)
    return ico_path, preview_path


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    ico_path, preview_path = save_outputs(out_dir)
    print(f'已生成图标: {ico_path}')
    print(f'已生成预览: {preview_path}')


if __name__ == '__main__':
    main()
