#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""国企PPT制作 · python-pptx 四页脚手架（可直接运行）

四页示例，同时验证两套视觉体系：
  01 深蓝体系封面        —— 路线 A（部门汇报）
  02 深蓝体系内容页      —— 大数字 + 故事矩阵 + 媒体位（可选嵌入 mp4）
  03 党建朱墨章节页      —— 红标题带 + 细金线 + 大留白
  04 党建朱墨尾页        —— 深朱底 + 原稿收口句

用法：
  python3 build_demo.py
  python3 build_demo.py --out /tmp/demo.pptx
  python3 build_demo.py --video /path/to/clip.mp4   # 第 2 页媒体位改为视频

依赖：python-pptx（DSH 捆绑运行时 1.0.2）+ lxml；不依赖网络、图片素材。
"""

from __future__ import annotations

import argparse
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# 画幅与两套色板（色值取自真实项目，已脱敏）
# ---------------------------------------------------------------------------
W, H = 13.333, 7.5

# 深蓝体系
NAVY = RGBColor(0x0B, 0x2B, 0x4E)
BLUE = RGBColor(0x00, 0x5D, 0xA2)
BLUE2 = RGBColor(0x2E, 0x86, 0xC8)
LIGHT = RGBColor(0xE6, 0xF2, 0xFA)
GOLD = RGBColor(0xFF, 0xC0, 0x00)
GRAY = RGBColor(0x5B, 0x6B, 0x7A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# 党建朱墨
PAPER = RGBColor(0xF6, 0xF2, 0xEA)
RED = RGBColor(0xA3, 0x2B, 0x2B)
RED_DK = RGBColor(0x7E, 0x1F, 0x1F)
GOLD2 = RGBColor(0xA9, 0x8A, 0x4B)
INK = RGBColor(0x1E, 0x1E, 0x1E)
INK2 = RGBColor(0x4A, 0x47, 0x41)
LINE = RGBColor(0xD8, 0xD0, 0xC0)
FADE = RGBColor(0xE9, 0xE0, 0xCF)

FONT = "微软雅黑"
SERIF = "宋体"
KAI = "楷体"


# ---------------------------------------------------------------------------
# 基础函数：所有页面共用，复制到新项目时先保留这一段
# ---------------------------------------------------------------------------
def _set_run(run, size, color, bold=False, font=FONT):
    """设置中文 run：a:latin 之外必须补 a:ea / a:cs，否则部分机器回退字体。"""
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.name = font
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)


def _kill_style_effects(shape):
    """去掉主题样式里的阴影/描边效果，避免 LibreOffice/部分 PowerPoint 出现默认阴影。"""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        eff = style.find(qn("a:effectRef"))
        if eff is not None:
            style.remove(eff)


def add_text(slide, x, y, w, h, text, size=14, color=INK, bold=False,
             font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.2, name=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        _set_run(r, size, color, bold, font)
    return box


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75,
             dash=None, radius=None):
    st = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(st, Inches(x), Inches(y), Inches(w), Inches(h))
    _kill_style_effects(shp)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
        if dash is not None:
            shp.line.dash_style = dash
    if radius is not None:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    return shp


def add_oval(slide, x, y, w, h, fill=None, line=None, line_w=0.75):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    _kill_style_effects(shp)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    return shp


def add_line(slide, x1, y1, x2, y2, color=LINE, width=0.75, dash=None):
    cn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    _kill_style_effects(cn)
    cn.line.color.rgb = color
    cn.line.width = Pt(width)
    if dash is not None:
        cn.line.dash_style = dash
    return cn


def solid_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def gradient_bg(slide, color_a, color_b, angle=45):
    """深色封面/尾页渐变；失败时退回纯色，保证脚本不崩。"""
    fill = slide.background.fill
    try:
        fill.gradient()
        fill.gradient_angle = angle
        stops = fill.gradient_stops
        stops[0].color.rgb = color_a
        stops[0].position = 0.0
        stops[1].color.rgb = color_b
        stops[1].position = 1.0
    except Exception:
        fill.solid()
        fill.fore_color.rgb = color_a


def add_footer(slide, text, color=GRAY, line_color=LINE):
    """页脚；页码先放占位，最后按实际页序统一回填。"""
    add_line(slide, 0.7, 6.98, 12.63, 6.98, color=line_color, width=0.75)
    add_text(slide, 0.7, 7.04, 8.4, 0.24, text, size=8.5, color=color)
    pn = add_text(slide, 11.8, 7.04, 0.83, 0.24, "00", size=9, color=color,
                  align=PP_ALIGN.RIGHT)
    pn.name = "PAGE_NO"
    return pn


def add_media_slot(slide, x, y, w, h, label="配图位", note="公司实景：船 / 码头 / 岸桥"):
    add_rect(slide, x, y, w, h, fill=LIGHT, line=BLUE2, line_w=1.0, dash=MSO_LINE.DASH)
    add_text(slide, x + 0.2, y + h / 2 - 0.35, w - 0.4, 0.3, label,
             size=11, color=BLUE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, x + 0.25, y + h / 2 + 0.02, w - 0.5, 0.6, note,
             size=9.5, color=GRAY, align=PP_ALIGN.CENTER, line_spacing=1.3)


def add_video(slide, path, x, y, w, h):
    """嵌入 mp4（H.264/AAC）；poster 可传封面图，这里省略。"""
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    return slide.shapes.add_movie(
        path, Inches(x), Inches(y), Inches(w), Inches(h), mime_type="video/mp4")


# ---------------------------------------------------------------------------
# 页面 01：深蓝体系封面
# ---------------------------------------------------------------------------
def page_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    gradient_bg(s, NAVY, BLUE, angle=315)

    # 右上角同心圆装饰（半出画）
    add_oval(s, 9.4, -1.6, 6.2, 6.2, fill=None, line=BLUE2, line_w=1.2)
    add_oval(s, 10.4, -0.6, 4.2, 4.2, fill=None, line=NAVY, line_w=1.0)
    add_oval(s, 11.2, 0.2, 2.6, 2.6, fill=None, line=GOLD, line_w=1.5)

    # logo 占位（替换成公司 logo 彩色/白色反白图）
    add_rect(s, 0.62, 0.42, 1.05, 0.64, fill=None, line=WHITE, line_w=0.9)
    add_text(s, 0.62, 0.56, 1.05, 0.3, "LOGO", size=8, color=WHITE,
             bold=True, align=PP_ALIGN.CENTER)

    add_text(s, 0.98, 6.52, 5.6, 0.3, "某年度务虚会 · 某部门",
             size=10, color=LIGHT)

    add_text(s, 0.95, 2.18, 8.9, 0.9, "年度工作汇报", size=44, color=WHITE,
             bold=True, line_spacing=1.05)
    add_text(s, 0.98, 3.18, 8.6, 0.55, "——在新格局下找准部门工作新坐标",
             size=20, color=LIGHT, line_spacing=1.15)
    add_line(s, 1.0, 4.28, 2.5, 4.28, color=GOLD, width=2.0)
    add_text(s, 0.98, 4.52, 5.4, 0.34, "某年 11 月 · 会议地", size=11.5, color=LIGHT)
    add_text(s, 10.1, 6.55, 2.6, 0.28, "16:9 · 演讲版", size=9.5, color=LIGHT,
             align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------------------
# 页面 02：深蓝体系内容页（大数字 + 故事矩阵 + 媒体位）
# ---------------------------------------------------------------------------
def page_content(prs, video=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    solid_bg(s, WHITE)

    add_text(s, 0.7, 0.42, 6.0, 0.26, "第二部分 · 持续改进", size=9.5, color=BLUE, bold=True)
    add_text(s, 0.7, 0.7, 7.8, 0.6, "某年，我们把发展故事讲到集团、公司",
             size=27, color=NAVY, bold=True, line_spacing=1.05)
    add_line(s, 0.72, 1.42, 1.62, 1.42, color=GOLD, width=2.5)

    # 左侧大数字：全篇只做一次重点
    add_text(s, 0.7, 1.62, 4.2, 1.2, "关键指标", size=72, color=BLUE, bold=True,
             line_spacing=0.95)
    add_text(s, 0.74, 2.9, 4.4, 0.3, "集团、公司外宣平台刊稿量同比", size=10.5, color=GRAY)

    # 2×2 故事矩阵：无卡片框，细线 + 序号 + 关键词
    stories = [
        ("01", "原创内容案例 A", "多家党组织、多名员工参与拍摄"),
        ("02", "案例 B", "与业务部门一起设计产品海报"),
        ("03", "案例 C", "发展故事被集团、公司快速转发"),
        ("04", "案例 D", "讲出属于区域分部的产业故事"),
    ]
    x0, y0, col_w, row_h = 0.7, 3.55, 3.55, 1.25
    for i, (num, head, note) in enumerate(stories):
        r, c = divmod(i, 2)
        x = x0 + c * (col_w + 0.25)
        y = y0 + r * row_h
        add_line(s, x, y, x + col_w, y, color=LINE, width=0.6)
        add_text(s, x, y + 0.12, 0.4, 0.24, num, size=9, color=GOLD, bold=True)
        add_text(s, x + 0.45, y + 0.08, col_w - 0.45, 0.3, head, size=13.5, color=NAVY, bold=True)
        add_text(s, x + 0.45, y + 0.45, col_w - 0.45, 0.5, note, size=9.2, color=GRAY,
                 line_spacing=1.25)

    # 右侧媒体位：有 --video 就嵌视频，没有就放配图占位
    if video:
        add_video(s, video, 8.62, 1.62, 4.05, 2.28)
        add_text(s, 8.62, 4.0, 4.05, 0.3, "一线有故事 · 约 30 秒", size=9.5, color=GRAY)
        add_text(s, 8.62, 4.4, 4.05, 1.2,
                 "视频单独成页或占据大画面，\n不要和三条内容挤在一页。",
                 size=10.5, color=INK2, line_spacing=1.35)
    else:
        add_media_slot(s, 8.62, 1.6, 4.05, 4.9, "媒体位", "配图：公司实景照片\n或在这里嵌入 mp4")

    add_text(s, 0.74, 6.28, 11.9, 0.42,
             "宣传不是“锦上添花”，而是凝聚力量、推动发展、提升形象的实功",
             size=12.5, color=BLUE, bold=True, line_spacing=1.1)
    add_footer(s, "年度工作汇报 · 某年度务虚会汇报")


# ---------------------------------------------------------------------------
# 页面 03：党建朱墨章节页（红标题带）
# ---------------------------------------------------------------------------
def page_section(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    solid_bg(s, PAPER)

    # 红标题带 + 金线
    add_rect(s, 0, 0, W, 1.98, fill=RED)
    add_rect(s, 0, 1.98, W, 0.045, fill=GOLD2)
    add_text(s, 0.98, 0.36, 4.0, 0.28, "第三部分 · 传达学习《决定》",
             size=9.5, color=RGBColor(0xF2, 0xD9, 0xA8), bold=True)
    add_text(s, 0.96, 0.72, 9.7, 0.66, "《决定》：成果 · 起草 · 三大板块",
             size=31, color=WHITE, bold=True, font=SERIF, line_spacing=1.0)
    add_text(s, 10.9, 0.48, 1.5, 1.1, "03", size=58, color=WHITE, bold=True,
             font=SERIF, align=PP_ALIGN.RIGHT, line_spacing=0.9)

    # 结论句
    add_text(s, 0.98, 2.52, 7.9, 1.2,
             "最重要的成果，是审议通过了《决定》",
             size=28, color=INK, bold=True, font=SERIF, line_spacing=1.15)
    add_line(s, 1.0, 3.95, 1.9, 3.95, color=GOLD2, width=2.0)

    # 右栏：三条结构，编号 + 细线
    rows = [("01", "第一板块", "总论"),
            ("02", "第二板块", "分论 · 7 个部分"),
            ("03", "第三板块", "监督与责任")]
    x, y, w = 8.55, 2.62, 3.9
    for num, head, note in rows:
        add_text(s, x, y, 0.45, 0.26, num, size=9, color=GOLD2, bold=True)
        add_text(s, x + 0.52, y - 0.04, w - 0.52, 0.32, head, size=14.5, color=RED, bold=True, font=SERIF)
        add_text(s, x + 0.52, y + 0.32, w - 0.52, 0.28, note, size=10, color=INK2)
        add_line(s, x, y + 0.76, x + w, y + 0.76, color=LINE, width=0.75)
        y += 0.95

    add_text(s, 0.98, 6.28, 8.0, 0.3, "三大板块构成《决定》全文",
             size=9.5, color=INK2, font=KAI)
    add_footer(s, "学习党的上级重要会议精神精神", color=INK2, line_color=LINE)


# ---------------------------------------------------------------------------
# 页面 04：党建朱墨尾页
# ---------------------------------------------------------------------------
def page_ending(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    gradient_bg(s, RGBColor(0x8C, 0x15, 0x13), RGBColor(0xB2, 0x21, 0x1E), angle=90)

    # 细金线边框（留白 0.42 英寸）
    add_rect(s, 0.42, 0.42, W - 0.84, H - 0.84, fill=None, line=GOLD2, line_w=0.9)

    add_text(s, 1.2, 0.9, 6.0, 0.28, "学习党的上级重要会议精神精神",
             size=9.5, color=RGBColor(0xF2, 0xD9, 0xA8))
    add_text(s, 1.2, 2.88, 10.8, 0.8, "持之以恒推进全面从严治党",
             size=40, color=WHITE, bold=True, font=SERIF, line_spacing=1.0)
    add_line(s, 1.22, 3.98, 2.75, 3.98, color=GOLD2, width=2.0)
    add_text(s, 1.2, 4.22, 6.0, 0.34, "某年 11 月", size=12.5,
             color=RGBColor(0xE8, 0xD9, 0xC8))
    add_text(s, 9.9, 6.5, 2.4, 0.28, "16:9 · 报告版", size=9.5,
             color=RGBColor(0xE8, 0xD9, 0xC8), align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------------------
# 组装 + 动态页码
# ---------------------------------------------------------------------------
def build(video=None):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    page_cover(prs)
    page_content(prs, video=video)
    page_section(prs)
    page_ending(prs)

    # 页码在全部页面排好序后统一回填：增删页不会错号
    for idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.name == "PAGE_NO" and shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.runs:
                        p.runs[0].text = f"{idx:02d}"
                        break
    return prs


def main(argv=None):
    parser = argparse.ArgumentParser(description="生成国企PPT四页 demo")
    default_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo.pptx")
    parser.add_argument("--out", default=default_out, help="输出 pptx 路径")
    parser.add_argument("--video", default=None, help="可选：嵌入第 2 页媒体位的 mp4")
    args = parser.parse_args(argv)

    prs = build(video=args.video)
    prs.save(args.out)

    # 回读校验，确认文件真的能打开
    check = Presentation(args.out)
    size_kb = os.path.getsize(args.out) / 1024
    print(f"saved: {args.out}")
    print(f"slides: {len(check.slides)}  size: {W}×{H} in  file: {size_kb:.1f} KB")
    if args.video:
        print(f"video: {args.video}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
