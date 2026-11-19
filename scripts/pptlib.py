# -*- coding: utf-8 -*-
"""python-pptx 公共函数：样式常量、文本框、面板、标签、表格、图片裁切、框高估算。

用法：新汇报复制本文件，只改下面“样式常量”一段（数值来自 可调设置/默认样式表.md 或
extract_style.py 的统计结果）；逐页脚本 from pptlib import * 后调用。
范例见 examples/02_课上文献汇报/脚本/build2.py。
约定：文字里 **x** 或 〖x〗 渲染为红色加粗（强调数字、关键词）。
"""
import re, os
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

# ---- 样式常量（按样式表改）----
FONT = '微软雅黑'
GREEN, GREEN_D, GREEN_L = '00873B', '006B2F', 'EAF4EE'
TEXT, RED, CAP, GRAY = '222222', 'C00000', '595959', '7F7F7F'
SZ_TITLE, SZ_KEY, SZ_LABEL, SZ_BODY, SZ_CAP, SZ_SRC, SZ_CONC = 28, 20, 16, 16, 12, 10, 20
MARGIN = 0.9
SLIDE_W, SLIDE_H = 33.867, 19.05
CW = SLIDE_W - 2 * MARGIN
LINE_Y = 2.05          # 页眉横线位置 cm
# ---- 样式常量结束 ----

def rgb(h): return RGBColor.from_string(h)

def set_font(run, size=None, bold=None, color=None):
    f = run.font
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    rPr.set('spc', '0')
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', FONT)
    if size: f.size = Pt(size)
    if bold is not None: f.bold = bold
    if color: f.color.rgb = rgb(color)

def parse(s, hot=RED):
    """'**x**' -> red bold ; '〖x〗' -> red bold (keyword)."""
    s = s.replace('〖', '**').replace('〗', '**')
    parts = re.split(r'(\*\*[^*]+\*\*)', s)
    return [(p[2:-2], True) if p.startswith('**') else (p, False) for p in parts if p]

SUBRE = re.compile(r'(CO₂|Q_[A-Za-z]+)')
def fill_par(p, text, size, bold=True, color=TEXT, hot=RED):
    for txt, h in parse(text):
        for piece in SUBRE.split(txt):
            if not piece: continue
            if piece == 'CO₂': parts = [('CO', 0), ('2', 1)]
            elif piece.startswith('Q_'): parts = [('Q', 0), (piece[2:], 1)]
            else: parts = [(piece, 0)]
            for t, sub in parts:
                r = p.add_run(); r.text = t
                set_font(r, size, bold, hot if h else color)
                if sub: r._r.get_or_add_rPr().set('baseline', '-25000')

def textbox(slide, x, y, w, h, lines, size, bold=True, color=TEXT, align=PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.TOP, bullet=False, space_after=4, name=None, line_spacing=None):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.1); tf.margin_top = tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = anchor
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        if line_spacing: p.line_spacing = line_spacing
        if bullet:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(Cm(0.6))); pPr.set('indent', str(-Cm(0.6)))
            bu = etree.SubElement(pPr, qn('a:buFont')); bu.set('typeface', 'Arial')
            bc = etree.SubElement(pPr, qn('a:buChar')); bc.set('char', '•')
        fill_par(p, line, size, bold, color)
    return tb

def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.5, rounded=True, name=None, radius=0.06, dash=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                 Cm(x), Cm(y), Cm(w), Cm(h))
    if name: shp.name = name
    if rounded: shp.adjustments[0] = radius
    if fill: shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    else: shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line); shp.line.width = Pt(line_w)
        if dash: shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    else: shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def shape_text(shp, lines, size, bold=True, color=TEXT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE):
    tf = shp.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.15); tf.margin_top = tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = anchor
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        fill_par(p, line, size, bold, color)

STYLE = {'panel': 'A', 'label': 'solid'}   # set per slide

def _line_alpha(shp, alpha=65000):
    ln = shp.line._get_or_add_ln()
    sf = ln.find(qn('a:solidFill'))
    if sf is not None and len(sf):
        clr = sf[0]
        for a in clr.findall(qn('a:alpha')): clr.remove(a)
        al = etree.SubElement(clr, qn('a:alpha')); al.set('val', str(alpha))

def _grad(shp, c1, c2, ang=5400000):
    spPr = shp._element.spPr
    for tag in ('a:solidFill', 'a:noFill', 'a:gradFill'):
        for e in spPr.findall(qn(tag)): spPr.remove(e)
    g = etree.Element(qn('a:gradFill')); g.set('rotWithShape', '1')
    gl = etree.SubElement(g, qn('a:gsLst'))
    for pos, c in ((0, c1), (100000, c2)):
        gs = etree.SubElement(gl, qn('a:gs')); gs.set('pos', str(pos))
        sc = etree.SubElement(gs, qn('a:srgbClr')); sc.set('val', c)
    lin = etree.SubElement(g, qn('a:lin')); lin.set('ang', str(ang)); lin.set('scaled', '0')
    spPr.find(qn('a:prstGeom')).addnext(g)

def label(slide, x, y, text, w=None, h=0.95, size=SZ_LABEL, style=None):
    style = style or STYLE['label']
    w = w or (0.7 + 0.62 * len(text) * size / 16)
    if style == 'soft':
        shp = rect(slide, x, y, w, h, fill=GREEN_L, line=GREEN, line_w=1.0, rounded=True, radius=0.25, name='标签')
        _line_alpha(shp, 60000)
        shape_text(shp, [text], size, True, GREEN_D)
    elif style == 'tab':
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUND_2_SAME_RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
        shp.adjustments[0] = 0.35
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(GREEN_D); shp.line.fill.background(); shp.shadow.inherit = False
        shp.name = '标签'
        shape_text(shp, [text], size, True, 'FFFFFF')
    else:
        shp = rect(slide, x, y, w, h, fill=GREEN, rounded=True, radius=0.25, name='标签')
        shape_text(shp, [text], size, True, 'FFFFFF')
    return shp

def panel(slide, x, y, w, h, lab=None, fill='FFFFFF', lab_w=None, line=GREEN, dash=False, style=None):
    style = style or STYLE['panel']
    if style in ('C', 'D'):     # C: top corners rounded; D: bottom corners rounded
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUND_2_SAME_RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
        if style == 'C': shp.adjustments[0] = 0.06; shp.adjustments[1] = 0.0
        else: shp.adjustments[0] = 0.0; shp.adjustments[1] = 0.06
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
        shp.line.color.rgb = rgb(line); shp.line.width = Pt(2.25); shp.shadow.inherit = False
        if dash: shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        _line_alpha(shp, 70000); shp.name = '面板'
    elif style == 'B':   # soft gradient fill
        shp = rect(slide, x, y, w, h, fill='FFFFFF', line=GREEN, line_w=1.75, radius=0.04, name='面板', dash=dash)
        _grad(shp, 'E3F1E8', 'FFFFFF'); _line_alpha(shp, 55000)
    else:                # A: white, thick translucent border
        shp = rect(slide, x, y, w, h, fill=fill, line=line, line_w=2.25, radius=0.04, name='面板', dash=dash)
        _line_alpha(shp, 70000)
    if lab:
        if STYLE['label'] == 'tab' and style in ('C', 'D'):
            label(slide, x, y, lab, w=lab_w)
        else:
            label(slide, x + 0.4, y + 0.25, lab, w=lab_w)
    return shp

def picture_fit(slide, path, x, y, w, h, name=None, mode='cover'):
    im = Image.open(path); iw, ih = im.size
    box = w / h; img = iw / ih
    if mode == 'contain':
        if img > box: nw, nh = w, w / img
        else: nw, nh = h * img, h
        pic = slide.shapes.add_picture(path, Cm(x + (w - nw) / 2), Cm(y + (h - nh) / 2), Cm(nw), Cm(nh))
    else:
        pic = slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))
        if img > box:
            c = (1 - box / img) / 2; pic.crop_left = pic.crop_right = c
        else:
            c = (1 - img / box) / 2; pic.crop_top = pic.crop_bottom = c
    if name: pic.name = name
    return pic

def placeholder(slide, x, y, w, h, text):
    shp = rect(slide, x, y, w, h, fill='F2F2F2', line='A6A6A6', line_w=1, rounded=False, dash=True, name='待放图片')
    shape_text(shp, ['待放：' + text], 12, False, GRAY)
    return shp

def image_or_placeholder(slide, path, x, y, w, h, text, mode='cover'):
    if path and os.path.exists(path):
        return picture_fit(slide, path, x, y, w, h, mode=mode)
    return placeholder(slide, x, y, w, h, text)

def line(slide, x1, y1, x2, y2, color=GREEN, width=1.25, name=None):
    ln = slide.shapes.add_connector(1, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    ln.line.color.rgb = rgb(color); ln.line.width = Pt(width)
    if name: ln.name = name
    return ln

def arrow(slide, x, y, w=0.8, h=0.7, color=GREEN):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(x), Cm(y), Cm(w), Cm(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(color); shp.line.fill.background()
    shp.shadow.inherit = False; shp.name = '箭头'
    return shp

def header(slide, num, title):
    tb = slide.shapes.add_textbox(Cm(MARGIN), Cm(0.3), Cm(24), Cm(1.55))
    tb.name = '页面标题'
    tf = tb.text_frame; tf.margin_left = tf.margin_right = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    if num:
        r = p.add_run(); r.text = num + '  '; set_font(r, SZ_TITLE, True, GREEN)
    r = p.add_run(); r.text = title; set_font(r, SZ_TITLE, True, TEXT)
    line(slide, MARGIN, LINE_Y, SLIDE_W - MARGIN, LINE_Y, GREEN, 1.5, name='页眉线')

def keylines(slide, lines, y=2.45):
    h = 1.05 * len(lines) + 0.1
    textbox(slide, MARGIN, y, CW, h, lines, SZ_KEY, True, TEXT, bullet=True, space_after=2, name='要点')
    return y + h

def source(slide, text):
    textbox(slide, MARGIN, 18.15, CW, 0.7, ['来源：' + text], SZ_SRC, False, GRAY, align=PP_ALIGN.CENTER, name='来源')

def conclusion(slide, text, y=17.15):
    textbox(slide, MARGIN, y, CW, 1.0, [text], SZ_CONC, True, RED, align=PP_ALIGN.CENTER, name='结论句')

def table(slide, x, y, w, rows, col_w=None, row_h=0.85, size=SZ_BODY, header_fill=GREEN):
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Cm(x), Cm(y), Cm(w), Cm(row_h * nr))
    tbl = gt.table
    tblPr = tbl._tbl.tblPr
    style = tblPr.find(qn('a:tableStyleId'))
    if style is not None: tblPr.remove(style)
    tblPr.set('firstRow', '0'); tblPr.set('bandRow', '0')
    if col_w:
        for i, cw in enumerate(col_w): tbl.columns[i].width = Cm(cw)
    for r in range(nr):
        tbl.rows[r].height = Cm(row_h)
        for c in range(nc):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(header_fill if r == 0 else ('FFFFFF' if r % 2 else 'F4F9F6'))
            cell.margin_left = cell.margin_right = Cm(0.15)
            cell.margin_top = cell.margin_bottom = Cm(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            fill_par(p, rows[r][c], size, True, 'FFFFFF' if r == 0 else TEXT)
    return gt

def remove(shape):
    shape._element.getparent().remove(shape._element)

def rich_box(slide, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, shape=None, name=None):
    """One text box (or text inside an existing shape) with several styled paragraphs.
    paras: list of (text, size, color, space_after)."""
    if shape is None:
        tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
        if name: tb.name = name
    else:
        tb = shape
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.2); tf.margin_top = tf.margin_bottom = Cm(0.1)
    tf.vertical_anchor = anchor
    for i, (text, size, color, sa) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(sa)
        fill_par(p, text, size, True, color)
    return tb


def est_lines(text, width_cm, size_pt):
    """Rough number of rendered lines for one paragraph (CJK = 1 em, ASCII ~ 0.55 em)."""
    import math, re as _re
    s = text.replace('**', '').replace('〖', '').replace('〗', '')
    em = sum(0.55 if ord(ch) < 128 else 1.0 for ch in s)
    per_line = max(1.0, (width_cm - 0.25) / (size_pt * 0.0353 * 1.02))
    return max(1, math.ceil(em / per_line))

def est_h(paras, width_cm, size_pt, line_spacing=1.15, space_after_pt=4, bullet=False):
    """Height (cm) needed by a list of paragraphs of one size."""
    w = width_cm - (0.6 if bullet else 0)
    n = sum(est_lines(p, w, size_pt) for p in paras)
    return n * size_pt * 0.0353 * 1.2 * line_spacing + len(paras) * space_after_pt * 0.0353 + 0.15
