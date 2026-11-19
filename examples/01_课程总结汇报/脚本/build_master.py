# -*- coding: utf-8 -*-
"""Build the adjusted CAU template-35 master deck + one sample content slide."""
import sys, copy, re, zipfile, shutil, os
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

SRC, OUT, IMGDIR = sys.argv[1], sys.argv[2], sys.argv[3]

# ---------------- style table (蒙古答辩 PPT 为基准，适当调整) ----------------
FONT = '微软雅黑'
GREEN, GREEN_D, GREEN_L = '00873B', '006B2F', 'EAF4EE'
TEXT, RED, CAP, GRAY = '222222', 'C00000', '595959', '7F7F7F'
SZ_TITLE, SZ_KEY, SZ_LABEL, SZ_BODY, SZ_CAP, SZ_SRC, SZ_CONC = 28, 18, 16, 14, 11, 10, 18
MARGIN = 0.9
SLIDE_W, SLIDE_H = 33.867, 19.05
LINE_Y = 2.05

def rgb(h): return RGBColor.from_string(h)

def set_font(run, size=None, bold=None, color=None):
    f = run.font
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', FONT)
    if size: f.size = Pt(size)
    if bold is not None: f.bold = bold
    if color: f.color.rgb = rgb(color)

def add_runs(par, segments, size, bold=True, color=TEXT):
    """segments: list of (text, is_number) ; numbers drawn red."""
    for txt, hot in segments:
        r = par.add_run(); r.text = txt
        set_font(r, size, bold, RED if hot else color)

def parse(s):
    """'**x**' marks red numbers."""
    parts = re.split(r'(\*\*[^*]+\*\*)', s)
    return [(p[2:-2], True) if p.startswith('**') else (p, False) for p in parts if p]

def textbox(slide, x, y, w, h, lines, size, bold=True, color=TEXT, align=PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.TOP, bullet=False, space_after=4, name=None):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Cm(0.1); tf.margin_top = tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = anchor
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        if bullet:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(Cm(0.6))); pPr.set('indent', str(-Cm(0.6)))
            bu = etree.SubElement(pPr, qn('a:buFont')); bu.set('typeface', 'Arial')
            bc = etree.SubElement(pPr, qn('a:buChar')); bc.set('char', '•')
        add_runs(p, parse(line), size, bold, color)
    return tb

def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.5, rounded=True, name=None, radius=0.06):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                 Cm(x), Cm(y), Cm(w), Cm(h))
    if name: shp.name = name
    if rounded: shp.adjustments[0] = radius
    if fill: shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    else: shp.fill.background()
    if line: shp.line.color.rgb = rgb(line); shp.line.width = Pt(line_w)
    else: shp.line.fill.background()
    shp.shadow.inherit = False
    shp.text_frame.text = ''
    return shp

def label(slide, x, y, text, w=None):
    w = w or (0.7 + 0.62 * len(text))
    shp = rect(slide, x, y, w, 0.95, fill=GREEN, rounded=True, radius=0.25, name='标签')
    tf = shp.text_frame; tf.margin_left = tf.margin_right = Cm(0.1)
    tf.margin_top = tf.margin_bottom = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text; set_font(r, SZ_LABEL, True, 'FFFFFF')
    return shp

def picture_fit(slide, path, x, y, w, h, name=None):
    """Insert picture cropped to fill the box (cover)."""
    im = Image.open(path); iw, ih = im.size
    box = w / h; img = iw / ih
    pic = slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))
    if img > box:   # too wide: crop left/right
        c = (1 - box / img) / 2; pic.crop_left = pic.crop_right = c
    else:
        c = (1 - img / box) / 2; pic.crop_top = pic.crop_bottom = c
    if name: pic.name = name
    return pic

def line(slide, x1, y1, x2, y2, color=GREEN, width=1.25, name=None):
    ln = slide.shapes.add_connector(1, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    ln.line.color.rgb = rgb(color); ln.line.width = Pt(width)
    if name: ln.name = name
    return ln

def remove(shape):
    shape._element.getparent().remove(shape._element)

def set_text_keep(shape, text, size=None, color=None, bold=None):
    """Replace all text in a shape with one run, keeping first paragraph props."""
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    for r in list(p0.runs)[1:]:
        r._r.getparent().remove(r._r)
    if p0.runs: r = p0.runs[0]
    else: r = p0.add_run()
    r.text = text
    set_font(r, size, bold, color)

def nospace(shape):
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r._r.get_or_add_rPr().set('spc', '0')

def all_text_shapes(container):
    for sh in container.shapes:
        if sh.shape_type == 6:
            yield from all_text_shapes(sh)
        elif sh.has_text_frame:
            yield sh

# ---------------- open template, keep slides 1,2,3,4,26 ----------------
prs = Presentation(SRC)
keep = [1, 2, 3, 4, 26]
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
for idx, sldId in enumerate(ids, 1):
    if idx not in keep:
        rId = sldId.get(qn('r:id'))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)
cover, toc, divider, content, end = list(prs.slides)

# rename shapes carrying the shop name
for s in prs.slides:
    for i, sh in enumerate(s.shapes):
        if '淘宝' in sh.name or '想搞' in sh.name:
            sh.name = '形状 %d' % (i + 1)

# ---------------- cover ----------------
for sh in list(cover.shapes):
    t = sh.text_frame.text if sh.has_text_frame else ''
    if t.startswith('HSVmode') or '勤奋严谨' in t:
        remove(sh)
for sh in all_text_shapes(cover):
    t = sh.text_frame.text
    if t.startswith('中国农业大学学校'):
        set_text_keep(sh, '农业生物环境与能源工程', 40, GREEN, True)
        sh.width = Cm(17.5); nospace(sh)
        p2 = sh.text_frame.add_paragraph(); r2 = p2.add_run(); r2.text = '研究进展'; set_font(r2, 40, True, GREEN); nospace(sh)
        sh.top = Cm(5.4); sh.height = Cm(4.2)
    elif t.startswith('题目保证两行'):
        set_text_keep(sh, '课程学习汇报', 28, GREEN_D, True); nospace(sh)
        sh.top = Cm(10.0)
    elif t.startswith('答辩人'):
        set_text_keep(sh, '汇报人：××', 16, 'FFFFFF', True)
    elif t.startswith('指导教师'):
        set_text_keep(sh, '××学院', 16, 'FFFFFF', True)
        sh.width = Cm(7.6)
    elif t.strip():
        pass
for sh in cover.shapes:   # move 汇报人 pill right of 学院 pill
    if sh.has_text_frame and sh.text_frame.text.startswith('汇报人'):
        sh.left = Cm(10.3)
# date line
textbox(cover, 2.2, 16.4, 12, 0.9, ['2026 年 11 月 10 日'], 14, False, CAP, name='日期')

# ---------------- TOC ----------------
toc_items = ['设施园艺生产与改造', '畜禽养殖智能化', '粪污资源化利用', '课程学习体会']
olds = ['课题背景', '研究内容', '结果讨论', '最后总结']
for sh in all_text_shapes(toc):
    t = sh.text_frame.text.strip()
    if t in olds:
        set_text_keep(sh, toc_items[olds.index(t)], 18, 'FFFFFF', True)
        bp = sh.text_frame._txBody.find(qn('a:bodyPr')); bp.set('anchor', 'ctr')
        sh.text_frame.margin_top = sh.text_frame.margin_bottom = 0
    elif t in ('1', '2', '3', '4'):
        set_text_keep(sh, t, 14, GREEN, True)

for g in toc.shapes:
    if g.shape_type == 6:
        kids = list(g.shapes)
        autos = [k for k in kids if k.shape_type == 1]
        if not autos or not any(k.shape_type == 17 for k in kids): continue
        pill = max(autos, key=lambda k: k.width)
        for k in kids:
            if k.shape_type == 17:
                k.top = pill.top; k.height = pill.height
                k._element.txBody.find(qn('a:bodyPr')).set('anchor', 'ctr')

# ---------------- divider ----------------
for sh in all_text_shapes(divider):
    t = sh.text_frame.text.strip()
    if t.startswith('Part'):
        set_text_keep(sh, '01', 54, GREEN, True)
    elif t == '课题背景':
        set_text_keep(sh, '设施园艺生产与改造', 36, GREEN, True)
        sh.width = Cm(16); nospace(sh)
    elif t.upper().startswith('RESEARCH'):
        set_text_keep(sh, 'Facility Horticulture and Mechanization', 12, GREEN, False)
        sh.width = Cm(14); nospace(sh)

# ---------------- end ----------------
for sh in all_text_shapes(end):
    t = sh.text_frame.text.strip()
    if t.startswith('感谢您的观看'):
        set_text_keep(sh, '敬请各位老师批评指正', 40, GREEN, True)
    elif t.upper().startswith('THANKS'):
        set_text_keep(sh, 'THANKS FOR YOUR ATTENTION', 12, GREEN, False)
    elif t.startswith('答辩人'):
        set_text_keep(sh, '汇报人：××', 16, 'FFFFFF', True)
    elif t.startswith('指导教师'):
        set_text_keep(sh, '××学院', 16, 'FFFFFF', True)
        sh.width = Cm(7.6)
for sh in end.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith('汇报人'):
        sh.left = Cm(10.3)

# ---------------- content header prototype (slide 4) ----------------
logo = None
for sh in list(content.shapes):
    if sh.shape_type == 13 and sh.left > Cm(20):
        logo = sh
    else:
        remove(sh)
# logo top-right, a bit smaller
ratio = logo.width / logo.height
logo.height = Cm(1.55); logo.width = int(Cm(1.55) * ratio)
logo.left = Cm(SLIDE_W - MARGIN) - logo.width; logo.top = Cm(0.25)
logo.name = '校徽'
HEADER_LOGO = copy.deepcopy(logo._element)

def header(slide, num, title):
    tb = slide.shapes.add_textbox(Cm(MARGIN), Cm(0.3), Cm(24), Cm(1.55))
    tb.name = '页面标题'
    tf = tb.text_frame; tf.margin_left = tf.margin_right = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = num + '  '; set_font(r, SZ_TITLE, True, GREEN)
    r = p.add_run(); r.text = title; set_font(r, SZ_TITLE, True, TEXT)
    line(slide, MARGIN, LINE_Y, SLIDE_W - MARGIN, LINE_Y, GREEN, 1.5, name='页眉线')

def source(slide, text):
    textbox(slide, MARGIN, 18.15, SLIDE_W - 2 * MARGIN, 0.7, [text], SZ_SRC, False, GRAY, name='来源')

# ---- sample content page: 1.3 日光温室生产存在的问题 ----
s = content
header(s, '1.3', '日光温室生产存在的问题')
textbox(s, MARGIN, 2.45, SLIDE_W - 2 * MARGIN, 1.9,
        ['装备水平低：设施蔬菜机械化率 **49.8%**，全国农作物为 **74.29%**',
         '依赖人工：生产以人工为主，50 岁以上工人占 **78.3%**'],
        SZ_KEY, True, TEXT, bullet=True, space_after=2, name='要点')
PW = (SLIDE_W - 2 * MARGIN - 0.8) / 2
PY, PH = 4.75, 10.6
for i, (lab, img, lines) in enumerate([
        ('装备水平低', 'H18.jpg', ['规划目标：设施种植机械化率 2025 年 **50%**，2030 年 **60%**',
                                   '中小拱棚、塑料大棚占 **70%** 以上，调控能力弱']),
        ('依赖人工', 'H09.jpg', ['耕整地、植保、采收、运输以人工为主'])]):
    x = MARGIN + i * (PW + 0.8)
    rect(s, x, PY, PW, PH, fill='FFFFFF', line=GREEN, line_w=1.5, radius=0.04, name='面板')
    picture_fit(s, os.path.join(IMGDIR, img), x + 0.5, PY + 1.35, PW - 1.0, 6.6, name='照片')
    label(s, x + 0.5, PY + 0.25, lab)
    textbox(s, x + 0.5, PY + 8.15, PW - 1.0, 2.3, lines, SZ_BODY, True, TEXT, bullet=False, space_after=3, name='说明')
# 原因 strip
rect(s, MARGIN, 15.65, SLIDE_W - 2 * MARGIN, 1.25, fill=GREEN_L, rounded=True, radius=0.2, name='原因底框')
tb = textbox(s, MARGIN + 0.4, 15.65, SLIDE_W - 2 * MARGIN - 0.8, 1.25,
             ['原因　农艺：南北向短垄｜农机：大田机具改小，与温室不匹配｜设施：未预留农机进出和转向空间'],
             SZ_BODY, True, TEXT, anchor=MSO_ANCHOR.MIDDLE, name='原因')
r0 = tb.text_frame.paragraphs[0].runs[0]
r0.text = '原因　'; set_font(r0, SZ_BODY, True, GREEN)
r1 = tb.text_frame.paragraphs[0].add_run(); r1.text = '农艺：南北向短垄｜农机：大田机具改小，与温室不匹配｜设施：未预留农机进出和转向空间'
set_font(r1, SZ_BODY, True, TEXT)
for extra in tb.text_frame.paragraphs[0].runs[1:-1]:
    extra._r.getparent().remove(extra._r)
source(s, '来源：《全国现代设施农业建设规划（2023—2030 年）》；李占台等，中国蔬菜，2019(8)')
s.notes_slide.notes_text_frame.text = '样张：用于检查页眉、字号、面板和图片版式。'

prs.save(OUT)

# ---------------- recolor template greens to CAU green ----------------
MAP = {}
for h in ['0AA046', '07A544', '09CA53', '0CBB53', '00B050', '0BA14A', '10A84F', '12A04A', '0DB14B']:
    MAP[h] = GREEN
for h in ['068436', '068838', '057C33', '088239', '099642', '087B36', '047A33', '0A8A3C']:
    MAP[h] = GREEN_D
MAP['0CFF69'] = '3FAE6E'   # gradient light end -> lighter CAU green
MAP['A5300F'] = GREEN      # theme accent1 (red-brown) -> CAU green
tmp = OUT + '.tmp'
with zipfile.ZipFile(OUT) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
    for item in zi.infolist():
        data = zi.read(item.filename)
        if item.filename.endswith('.xml') and ('slides/' in item.filename or 'slideLayouts/' in item.filename or 'slideMasters/' in item.filename or 'theme/' in item.filename):
            txt = data.decode('utf-8')
            for k, v in MAP.items():
                txt = re.sub('val="%s"' % k, 'val="%s"' % v, txt, flags=re.I)
            data = txt.encode('utf-8')
        zo.writestr(item, data)
os.replace(tmp, OUT)
print('saved', OUT)
