# -*- coding: utf-8 -*-
"""11-19 低碳建筑专题研讨（Sugano 2026 ZEG 温室）v2：按 11-10 汇报版式重排。"""
import sys, os, io
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from pptlib import *

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = sys.argv[1]
FIG = os.path.join(ROOT, '素材', '论文图')
CROP = os.path.join(ROOT, '素材', '裁图')
KJ = os.path.join(ROOT, '素材', '课件图')
_figs = os.listdir(FIG)
def F(n): return os.path.join(FIG, [x for x in _figs if x.startswith('图%d_' % n)][0])
def C(n): return os.path.join(CROP, n + '.png')
BLUE = '1F5FA8'
BOT = 17.9
LS = 1.1

prs = Presentation(os.path.join(HERE, '母版.pptx'))
slides = list(prs.slides)
cover, toc, div1, sample, end = slides[:5]
layout = sample.slide_layout
logo_pic = [s for s in sample.shapes if s.name == '校徽'][0]
LOGO = io.BytesIO(logo_pic.image.blob)
LOGO_GEOM = (logo_pic.left, logo_pic.top, logo_pic.width, logo_pic.height)
STYLE['panel'], STYLE['label'] = 'A', 'solid'

def new_slide(num, title):
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    LOGO.seek(0)
    pic = s.shapes.add_picture(LOGO, *LOGO_GEOM); pic.name = '校徽'
    header(s, num, title)
    return s

def keys(s, lines, size=18):
    h = 0.92 * len(lines) + 0.15
    textbox(s, MARGIN, 2.35, CW, h, lines, size, True, TEXT, bullet=True, space_after=2, line_spacing=1.05, name='要点')
    return 2.35 + h + 0.35

def th(lines, w, size, sa=6, bullet=True):
    return est_h(lines, w, size, LS, sa, bullet)

def body(s, x, y, w, h, lines, size=16, bullet=True, sa=6, color=TEXT, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    return textbox(s, x, y, w, h, lines, size, True, color, align=align, anchor=anchor,
                   bullet=bullet, space_after=sa, line_spacing=LS, name='正文')

def cap(s, x, y, w, lines, size=12, align=PP_ALIGN.CENTER, color=TEXT, bold=False):
    h = est_h(lines, w, size, 1.05, 2)
    textbox(s, x, y, w, h, lines, size, bold, color, align=align, space_after=2, line_spacing=1.05, name='图注')
    return y + h

def pic(s, path, x, y, w, h, mode='contain'):
    return picture_fit(s, path, x, y, w, h, mode=mode)

def mark(s, p, rx, ry, rw, rh):
    L, T, W, H = p.left, p.top, p.width, p.height
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(L + rx * W), int(T + ry * H), int(rw * W), int(rh * H))
    shp.fill.background(); shp.line.color.rgb = rgb(RED); shp.line.width = Pt(2.0); shp.shadow.inherit = False
    shp.name = '图上标注'

def conn(s, x1, y1, x2, y2, color=GREEN, width=2.0):
    ln = line(s, x1, y1, x2, y2, color, width, name='箭头线')
    l = ln.line._get_or_add_ln()
    te = etree.SubElement(l, qn('a:tailEnd')); te.set('type', 'triangle'); te.set('w', 'med'); te.set('len', 'med')

def bar(s, y, text, size=17, h=None):
    h = h or (th([text], CW - 0.8, size, 0, False) + 0.35)
    r = rect(s, MARGIN, y, CW, h, fill=GREEN_L, radius=0.25, name='底框')
    shape_text(r, [text], size, True, TEXT)
    return y + h

def tbl(s, x, y, w, rows, col_w, row_h, size=14, merges=(), first_h=None):
    gt = table(s, x, y, w, rows, col_w=col_w, row_h=row_h, size=size)
    t = gt.table
    if first_h: t.rows[0].height = Cm(first_h)
    for (r, c1, c2) in merges:
        t.cell(r, c1).merge(t.cell(r, c2))
        cell = t.cell(r, c1); tf = cell.text_frame
        for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
        p = tf.paragraphs[0]
        for rr in list(p.runs): rr._r.getparent().remove(rr._r)
        p.alignment = PP_ALIGN.CENTER
        fill_par(p, rows[r][c1], size, True, TEXT)
    return gt

S = []
SRC = 'Sugano 等，Scientific Reports，2026'

# =============== 封面 ===============
def set_tb(shape, lines, size, color, bold=True):
    tf = shape.text_frame; tf.word_wrap = True
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    p = tf.paragraphs[0]
    for r in list(p.runs): r._r.getparent().remove(r._r)
    for i, t in enumerate(lines):
        pp = p if i == 0 else tf.add_paragraph()
        fill_par(pp, t, size, bold, color)
shp = {x.name: x for x in cover.shapes}
t3, t4 = shp['形状 3'], shp['形状 4']
t3.left, t3.top, t3.width, t3.height = Cm(2.2), Cm(5.0), Cm(17.6), Cm(3.4)
set_tb(t3, ['融合电气化供暖、围护改进与现场光伏的', '净零能耗温室全尺寸现场评估'], 26, GREEN)
t4.left, t4.top, t4.width, t4.height = Cm(2.2), Cm(8.6), Cm(17.6), Cm(2.2)
set_tb(t4, ['Full-scale field evaluation of a net zero energy greenhouse concept integrating electrified heating, envelope improvements, and on-site photovoltaics'], 12, '595959', bold=False)
tb = cover.shapes.add_textbox(Cm(2.2), Cm(10.75), Cm(17.6), Cm(1.6))
tf = tb.text_frame; tf.word_wrap = True
fill_par(tf.paragraphs[0], 'Soma Sugano, Masahisa Ishii, Ryota Tsuchiya, et al.', 12, False, '595959')
p = tf.add_paragraph(); fill_par(p, 'Scientific Reports, 2026. DOI: 10.1038/s41598-026-67257-8', 12, False, '595959')
tb = cover.shapes.add_textbox(Cm(2.2), Cm(4.0), Cm(17.6), Cm(0.9))
fill_par(tb.text_frame.paragraphs[0], '低碳农业建筑与通风 · 专题研讨（低碳建筑）', 14, True, GREEN_D)
set_tb(shp['日期'], ['2026 年 11 月 19 日'], 14, '595959', bold=False)

# =============== 文章主要内容 ===============
s = new_slide(None, '文章主要内容')
y0 = 2.75
ph_ = BOT - y0; pw_ = ph_ * 0.7727
pp = pic(s, os.path.join(ROOT, '素材', '论文首页.png'), MARGIN, y0, pw_, ph_)
pp.line.color.rgb = rgb('BFBFBF'); pp.line.width = Pt(0.75)
rx = MARGIN + pw_ + 1.2; rw = SLIDE_W - MARGIN - rx
textbox(s, rx, y0, rw, 0.8, ['Scientific Reports, 2026｜早稻田大学、日本农研机构（NARO）等'], 13, False, '595959')
H, T = 16, 19
paras = [('研究内容', H, GREEN_D, 3),
         ('日本三重县两栋相同的草莓温室并排运行，实测对比一个冬季', T, TEXT, 16),
         ('对照方案', H, GREEN_D, 3),
         ('常规温室：燃油热风机＋单层保温幕', T, TEXT, 3),
         ('ZEG 温室：地下水源热泵＋双层保温幕＋现场光伏', T, TEXT, 16),
         ('主要结果', H, GREEN_D, 3),
         ('供暖一次能源：减少〖58%〗', T, TEXT, 3),
         ('CO₂ 排放：减少〖68%〗', T, TEXT, 3),
         ('供暖费用：减少〖76%〗', T, TEXT, 3),
         ('建造费用：增加〖34%〗', T, TEXT, 0)]
rich_box(s, rx, y0 + 1.1, rw, BOT - y0 - 1.1, paras, anchor=MSO_ANCHOR.MIDDLE)
line(s, rx - 0.6, y0 + 1.2, rx - 0.6, BOT - 0.3, GREEN, 1.5)
source(s, SRC)
S.append(s)

# =============== 1.1 研究背景与目的 ===============
s = new_slide('1.1', '研究背景与关注的问题')
y0 = 2.6
pw_ = 9.6; ph_ = (BOT - 1.55 - 0.4 - y0 - 0.3 - 0.7) / 2
pic(s, C('9a'), MARGIN, y0, pw_, ph_, mode='cover')
pic(s, C('9c'), MARGIN, y0 + ph_ + 0.3, pw_, ph_, mode='cover')
cap(s, MARGIN, y0 + 2 * ph_ + 0.35, pw_, ['常规温室与燃油热风机（原文图 9）'], 12)
rx = MARGIN + pw_ + 0.8; rw = SLIDE_W - MARGIN - rx
a = ['日本约〖90%〗的加温温室使用燃油等燃烧式供暖，冬季供暖是温室能耗的主体，农户还要承受油价波动',
     '保温幕、热泵、光伏等单项技术都有研究，作者团队 2024 年也提出了零能耗温室（ZEG）的评价方法']
b = ['这些技术组合起来，在真实生产中到底能省多少能、减多少碳？',
     '能不能在不影响种植环境的前提下做到？',
     '要多花多少钱？']
c = ['已有研究多为模型和情景模拟，缺少同场地、同结构、同种植条件下与常规温室的并排实测，也很少报告建造费用']
ha = th(a, rw - 0.7, 16, 8) + 1.55
hb = th(b, rw - 0.7, 17, 6) + th(c, rw - 0.7, 15, 4) + 1.9
gap = (BOT - y0 - ha - hb - 1.55 - 0.4)
panel(s, rx, y0, rw, ha, '现状')
body(s, rx + 0.4, y0 + 1.4, rw - 0.7, ha - 1.45, a, 16, sa=8)
y2 = y0 + ha + gap
panel(s, rx, y2, rw, hb, '关注的问题', line='C55A11')
hc = th(c, rw - 0.7, 15, 4)
body(s, rx + 0.4, y2 + 1.35, rw - 0.7, hc, c, 15, bullet=False, sa=4, color='404040')
body(s, rx + 0.4, y2 + 1.45 + hc, rw - 0.7, hb - 1.5 - hc, b, 17, sa=6)
bar(s, BOT - 1.55, '研究目的：建两栋相同的温室并排实测，回答上面三个问题', 17, 1.55)
source(s, SRC + '，Introduction；日本农林水产省，2023')
S.append(s)

# =============== 1.2 零能耗建筑与 ZEG ===============
s = new_slide('1.2', '从零能耗建筑到零能耗温室')
y0 = keys(s, ['零能耗建筑先降负荷、再提效率、最后用可再生能源补足；作者把这一思路用于温室，提出 ZEG'])
nodes = [('1975', '丹麦建成“零能耗住宅”'), ('2006', 'NREL 比较零能耗定义'),
         ('2010', '欧盟提出近零能耗建筑'), ('2019', '我国发布 GB/T 51350'), ('2024', 'Sugano 等提出 ZEG/ZEF')]
ly = y0 + 0.25; seg = CW / len(nodes)
line(s, MARGIN + 0.4, ly, SLIDE_W - MARGIN - 0.4, ly, 'A6A6A6', 1.5)
for i, (yr, txt) in enumerate(nodes):
    cx = MARGIN + seg * (i + 0.5)
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(cx - 0.2), Cm(ly - 0.2), Cm(0.4), Cm(0.4))
    c.fill.solid(); c.fill.fore_color.rgb = rgb(GREEN_D); c.line.fill.background(); c.shadow.inherit = False
    textbox(s, cx - seg / 2, ly + 0.2, seg, 0.65, [yr], 14, True, GREEN_D, align=PP_ALIGN.CENTER)
    textbox(s, cx - seg / 2, ly + 0.8, seg, 0.65, [txt], 13, True, TEXT, align=PP_ALIGN.CENTER)
y1 = ly + 1.75
fh = BOT - y1 - 1.0; fw = fh * 2008 / 1312
fx = SLIDE_W - MARGIN - fw
f1 = pic(s, F(1), fx, y1, fw, fh)
mark(s, f1, 0.505, 0.035, 0.32, 0.13); mark(s, f1, 0.65, 0.775, 0.22, 0.14)
cap(s, fx, y1 + fh + 0.05, fw, ['图 1　D 输入能源　G 现场可再生能源　C 能源消耗　E 输出能源；ZEG：全年 G ≥ C'], 12)
lw = fx - MARGIN - 0.7
steps = [('用能前', '减少需求：围护保温、气密、被动设计'), ('用能中', '提高效率：如热泵'), ('能源类型', '可再生能源补足：如光伏')]
for i, (a, b) in enumerate(steps):
    yy = y1 + i * 1.5
    label(s, MARGIN, yy, a, w=3.0, h=1.15, size=15)
    textbox(s, MARGIN + 3.3, yy, lw - 3.3, 1.15, [b], 16, True, TEXT, anchor=MSO_ANCHOR.MIDDLE)
dy = y1 + 3 * 1.5 + 0.2
r = rect(s, MARGIN, dy, lw, BOT - dy, fill=GREEN_L, radius=0.06, name='定义框')
rich_box(s, 0, 0, 0, 0, [('本文对 ZEG 的定义', 17, GREEN_D, 6),
    ('在不影响作物产量和品质的前提下，以经济可行的方式维持室内环境；通过降低负荷、被动措施和高效设备显著节能，再配合可再生能源，使全年运行能耗与现场能源供给的收支大致为净零或正值', 17, TEXT, 0)],
    anchor=MSO_ANCHOR.MIDDLE, shape=r)
source(s, 'Esbensen & Korsgaard, Solar Energy, 1977；Torcellini 等, ACEEE, 2006；Directive 2010/31/EU；GB/T 51350—2019；Sugano 等, AIJ J. Technol. Des., 2024')
S.append(s)

# =============== 2.1 试验设计 ===============
s = new_slide('2.1', '试验设计')
y0 = keys(s, ['两栋温室结构、作物和控温程序相同，只在供暖、保温幕和光伏上不同'])
capl = ['图 11　平面布置（mm），左为常规温室，右为 ZEG 温室', '绿框 热泵　红框 燃油热风机　黄框 CO₂ 增施机　绿点 温湿度传感器　紫框 逆变器与蓄电池　深蓝条 光伏板']
fw = 17.2; fh = fw * 1379 / 2008
ch = est_h(capl, fw, 12, 1.05, 2)
fy = y0 + (BOT - y0 - fh - ch) / 2
pic(s, F(11), MARGIN, fy, fw, fh)
cap(s, MARGIN, fy + fh + 0.05, fw, capl, 12, PP_ALIGN.LEFT)
tx = MARGIN + fw + 0.7; tw = SLIDE_W - MARGIN - tx
rows = [['项目', '常规温室', 'ZEG 温室'],
        ['结构', '两栋相同：6 连栋，单跨 7.2 m，长 50 m，面积 2160 m²', ''],
        ['作物与控温', '两栋相同：高架草莓，夜间设定 6–7 °C，同一套控温程序', ''],
        ['供暖', '燃油热风机 2 台（各 116 kW），燃料为 A 重油', '地下水源热泵 2 台（各 32 kW）；燃油机 2 台作备用'],
        ['保温幕', '单层保温遮阳幕', '单层幕＋近红外反射膜幕'],
        ['光伏', '无', '26 块竖直光伏板（13.65 kW）＋5 kWh 蓄电池']]
rh = (BOT - y0 - 1.1) / 5
gt = tbl(s, tx, y0, tw, rows, [3.0, (tw - 3.0) / 2, (tw - 3.0) / 2], rh, 15, merges=[(1, 1, 2), (2, 1, 2)], first_h=1.1)
for r in (3, 4, 5):
    gt.table.cell(r, 2).fill.solid(); gt.table.cell(r, 2).fill.fore_color.rgb = rgb('E3F1E8')
source(s, SRC + '，Methods、图 11')
S.append(s)

# =============== 2.2 关键技术原理 ===============
s = new_slide('2.2', '关键技术原理')
y0 = keys(s, ['热泵与家用空调冬季制热原理相同，热源是 15–16 °C 的井水',
              '近红外反射膜作为第二层幕，冬季夜间展开，减少散热'])
ph = BOT - y0
ih = ph - 0.75; pw5 = ih * 1374 / 1145
pic(s, os.path.join(ROOT, '素材', '热泵原理图.png'), MARGIN, y0, pw5, ih)
cap(s, MARGIN, y0 + ih + 0.05, pw5, ['示意图，图中温度为水源热泵一般工况，不是本文实测值'], 12)
rx = MARGIN + pw5 + 0.6; rw = SLIDE_W - MARGIN - rx
bh = (ph - 0.5) / 2
blocks = [('地下水源热泵', '10d', ['耗 1 份电得 3–4 份热（COP）；燃油机烧 1 份油得热不到 1 份', '直接用灌溉井水，冬季不结霜，不打地埋孔']),
          ('近红外反射膜幕', '10b', ['反射阳光中带热量的近红外，透过可见光', '16:00 展开、次日 9:30 收起；夏季可减少太阳得热'])]
for i, (lab, img, lines_) in enumerate(blocks):
    yy = y0 + i * (bh + 0.5)
    panel(s, rx, yy, rw, bh, lab)
    iw = 6.4; ih = bh - 1.35 - 0.35
    pic(s, C(img), rx + 0.4, yy + 1.35, iw, ih, mode='cover')
    tw_ = rw - iw - 1.1
    hh = th(lines_, tw_, 15, 6)
    body(s, rx + iw + 0.75, yy + 1.35 + (ih - hh) / 2, tw_, hh, lines_, 15, sa=6)
source(s, SRC + '；U.S. DOE, Pump Up Your Savings with Heat Pumps, 2023')
S.append(s)

# =============== 2.3 测量方法与室内环境 ===============
s = new_slide('2.3', '测量方法与室内环境')
y0 = keys(s, ['2025 年 1 月 10 日至 4 月 10 日，每分钟记录温湿度、各设备电流和燃油流量',
              '两栋温室室温变化基本一致，全期温湿度差异在传感器误差范围内'])
capl = ['图 2　1 月 15—21 日气温。绿：ZEG 温室两测点；红：常规温室两测点；灰：室外']
fh = BOT - y0 - 0.75; fw = fh * 2008 / 1230
pic(s, F(2), MARGIN, y0, fw, fh)
cap(s, MARGIN, y0 + fh + 0.05, fw, capl, 12)
tx = MARGIN + fw + 0.6; tw = SLIDE_W - MARGIN - tx
rows = [['测量量', '精度'], ['气温', '±0.3 °C'], ['相对湿度', '±2.5%'], ['设备电流', '±(0.5 mV＋0.3%)'], ['燃油流量', '±1.0%']]
r1 = 1.25
textbox(s, tx, y0, tw, 0.7, ['测量精度（表 2）'], 14, True, GREEN_D)
tbl(s, tx, y0 + 0.75, tw, rows, [tw * 0.42, tw * 0.58], r1, 13)
y2 = y0 + 0.75 + 5 * r1 + 0.5
textbox(s, tx, y2, tw, 0.7, ['用量换算'], 14, True, GREEN_D)
rows = [['用量', '→ 一次能源', '→ CO₂'], ['电网用电 1 kWh', '2.71 kWh', '0.411 kg'], ['A 重油 1 L', '10.8 kWh', '2.75 kg']]
note = ['A 重油：日本农林渔业常用的轻质燃料油（JIS K 2205）']
nh = est_h(note, tw, 12, 1.05, 2)
r2 = (BOT - (y2 + 0.75) - nh - 0.1) / 3
tbl(s, tx, y2 + 0.75, tw, rows, [tw * 0.36, tw * 0.32, tw * 0.32], r2, 13)
cap(s, tx, BOT - nh, tw, note, 12, PP_ALIGN.LEFT)
source(s, SRC + '，图 2、表 2、Methods')
S.append(s)

# =============== 3.1 供暖运行与燃油消耗 ===============
s = new_slide('3.1', '燃油与用电')
y0 = keys(s, ['燃油减少〖90%〗：ZEG 温室夜间由两台热泵供热，290 L 燃油主要用在 2 月最冷时段',
              '总用电基本持平：热泵新增 3322 kWh，燃油机风机少用约 2960 kWh'])
fh6 = BOT - y0 - 0.8; fw6 = fh6 * 1832 / 1505
pic(s, F(6), MARGIN + 0.5, y0, fw6, fh6)
cap(s, MARGIN, y0 + fh6 + 0.05, fw6 + 1.0, ['图 6　监测期累计燃油消耗（L）'], 13)
rx = MARGIN + fw6 + 1.6; rw = SLIDE_W - MARGIN - rx
pv = ['三个月发电〖1846 kWh〗，全部被两栋温室自用', '约相当于 ZEG 温室用电的〖25%〗', '为避免向电网倒送电，部分时段被限发']
ph_ = th(pv, rw - 0.7, 16, 8) + 1.6
pih = BOT - y0 - ph_ - 0.4
pic(s, C('10e'), rx, y0, rw, pih, mode='cover')
panel(s, rx, y0 + pih + 0.4, rw, ph_, '现场光伏')
body(s, rx + 0.4, y0 + pih + 1.8, rw - 0.7, ph_ - 1.5, pv, 16, sa=8)
source(s, SRC + '，图 3–6、Results')
S.append(s)

# =============== 3.2 用电与光伏发电 ===============

# =============== 3.3 一次能源与碳排放 ===============
s = new_slide('3.2', '一次能源与碳排放')
y0 = keys(s, ['监测期（2025 年 1—4 月）供暖一次能源减少〖58%〗，CO₂ 减少〖68%〗'])
fh7 = BOT - y0 - 1.4; fw7 = fh7 * 2008 / 1356
pic(s, F(7), MARGIN, y0, fw7, fh7)
cap(s, MARGIN, y0 + fh7 + 0.05, fw7, ['图 7　单位面积供暖一次能源（a）与 CO₂ 排放（b）', '图例：Fuel oil A＝A 重油　Electricity (oil heater)＝燃油机风机用电　Electricity (HP)＝热泵用电　Photovoltaic＝光伏'], 12)
rx = MARGIN + fw7 + 0.8; rw = SLIDE_W - MARGIN - rx
paras = [('一次能源', 16, GREEN_D, 2), ('21.7 → 9.1 kWh/m²，〖−58%〗', 18, TEXT, 14),
         ('CO₂', 16, GREEN_D, 2), ('4.8 → 1.5 kg/m²，〖−68%〗', 18, TEXT, 14),
         ('光伏另抵扣', 16, GREEN_D, 2), ('一次能源 11%，CO₂ 7%', 18, TEXT, 14),
         ('CO₂ 降得比一次能源多，是因为燃油换成了电，热泵用较少的电提供同样的热', 14, TEXT, 8),
         ('一次能源＝A 重油热值＋电网用电 ×2.71', 12, '595959', 0)]
rich_box(s, rx, y0, rw, fh7, paras, anchor=MSO_ANCHOR.MIDDLE)
source(s, SRC + '，图 7、表 1')
S.append(s)

# =============== 3.4 运行费用与建造成本 ===============
s = new_slide('3.3', '经济性')
y0 = keys(s, ['供暖费用减少〖76%〗（2652.9 → 631.6 美元），建造费用增加〖34%〗（多出 20.27 万美元）'])
fh8 = BOT - y0 - 0.8; fw8 = fh8 * 1811 / 1546
pic(s, F(8), MARGIN + 0.5, y0, fw8, fh8)
cap(s, MARGIN, y0 + fh8 + 0.05, fw8 + 1.0, ['图 8　监测期供暖费用（美元）。紫：A 重油；浅灰：电费'], 12)
rx = MARGIN + fw8 + 1.6; rw = SLIDE_W - MARGIN - rx
rows = [['', '常规温室', 'ZEG 温室'], ['建造费用（美元）', '592,900', '795,600'], ['监测期供暖费用（美元）', '2652.9', '631.6']]
tbl(s, rx, y0 + 0.4, rw, rows, [rw * 0.44, rw * 0.28, rw * 0.28], 1.4, 14)
a = ['作者认为回收期较长', '监测只覆盖冬季三个月，全年光伏收益、维护费用和电力基本电费都没有计入']
ha = th(a, rw - 0.7, 16, 8) + 1.6
panel(s, rx, y0 + 5.2, rw, ha, '作者的评价')
body(s, rx + 0.4, y0 + 6.65, rw - 0.7, ha - 1.5, a, 16, sa=8)
cap(s, rx, y0 + 5.2 + ha + 0.4, rw, ['我们的估算：按三个月省约 2021 美元简单相除，收回多出的建造费用约需 100 个同样的三个月'], 14, PP_ALIGN.LEFT)
source(s, SRC + '，图 8、表 1、表 S1')
S.append(s)

# =============== 4.1 节能减排的原因 ===============
s = new_slide('4.1', '节能原因与推广障碍')
y0 = keys(s, ['作者认为减排主要来自双层幕和热泵；多项措施同时改变，无法定量拆分各自贡献'])
pw = (CW - 2 * 0.7) / 3
items = [('减少需求：双层幕', ['已有试验：加一层保温幕可减少约〖30%〗供暖需求', '本文双层幕气密性更好']),
         ('提高效率：热泵', ['季节 COP 大于 4，冬季不除霜', '燃油换成电，CO₂ 降幅大于一次能源降幅']),
         ('可再生能源：光伏', ['本次只测冬季三个月，且部分被限发', '全年运行、减少限发、配合储能，作用可能更大'])]
ph = max(th(t_, pw - 0.7, 16, 8) for _, t_ in items) + 1.7
for i, (lab, lines_) in enumerate(items):
    x = MARGIN + i * (pw + 0.7)
    panel(s, x, y0, pw, ph, lab, style='C')
    body(s, x + 0.4, y0 + 1.4, pw - 0.7, ph - 1.5, lines_, 16, sa=8)
    if i < 2: arrow(s, x + pw + 0.1, y0 + ph / 2 - 0.3, 0.5, 0.6)
y2 = y0 + ph + 0.6
textbox(s, MARGIN, y2, CW, 0.8, ['推广障碍'], 17, True, GREEN_D)
rows = [['障碍', '作者的说明'], ['投入门槛', '日本设施园艺以小农户为主，热泵和高性能保温材料前期投入高'],
        ['回收期', '建造费用高 34%，回收期较长'],
        ['碳价', '日本碳税和 J-Credit（碳信用）价格远低于欧洲，减排本身收益有限'],
        ['电网', '日本电网火电比例高；电网清洁化后，ZEG 温室减碳效果会更大']]
rh = (BOT - y2 - 0.9 - 0.9) / 4
gt = tbl(s, MARGIN, y2 + 0.85, CW, rows, [4.2, CW - 4.2], rh, 15, first_h=0.9)
for r in range(1, 5):
    c = gt.table.cell(r, 0); c.fill.solid(); c.fill.fore_color.rgb = rgb(GREEN_L)
    c1 = gt.table.cell(r, 1)
    for p_ in c1.text_frame.paragraphs: p_.alignment = PP_ALIGN.LEFT
source(s, SRC + '，Discussion；Kim 等，Energies，2018')
S.append(s)

# =============== 4.2 研究贡献与创新点 ===============
s = new_slide('4.2', '结论与创新点')
textbox(s, MARGIN, 2.5, CW, 0.8, ['结论'], 18, True, GREEN_D)
cl = ['1　草莓夜间设定 6–7 °C 的条件下，ZEG 温室与常规温室室内环境相当',
      '2　供暖一次能源减少〖58%〗，CO₂ 减少〖68%〗，燃油减少〖90%〗，供暖费用减少〖76%〗',
      '3　地下水源热泵和双层幕是主要减排来源，光伏另抵扣约 25% 一次能源',
      '4　建造费用高 34%，推广需要全年经济评价和政策支持']
ch = th(cl, CW - 0.8, 18, 10, False) + 0.4
r = rect(s, MARGIN, 3.35, CW, ch, fill='FFFFFF', line=GREEN, line_w=1.5, radius=0.05, name='结论框')
textbox(s, MARGIN + 0.4, 3.5, CW - 0.8, ch - 0.2, cl, 18, True, TEXT, space_after=10, line_spacing=LS)
y2 = 3.35 + ch + 0.6
textbox(s, MARGIN, y2, CW, 0.8, ['创新点'], 18, True, GREEN_D)
y3 = y2 + 0.85
pw = (CW - 2 * 0.6) / 3
items = [('全尺寸并排实测', '在日本首次于同一场地、商业化种植条件下，对 ZEG 温室和常规温室同时实测'),
         ('能耗与成本同时报告', '同一套实测数据给出能耗、CO₂、运行费用和实际建造费用'),
         ('评价框架用于温室', '把零能耗建筑按场地边界核算能量收支的方法用于温室，结果可作模型验证的基准')]
for i, (a_, b_) in enumerate(items):
    x = MARGIN + i * (pw + 0.6)
    rect(s, x, y3, pw, BOT - y3, fill=GREEN_L, radius=0.08, name='创新点框')
    rich_box(s, x + 0.3, y3 + 0.2, pw - 0.6, BOT - y3 - 0.4, [(a_, 19, GREEN_D, 10), (b_, 17, TEXT, 0)], anchor=MSO_ANCHOR.MIDDLE)
source(s, SRC + '，Abstract、Discussion')
S.append(s)

# =============== 4.3 研究不足 ===============
s = new_slide('4.3', '研究不足')
y0 = keys(s, ['结论适用于草莓、冬季、日本电网条件下的这一对温室'])
textbox(s, MARGIN, y0, CW, 0.8, ['作者提到的不足'], 17, True, GREEN_D)
la = [('时间', '只测了 1—4 月，未包括制冷，全年是否达到净零尚未验证'),
      ('作物与气候', '草莓夜温设定低、供暖负荷小，推广到其他作物需谨慎'),
      ('技术贡献', '多项技术同时改变，无法定量区分各自贡献'),
      ('经济评价', '未计全年光伏收益、电力基本电费、设备寿命与维护')]
gw = (CW - 0.6) / 2; gh = 2.9
for i, (a_, b_) in enumerate(la):
    x = MARGIN + (i % 2) * (gw + 0.6); yy = y0 + 0.9 + (i // 2) * (gh + 0.4)
    rect(s, x, yy, gw, gh, fill='FFFFFF', line=GREEN, line_w=1.25, radius=0.08, name='不足框')
    rich_box(s, x + 0.2, yy + 0.1, gw - 0.4, gh - 0.2, [(a_, 17, GREEN_D, 4), (b_, 17, TEXT, 0)], anchor=MSO_ANCHOR.MIDDLE)
y2 = y0 + 0.9 + 2 * gh + 0.4 + 0.6
lb = ['保温幕和热泵的作用分不开：只加一层幕、继续烧油，能省多少？',
      '只比较了温湿度，没有比较草莓产量和品质，“不影响作物”缺少直接证据']
hh = th(lb, CW - 0.8, 18, 10)
panel(s, MARGIN, y2, CW, hh + 1.6, '我们的疑问', line='C55A11')
body(s, MARGIN + 0.4, y2 + 1.4, CW - 0.8, hh, lb, 18, sa=10)
source(s, SRC + '，Discussion、Limitations')
S.append(s)

# =============== 讨论问题 ===============
s = new_slide(None, '讨论问题')
qy = 2.6; qh = 3.2
r = rect(s, MARGIN, qy, CW, qh, fill=GREEN_L, radius=0.12, name='问题框')
shape_text(r, ['把这套方案搬到华北的日光温室或牛舍，', '哪些做法可以沿用，哪些需要改，哪些行不通？'], 26, True, TEXT)
textbox(s, MARGIN, qy + qh + 0.2, CW, 0.8, ['我国正在推进“双碳”，设施农业的冬季供暖和夏季降温都面临减碳要求'], 15, False, '404040', align=PP_ALIGN.CENTER)
hy = qy + qh + 1.25
textbox(s, MARGIN, hy, CW, 0.8, ['思考方向：碳排放 ≈ 热需求 ÷ 转换效率 × 碳强度 − 可再生抵扣'], 17, True, GREEN_D)
hints = [('热需求', '日光温室、牛舍的围护和散热方式，与本文的连栋塑料温室有什么不同？'),
         ('转换效率', '华北的水源、空气温度和可用性如何？热泵还能达到这么高的 COP 吗？'),
         ('碳强度', '我国电网排放因子 0.53 kg/kWh，本文为 0.41'),
         ('可再生抵扣', '光伏装在哪里？能不能全年发电？')]
hw = (CW - 0.8) / 2; rowh = 1.75
for i, (k, t) in enumerate(hints):
    x = MARGIN + (i % 2) * (hw + 0.8); yy = hy + 0.95 + (i // 2) * rowh
    textbox(s, x, yy, 3.7, rowh, [k], 17, True, GREEN_D, anchor=MSO_ANCHOR.TOP)
    textbox(s, x + 3.7, yy, hw - 3.7, rowh, [t], 16, True, TEXT, anchor=MSO_ANCHOR.TOP, line_spacing=1.05)
by = hy + 0.95 + 2 * rowh + 0.15
line(s, MARGIN, by, SLIDE_W - MARGIN, by, 'BFBFBF', 1.0)
textbox(s, MARGIN, by + 0.2, CW, 0.8, ['零能耗的三步：用能前减少需求 → 用能中提高效率 → 能源类型换成可再生'], 15, True, TEXT)
fy = by + 1.1; fh = BOT - fy
fw = (CW - 0.8) / 2 - 2.4
textbox(s, MARGIN, fy, 2.4, fh, ['温室'], 15, True, GREEN_D, anchor=MSO_ANCHOR.MIDDLE)
pic(s, os.path.join(KJ, '温室热负荷公式.png'), MARGIN + 2.4, fy, fw, fh)
x2 = MARGIN + (CW + 0.8) / 2
textbox(s, x2, fy, 2.4, fh, ['畜禽舍'], 15, True, GREEN_D, anchor=MSO_ANCHOR.MIDDLE)
pic(s, os.path.join(KJ, '畜禽舍热负荷公式.png'), x2 + 2.4, fy, fw, fh)
source(s, '课件《节能减排》《绪论》；生态环境部、国家统计局，2023 年电力二氧化碳排放因子，2025')
S.append(s)

# =============== 讨论小结 ===============
s = new_slide(None, '讨论小结')
H, T = 20, 18
paras = [('热需求', H, GREEN_D, 3),
         ('日光温室：有后墙蓄热、夜间盖保温被，被动保温已经较强，可借鉴气密和多层保温幕', T, TEXT, 2),
         ('牛舍：要保证通风，不能靠封闭保温；冬季靠牛自身产热，主要能耗在夏季降温', T, TEXT, 14),
         ('转换效率', H, GREEN_D, 3),
         ('水源：华北地下水温可能更低，且地下水超采、取水受限；只取热不耗水的回灌式或土壤源热泵可以考虑', T, TEXT, 2),
         ('空气源：冬季低温下 COP 下降，还要除霜', T, TEXT, 14),
         ('碳强度', H, GREEN_D, 3),
         ('我国电网排放因子高于日本；用本文数据按 0.53 kg/kWh 重算，CO₂ 减少约 63%（本文 68%）', T, TEXT, 14),
         ('可再生抵扣', H, GREEN_D, 3),
         ('牛舍屋顶不需要采光，光伏可斜铺在屋顶上全年发电，比本文竖装在场地边发电更多', T, TEXT, 2),
         ('光伏板和屋面之间形成遮阳和通风间层，夏季还能减少屋面得热（课件：隔热构造）', T, TEXT, 16),
         ('三步的思路可以沿用，具体措施要按当地的建筑类型、热源和电网重新选', 19, RED, 0)]
rich_box(s, MARGIN + 0.2, 2.5, CW - 0.4, BOT - 2.5, paras, anchor=MSO_ANCHOR.MIDDLE)
source(s, '课件《节能减排》；Sugano 等，Scientific Reports，2026；生态环境部、国家统计局，2025')
S.append(s)

# =============== order ===============
sld = prs.slides._sldIdLst
ids = list(sld)
id_of = {x.slide_id: e for x, e in zip(prs.slides, ids)}
elems = [id_of[x.slide_id] for x in [cover] + S + [end]]
for e in ids: sld.remove(e)
for e in elems: sld.append(e)
prs.save(OUT)
print('saved', len(elems))
