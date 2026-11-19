# -*- coding: utf-8 -*-
"""Build the full 20-page deck (v2 layout pass)."""
import sys, os, re, io
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptlib import *

MASTER, OUT, IMG, TEXTMD = sys.argv[1:5]
def P(*a): return os.path.join(IMG, *a)
Y = lambda n: P('Y_大牧人环控部件', n)
H = lambda n: P('检索H_下载', n)
LS = 1.15   # body line spacing

prs = Presentation(MASTER)
slides = list(prs.slides)
cover, toc, div1, sample, end = slides[:5]
divs = [div1] + slides[5:8]
layout = sample.slide_layout
logo_pic = [s for s in sample.shapes if s.name == '校徽'][0]
LOGO = io.BytesIO(logo_pic.image.blob)
LOGO_GEOM = (logo_pic.left, logo_pic.top, logo_pic.width, logo_pic.height)

def walk(shapes):
    for x in shapes:
        if x.shape_type == 6: yield from walk(x.shapes)
        else: yield x

def set_divider(s, num, title, en):
    for x in walk(s.shapes):
        if not x.has_text_frame or not x.text_frame.paragraphs[0].runs: continue
        t = x.text_frame.text.strip(); r = x.text_frame.paragraphs[0].runs
        if t == '01': r[0].text = num
        elif t == '设施园艺生产与改造': r[0].text = title
        elif t.startswith('Facility'): r[0].text = en
for s, v in zip(divs[1:], [('02', '畜禽养殖智能化', 'Intelligent Livestock Farming'),
                           ('03', '粪污资源化利用', 'Manure Resource Utilization'),
                           ('04', '课程学习体会', 'Reflections')]):
    set_divider(s, *v)
for d in divs:
    for x in walk(d.shapes):
        if x.has_text_frame and x.text_frame.text.strip()[:1].isascii() and len(x.text_frame.text.strip()) > 4:
            for pp in x.text_frame.paragraphs: pp.alignment = PP_ALIGN.LEFT

PAGE_STYLES = {'1.1': ('A', 'solid'), '1.2': ('C', 'tab'), '1.3': ('A', 'soft'), '1.4': ('B', 'soft'),
               '2.1': ('C', 'tab'), '2.2': ('B', 'solid'), '2.3': ('A', 'soft'), '2.4': ('C', 'tab'),
               '3.1': ('A', 'solid'), '3.2': ('B', 'soft'), '3.3': ('A', 'tab'), '3.4': ('C', 'tab')}
def new_slide(num, title):
    if num in PAGE_STYLES:
        STYLE['panel'], STYLE['label'] = PAGE_STYLES[num]
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    LOGO.seek(0)
    pic = s.shapes.add_picture(LOGO, *LOGO_GEOM); pic.name = '校徽'
    header(s, num, title)
    return s

def body(s, x, y, w, h, lines, size=SZ_BODY, align=PP_ALIGN.LEFT, bullet=False, sa=4, anchor=MSO_ANCHOR.TOP, color=TEXT):
    return textbox(s, x, y, w, h, lines, size, True, color, align=align, bullet=bullet,
                   space_after=sa, line_spacing=LS, anchor=anchor)

def keys(s, lines, w=None):
    w = w or CW
    textbox(s, MARGIN, 2.45, w, 1.15 * len(lines) + 0.15, lines, SZ_KEY, True, TEXT,
            bullet=True, space_after=3, line_spacing=1.1, name='要点')
    return 2.45 + 1.15 * len(lines) + 0.25

def fit(lines, w, size=SZ_BODY, sa=4, bullet=False):
    return est_h(lines, w, size, LS, sa, bullet)

K = lambda n: P(n)
J = lambda n: P('检索J_下载', n)

S = {}
# =============== 1.1 ===============
s = new_slide('1.1', '设施类型与产出水平'); S[4] = s
y0 = keys(s, ['我国设施种植 **4270 多万亩**，居世界首位，塑料棚室占 **69.4%**，日光温室占 **29%**',
              '环境控制程度越高，单位面积产量越高，对能源、设备和管理的投入要求也越高'])
ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(MARGIN), Cm(y0), Cm(CW), Cm(1.3))
ar.fill.solid(); ar.fill.fore_color.rgb = rgb(GREEN); ar.line.fill.background(); ar.shadow.inherit = False
ar.adjustments[0] = 0.62; ar.adjustments[1] = 0.45
shape_text(ar, ['环境控制程度提高：单产提高，投入增加'], 18, True, 'FFFFFF')
cw = (CW - 3 * 0.55) / 4; cy = y0 + 1.6; ch = 17.95 - cy
cards = [('日光温室', H('N07_沈阳农大日光温室外景_裁切.jpg'), ['依靠后墙蓄热，冬季不加温', '新疆北疆越冬栽培', '番茄 **45–50 kg/m²**']),
         ('连栋温室', P('H12.jpg'), ['荷兰玻璃温室', '加温、补光等综合调控', '番茄 **56.5 kg/m²**']),
         ('植物工厂', P('E01.jpg'), ['福建泉州，全人工光', '多层立体栽培，周年生产', '电费在成本中占比大']),
         ('空间站植物栽培', P('E04.jpg'), ['天宫水稻实验', '氧、水回收闭合度 **>95%**', '食物仍需地面补给'])]
for i, (lab, img, txt) in enumerate(cards):
    x = MARGIN + i * (cw + 0.55)
    panel(s, x, cy, cw, ch, lab)
    image_or_placeholder(s, img, x + 0.35, cy + 1.35, cw - 0.7, ch - 5.05, lab)
    body(s, x + 0.25, cy + ch - 3.55, cw - 0.5, 3.4, txt, align=PP_ALIGN.CENTER, sa=2)
source(s, '吴杨焕等，北方园艺，2023(2)；Montero et al., Wageningen UR, 2011；农业农村部发展规划司，2024')

# =============== 1.4 植物工厂 ===============

s = new_slide('1.4', '植物工厂的经营情况'); S[5] = s
y0 = keys(s, ['植物工厂单产高、可周年生产，但电力和建设投入大，经营压力明显',
              '日本人工光型植物工厂 **59%** 亏损，国内多数植物工厂同样面临盈利困难'])
lw = 21.3; ph = 15.95 - y0
panel(s, MARGIN, y0, lw, ph, '经营收支（日本设施园艺协会调查，2024 年度）', lab_w=13.6)
cx = MARGIN + 0.4
textbox(s, cx, y0 + 1.25, 11.6, 0.7, ['太阳光型温室'], 14, True, GREEN)
picture_fit(s, P('JP_图表88_太阳光型.png'), cx, y0 + 1.9, 11.6, 2.85, mode='contain')
textbox(s, cx, y0 + 4.8, 11.6, 0.7, ['人工光型植物工厂'], 14, True, GREEN)
picture_fit(s, P('JP_图表88_人工光型.png'), cx, y0 + 5.45, 11.6, 3.15, mode='contain')
px = cx + 12.0; pw_ = lw - 12.8
picture_fit(s, P('H01.jpeg'), px, y0 + 1.3, pw_, 3.35)
textbox(s, px, y0 + 4.65, pw_, 0.6, ['多层栽培架与 LED 补光（福建泉州）'], 12, True, CAP, align=PP_ALIGN.CENTER)
picture_fit(s, P('H02.jpeg'), px, y0 + 5.3, pw_, 3.0)
textbox(s, px, y0 + 8.3, pw_, 0.6, ['定植板上的生菜幼苗（同上）'], 12, True, CAP, align=PP_ALIGN.CENTER)
body(s, cx, y0 + 8.95, lw - 0.8, ph - 9.0,
     ['日本：电费占总成本 **24%**，照明、空调分别占电费的 **58%**、**31%**',
      '我国：植物工厂 **493 座**，能源约占成本 **60%**，生菜保本售价约 **15 元/kg**'], sa=2)
x2 = MARGIN + lw + 0.5; w2 = CW - lw - 0.5
label(s, x2, y0, '多多农研科技大赛')
picture_fit(s, P('D11.jpg'), x2, y0 + 1.2, w2, 4.8)
textbox(s, x2, y0 + 6.0, w2, 0.6, ['2024 年第四届评审现场'], 12, True, CAP, align=PP_ALIGN.CENTER)
table(s, x2, y0 + 6.65, w2, [['届次', '评价中增加的内容'], ['第三届', '能耗成本'],
      ['第四届', '建设成本（改造费上限 40 万元）'], ['第五届', '成本与能耗单列为一项']],
      col_w=[2.2, w2 - 2.2], row_h=(ph - 6.65) / 4, size=13)
conclusion(s, '体会：技术能否推广，要看增加的投入能否由增产和省工收回', y=16.3)
source(s, '日本施設園芸協会《大規模施設園芸・植物工場 実態調査》，2024、2025；科技日报，2025')
# =============== 1.2 日光温室问题 ===============

s = new_slide('1.2', '日光温室生产存在的问题'); S[6] = s
y0 = keys(s, ['装备水平低：设施蔬菜综合机械化率 **49.8%**，明显低于全国农作物耕种收综合机械化率 **74.29%**',
              '依赖人工：耕整地、植保、采收、运输以人工为主，劳动力老龄化明显'])
lw = 17.4; rw = CW - lw - 0.5
ltx = ['课上概括为：机不好用、无机可用、无好机用、用不好机',
       '规划目标：设施种植机械化率 2025 年 **50%**，2030 年 **60%**',
       '中小拱棚、塑料大棚占 **70%** 以上，调控能力弱']
rtx = ['北京 20 家园区：50 岁以上工人占 **78.3%**',
       '江西 138 个经营主体：41–60 岁占 **78.26%**',
       '江西调查中 30 岁以下仅占 **2.9%**']
th = max(fit(ltx, lw - 0.6, sa=2), fit(rtx, rw - 0.6, sa=2))
img_h = 6.6
ph = 1.35 + img_h + 0.2 + th + 0.25
panel(s, MARGIN, y0, lw, ph, '装备水平低')
iw = (lw - 0.8 - 0.3) / 2
picture_fit(s, P('H18.jpg'), MARGIN + 0.4, y0 + 1.35, iw, img_h)
picture_fit(s, P('H19.jpg'), MARGIN + 0.7 + iw, y0 + 1.35, iw, img_h)
body(s, MARGIN + 0.4, y0 + 1.55 + img_h, lw - 0.6, th, ltx, sa=2)
x2 = MARGIN + lw + 0.5
panel(s, x2, y0, rw, ph, '依赖人工')
iw2 = (rw - 0.8 - 0.3) / 2
picture_fit(s, P('H09.jpg'), x2 + 0.4, y0 + 1.35, iw2, img_h)
picture_fit(s, P('H08.jpg'), x2 + 0.7 + iw2, y0 + 1.35, iw2, img_h)
body(s, x2 + 0.4, y0 + 1.55 + img_h, rw - 0.6, th, rtx, sa=2)
ry = y0 + ph + 0.35
r = rect(s, MARGIN, ry, CW, 1.65, fill=GREEN_L, radius=0.25, name='原因底框')
shape_text(r, ['原因：农艺上多为南北向短垄，农机多由大田机具改小，温室建设时没有预留农机进出和转向空间'], 18, True, TEXT)
source(s, '《全国现代设施农业建设规划（2023—2030 年）》；李占台等，中国蔬菜，2019(8)；中国蔬菜，2024(4)')
# =============== 1.3 宜机化 ===============

s = new_slide('1.3', '宜机化改造与覆盖材料'); S[7] = s
y0 = keys(s, ['机具尺寸难以继续缩小，从栽培模式和温室结构两方面为农机作业创造条件',
              '采用宜机化栽培模式后，日光温室机械化水平由 **33%** 提高到 **70% 以上**'])
pw = (CW - 2 * 1.1) / 3
items = [('栽培模式', H('N02_北京日光温室移栽机.png'), 'cover', ['南北向短垄改为东西向长垄', '畦中心距 **180–200 cm**，宽沟窄畦']),
         ('温室结构', H('N05_Stigas管轨作业车.jpg'), 'cover', ['跨度按垄数确定，山墙留农机进出口', '荷兰温室建设时即为作业车铺设轨道']),
         ('覆盖材料', P('H16.jpg'), 'cover', ['北墙前光量：普通膜为南侧的 **13%**，散射膜为 **23%–34%**', '东西向、南北向番茄产量差异不显著'])]
th = max(fit(t_, pw - 0.5, sa=2) for _, _, _, t_ in items)
img_h = 5.9
ph = 1.35 + img_h + 0.2 + th + 0.2
for i, (lab, img, mode, txt) in enumerate(items):
    x = MARGIN + i * (pw + 1.1)
    panel(s, x, y0, pw, ph, lab)
    picture_fit(s, img, x + 0.35, y0 + 1.35, pw - 0.7, img_h, mode=mode)
    body(s, x + 0.25, y0 + 1.55 + img_h, pw - 0.5, th, txt, align=PP_ALIGN.CENTER, sa=2)
    if i < 2: arrow(s, x + pw + 0.15, y0 + ph / 2 - 0.4, 0.8, 0.8)
r = rect(s, MARGIN, y0 + ph + 0.35, CW, 1.55, fill=GREEN_L, radius=0.25)
shape_text(r, ['东西向长垄会出现南行遮挡北行的问题，散射膜使北侧光照更均匀；性能好的覆盖材料价格也更高'], 18, True, TEXT)
source(s, '××老师，李明. 农业工程学报，2020；NY/T 1086.6—2016；××老师等，农业工程学报，2017')
# =============== 2.1 ===============

s = new_slide('2.1', '畜禽养殖的规模与智能化进展'); S[9] = s
y0 = keys(s, ['2025 年全国生猪出栏 **71973 万头**，生猪规模化率超过 **70%**，肉鸡超过 **88%**',
              '规模场装备基础较好，智能感知从研究走向生产应用仍有距离'])
ph = 17.95 - y0
rect(s, MARGIN, y0, 12.6, ph, fill='FFFFFF', line=GREEN, line_w=1.25, radius=0.04, dash=True, name='面板')
label(s, MARGIN + 0.4, y0 + 0.25, '规模化养殖')
picture_fit(s, P('X01_京鹏大型牛舍航拍.jpeg'), MARGIN + 0.35, y0 + 1.35, 11.9, ph - 1.7)
x2 = MARGIN + 13.1; w2 = CW - 13.1; hh = (ph - 0.4) / 2
for j, (lab, txt, img, st) in enumerate([
        ('设施装备较为成熟', ['规模场关键环节已基本机械化，全国奶牛机械挤奶率 **95%**（2019）',
                              '天津天食智慧牧场：项圈、推料机器人、AI 影像喷淋，繁殖率提高 **3%**'],
         P('X05_北京工业大学巡检推料机器人_展会实拍.jpg'), 'C'),
        ('智能感知仍在探索', ['设备和算法研究持续进展，成熟的商业化方案仍然不足',
                              '规划目标：奶牛智慧牧场 2025 年 **100 个**、2030 年 **200 个**'],
         P('X03_项圈监测界面_展会实拍.jpeg'), 'D')]):
    y = y0 + j * (hh + 0.4)
    panel(s, x2, y, w2, hh, lab, style=st)
    body(s, x2 + 0.4, y + 1.35, w2 - 7.9, hh - 1.5, txt, sa=8, anchor=MSO_ANCHOR.MIDDLE)
    picture_fit(s, img, x2 + w2 - 7.2, y + 0.35, 6.85, hh - 0.7)
source(s, '《全国现代设施农业建设规划（2023—2030 年）》；国家统计局，2025 年统计公报；农业农村部农机化司，2019；ECPLF 2026 会议结论')
# =============== 2.2 ===============
s = new_slide('2.2', '感知的对象与手段'); S[10] = s
lw = 19.6
y0 = keys(s, ['舍内环境：传感器采集数据，环控器联动调控设备',
              '动物个体：项圈、摄像头、巡检机器人、拾音器'], w=lw)
label(s, MARGIN, y0, '环境感知与调控（以大牧人猪舍环控系统为例）', w=13.6)
def tree(root, items, ytop, img_sz):
    n = len(items); iw_ = lw / n
    rx = MARGIN + lw / 2 - 2.0
    rb = rect(s, rx, ytop, 4.0, 0.95, fill=GREEN, radius=0.3)
    shape_text(rb, [root], 16, True, 'FFFFFF')
    ybar = ytop + 1.35
    line(s, MARGIN + lw / 2, ytop + 0.95, MARGIN + lw / 2, ybar, GREEN, 1.25)
    line(s, MARGIN + iw_ / 2, ybar, MARGIN + lw - iw_ / 2, ybar, GREEN, 1.25)
    for k, (nm, f) in enumerate(items):
        cx_ = MARGIN + iw_ * k + iw_ / 2
        line(s, cx_, ybar, cx_, ybar + 0.35, GREEN, 1.25)
        picture_fit(s, Y(f), cx_ - img_sz / 2, ybar + 0.45, img_sz, img_sz, mode='contain')
        textbox(s, cx_ - iw_ / 2, ybar + 0.5 + img_sz, iw_, 0.75, [nm], 14, True, TEXT, align=PP_ALIGN.CENTER)
    return ybar + 0.5 + img_sz + 0.75
yb = tree('感知监测', [('温度', 'Y02_温度传感器.png'), ('湿度', 'Y03_湿度传感器.png'), ('CO₂', 'Y04_CO2传感器.png'),
                       ('氨气', 'Y05_氨气传感器.png'), ('压差', 'Y06_压差传感器.png')], y0 + 1.25, 2.1)
cy_ = yb + 0.15
picture_fit(s, Y('Y01_环境控制器BH8800.png'), MARGIN + 1.0, cy_, 2.6, 2.4, mode='contain')
body(s, MARGIN + 3.8, cy_ + 0.15, lw - 4.0, 2.2,
     ['环境控制器：接入温度、湿度、CO₂、氨气等传感器，按设定曲线联动风机、湿帘、加热等 12 路设备'], size=14, sa=0,
     anchor=MSO_ANCHOR.MIDDLE)
tree('调控设备', [('风机', 'Y07_风机.png'), ('进风窗', 'Y08_吊顶进风窗.png'), ('湿帘', 'Y09_湿帘.png'),
                  ('加热', 'Y10_加热器.png'), ('灯光', 'Y11_灯光.png'), ('喷淋', 'Y12_喷淋.png')], cy_ + 2.65, 1.8)
x2 = MARGIN + lw + 0.5; w2 = CW - lw - 0.5; py2 = 2.45; ph = 17.95 - py2
panel(s, x2, py2, w2, ph, '个体感知')
rows = [(P('X02_项圈传感器.png'), '项圈', '采食、反刍、躺卧、跛行；热应激和行为变化放在同一时间线上'),
        (P('X07_国科农牧牛舍摄像头_展会实拍.jpg'), '摄像头', '体重、体况、行为、粪便形态；群养生猪俯拍估重，不改猪舍'),
        (P('H21.png'), '巡检机器人', '死鸡、寡产鸡巡检；十万羽鸡舍运行 **300 天**，检出寡产鸡 **1920 只**'),
        (J('j5_soundtalks.jpg'), '拾音器', '夜间咳嗽、呼噜声识别，用于健康预警')]
rh = (ph - 1.45) / 4
for k, (img, n, d) in enumerate(rows):
    y = py2 + 1.3 + k * rh
    image_or_placeholder(s, img, x2 + 0.35, y + 0.1, 4.2, rh - 0.3, '拾音器图标')
    rich_box(s, x2 + 4.7, y, w2 - 4.9, rh, [(n, 16, GREEN, 2), (d, 14, TEXT, 0)], anchor=MSO_ANCHOR.MIDDLE)

# =============== 2.3 ===============
s = new_slide('2.3', '生产现场数据获取的难点与应对'); S[11] = s
y0 = keys(s, ['图像：光照变化大、遮挡多，同一台相机一天内画面差别很大',
              '验证：多数系统在实验室或小群体中搭建，缺少生产验证'])
iw = (CW - 3 * 0.45) / 4
for k, (f, n) in enumerate([('Z01_牛舍白天正常光照.jpg', '白天正常'), ('Z02_牛舍阳光条纹.jpg', '阳光条纹'),
                            ('Z03_牛舍逆光眩光.jpg', '逆光眩光'), ('Z04_牛舍夜间.jpg', '夜间低照度')]):
    x = MARGIN + k * (iw + 0.45)
    picture_fit(s, P(f), x, y0, iw, 4.7)
    textbox(s, x, y0 + 4.72, iw, 0.75, [n], SZ_BODY, True, TEXT, align=PP_ALIGN.CENTER)
by = y0 + 5.6; bh = 17.95 - by; pw = (CW - 0.6) / 2
panel(s, MARGIN, by, pw, bh, '2026 欧洲精准畜牧业大会结论', lab_w=10.0)
body(s, MARGIN + 0.4, by + 1.3, pw - 0.8, bh - 1.4, ['技术开发与验证多依赖小规模数据集，成熟、完整的商业化方案仍然不足',
     '动物在商业生产中的“正常”行为缺少统一定义，难以作为识别异常的基准',
     '测量准确性和投资回报仍是推广的主要障碍，需要在设备成本和精度之间权衡'], size=15, bullet=True, sa=5)
x2 = MARGIN + pw + 0.6
panel(s, x2, by, pw, bh, '课上的应对思路', dash=True)
body(s, x2 + 0.4, by + 1.3, pw - 0.8, bh - 1.4, ['群养生猪估重：群养条件下难以获取标准站姿，研究发现侧卧与站立姿态的估重效果无显著差异，两种姿态融合后误差为 **3.83%**',
     '异常鸡粪识别：异常样本在生产中很少，用生成模型合成样本扩充数据集，再训练检测模型',
     '两例的共同点：现场条件难以改变，就让方法去适应现场'], size=15, bullet=True, sa=5)
source(s, 'ECPLF 2026 会议结论')

# =============== 2.4 ===============

s = new_slide('2.4', '智能装备与养殖场的相互适应'); S[12] = s
y0 = keys(s, ['智能装备进入养殖场后，装备和场舍要相互适应，在使用中逐代改进',
              '挤奶机器人从原型到大规模应用经历了 30 多年，日本蛋鸡场的巡检机器人也已到第 6 代'])
lw = 17.6; rw = CW - lw - 0.5; ph = 16.1 - y0
panel(s, MARGIN, y0, lw, ph, '挤奶机器人（Lely Astronaut）', lab_w=9.5)
iw = (lw - 0.8 - 0.3) / 2
picture_fit(s, H('N04_Lely挤奶工作照.jpg'), MARGIN + 0.4, y0 + 1.35, iw, 5.6)
picture_fit(s, P('H23.JPG'), MARGIN + 0.7 + iw, y0 + 1.35, iw, 5.6)
body(s, MARGIN + 0.4, y0 + 7.2, lw - 0.8, ph - 7.3,
     ['1992 年出现原型，1995 年商用，此后多代改进',
      '2024 年累计交付第 **5 万台**，用于 50 多个国家',
      '每天为 **250 万头**以上奶牛挤奶'], bullet=True, sa=6)
x2 = MARGIN + lw + 0.5
panel(s, x2, y0, rw, ph, '日本蛋鸡场导入巡检机器人', lab_w=8.6)
picture_fit(s, H('N03_木鸡郎6光阳蛋业.jpg'), x2 + 0.4, y0 + 1.35, rw - 0.8, 5.6)
body(s, x2 + 0.4, y0 + 7.2, rw - 0.8, ph - 7.3,
     ['设计：首家厂商临近完工才提出空间要求，计划作废',
      '使用：死鸡巡检由每天 **8 工时**降到 **1 工时**',
      '改造：鸡舍拟改为 6 层 × 2 层楼，机器人由 **6 台**减到 **2 台**'], size=15, bullet=True, sa=5)
conclusion(s, '体会：智能装备的成熟依赖现场反馈与持续迭代', y=16.35)
# =============== 3.1 ===============
s = new_slide('3.1', '粪污利用现状与处理目标的变化'); S[14] = s
y0 = keys(s, ['我国畜禽粪污年产生量约 **30.5 亿吨**（2019），综合利用率已达 **80.1%**',
              '利用方式以肥料化为主：多数粪污经无害化处理后就地还田，商品有机肥还田比例不足 **10%**'])
L = lambda n: P('L_废水课件', n)
bw = (CW - 2 * 1.1) / 3
label(s, MARGIN, y0, '污水处理目标的演变')
ch = 6.75; cy = y0 + 1.1
for k, (a, b, img) in enumerate([
        ('卫生排放（公元前 800 年—1800 年）', '管网排污，防止疾病与恶臭', L('L01_古罗马下水道.png')),
        ('达标处理（1893 年起）', '生物滤池、活性污泥法，应对有机负荷增长和水体富营养化', J('j6_aeration.jpg')),
        ('资源回收（1970 年代起）', '再生回用，回收沼气、营养物等资源', L('L03_污水厂沼气回收.png'))]):
    x = MARGIN + k * (bw + 1.1)
    panel(s, x, cy, bw, ch)
    picture_fit(s, img, x + 0.3, cy + 0.3, bw - 0.6, 3.75)
    rich_box(s, x + 0.15, cy + 4.15, bw - 0.3, ch - 4.25, [(a, 16, GREEN, 2), (b, 15, TEXT, 0)], align=PP_ALIGN.CENTER)
    if k < 2: arrow(s, x + bw + 0.15, cy + 1.8, 0.8, 0.8)
y1 = cy + ch + 0.25
label(s, MARGIN, y1, '畜禽粪污处理目标也在发生类似转变', style='soft', w=11.5)
row2 = [('污染防治', '2014 年条例：不得未经处理直接排放，鼓励综合利用', P('X06_固液分离机_展会实拍.jpg')),
        ('资源化利用', '2017 年起以肥料化为基础，利用率目标 **75%** → **85%**', H('M05_槽式堆肥车间.jpg')),
        ('能源化探索', '沼气、生物天然气，以及水热液化制取液体燃料', H('M02_君乐宝万头牧场沼气.jpg'))]
ry = y1 + 1.1; rh = 16.2 - ry
for k, (a, b, img) in enumerate(row2):
    x = MARGIN + k * (bw + 1.1)
    rect(s, x, ry, bw, rh, fill=GREEN_L, radius=0.08)
    picture_fit(s, img, x + 0.2, ry + 0.2, 3.3, rh - 0.4)
    rich_box(s, x + 3.6, ry + 0.05, bw - 3.7, rh - 0.1, [(a, 15, GREEN, 1), (b, 13, TEXT, 0)], anchor=MSO_ANCHOR.MIDDLE)
    if k < 2: arrow(s, x + bw + 0.15, ry + rh / 2 - 0.35, 0.8, 0.7)
conclusion(s, '粪污处理的重心正转向资源化、能源化利用', y=16.45)
source(s, '《畜禽规模养殖污染防治条例》（国务院令第 643 号）；国办发〔2017〕48 号；农规发〔2024〕27 号')

# =============== 3.2 ===============
s = new_slide('3.2', '粪污处理利用中的问题'); S[15] = s
y0 = keys(s, ['种养脱节：种植与养殖结合不紧密，粪污缺少高效利用的途径',
              '成本与运行：工艺选择过于复杂、处理成本高，寒区冬季沼气工程难以正常运行'])
tw = 11.6; lw = CW - tw - 0.6; ph = 16.1 - y0
panel(s, MARGIN, y0, lw, ph, '粪污处理设备')
iw = (lw - 0.8 - 0.35) / 2
picture_fit(s, P('X06_固液分离机_展会实拍.jpg'), MARGIN + 0.4, y0 + 1.35, iw, ph - 4.25)
picture_fit(s, H('M03_Farmnote再生垫料.jpg'), MARGIN + 0.75 + iw, y0 + 1.35, iw, ph - 4.25)
textbox(s, MARGIN + 0.4, y0 + ph - 2.85, iw, 0.75, ['固液分离机（奶业大会展会）'], 14, True, CAP, align=PP_ALIGN.CENTER)
textbox(s, MARGIN + 0.75 + iw, y0 + ph - 2.85, iw, 0.75, ['再生牛粪垫料（日本北海道牧场）'], 14, True, CAP, align=PP_ALIGN.CENTER)
body(s, MARGIN + 0.4, y0 + ph - 2.05, lw - 0.8, 1.9, ['固液分离后，固体做牛床垫料或有机肥，粪水进入厌氧发酵'], sa=0)
x2 = MARGIN + lw + 0.6
textbox(s, x2, y0, tw, 1.0, ['盈利案例：嘉立荷万头牛场沼气发电'], 16, True, TEXT, anchor=MSO_ANCHOR.MIDDLE)
table(s, x2, y0 + 1.2, tw, [['项目', '数值'], ['总投资', '约 **2000 万元**'], ['年运行费用', '**160 万元**'],
      ['年收益', '**766.7 万元**'], ['日产沼气', '**16000 m³**'], ['沼气用途', '发电自用、奶厅烧热水、食堂做饭'],
      ['沼渣、沼液', '牛床垫料、有机肥、青贮玉米']], col_w=[3.3, tw - 3.3], row_h=(ph - 1.2) / 7, size=15)
conclusion(s, '处理工艺力求简单实用，前提是源头减量', y=16.45)

# =============== 3.3 ===============
s = new_slide('3.3', '水热液化：粪便制取液体燃料'); S[16] = s
y0 = keys(s, ['粪便含水率高，水热液化无需干燥，可将全组分转化为生物原油',
              '生物原油含氧 **23.5%**、酸值高，需要提质才能替代交通燃料'])
picture_fit(s, J('j1_htl_aarhus.jpg'), MARGIN, y0, 8.6, 5.0)
textbox(s, MARGIN, y0 + 5.0, 8.6, 0.55, ['丹麦奥胡斯大学连续式 HTL 中试装置'], 11, True, CAP, align=PP_ALIGN.CENTER)
fx = MARGIN + 9.1; fw = CW - 9.1
label(s, fx, y0, '技术路线', w=3.6)
steps = [('畜禽粪便', '含水率高'), ('水热液化', '高温高压水'), ('生物原油', '含氧高'), ('提质', '脱氧、分馏'), ('柴油航煤', '交通燃料')]
sw = (fw + 4 * 0.45) / 5
for k, (st, sub) in enumerate(steps):
    x = fx + k * (sw - 0.45)
    shp = s.shapes.add_shape(MSO_SHAPE.PENTAGON if k == 0 else MSO_SHAPE.CHEVRON, Cm(x), Cm(y0 + 1.3), Cm(sw), Cm(3.7))
    shp.adjustments[0] = 0.2
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(GREEN if k % 2 == 0 else GREEN_D)
    shp.line.color.rgb = rgb('FFFFFF'); shp.line.width = Pt(2); shp.shadow.inherit = False
    rich_box(s, 0, 0, 0, 0, [(st, 15, 'FFFFFF', 2), (sub, 11, 'EAF4EE', 0)], shape=shp, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    shp.text_frame.margin_left = Cm(0.55 if k else 0.1); shp.text_frame.margin_right = Cm(0.25)
py = y0 + 5.85; pw = (CW - 2 * 0.5) / 3; ph = 15.85 - py
for k, (yr, tx) in enumerate([('2018', ['分馏耦合酯化制柴油：掺混 **10%** 时内燃机动力输出达对照的 **96%–100%**', '提质燃油热值与商用柴油相当']),
                              ('2025', ['无氢炼制：比加氢工艺温室气体排放降低 **28%**、能耗减少 **35%**', '绿煤与生物润滑油联产，约 **90%** 组分高值利用']),
                              ('2026', ['催化制航煤：氧去除率 **98%** 以上，航煤馏分占提质油 **45%**', '平均碳数 **10.5**，接近航煤的 **11.4**'])]):
    x = MARGIN + k * (pw + 0.5)
    panel(s, x, py, pw, ph, yr, dash=True)
    body(s, x + 0.35, py + 1.3, pw - 0.6, ph - 1.4, tx, size=16, sa=6)
textbox(s, MARGIN, 15.95, CW, 2.1, ['水热液化（HTL）：在高温高压水中把湿生物质转化为油状液体的工艺，原料不用烘干',
        '生物原油：水热液化得到的粗油，含氧多、酸性强，要提质后才能当燃料用',
        '亚/超临界水：温度、压力接近或超过水的临界点（374 ℃、22 MPa）的水'], 12, False, TEXT, space_after=0)
source(s, 'Nature Sustainability, 2018；Nature Communications, 2025；Nature Sustainability, 2026（课程课件）')

# =============== 3.4 ===============
s = new_slide('3.4', '碳减排与可持续航空燃料需求'); S[17] = s
y0 = keys(s, ['2021 年我国农业温室气体排放 **9.31 亿吨** CO₂ 当量，其中畜禽粪便管理 **1.65 亿吨**',
              '中长途航空难以电气化，SAF 预计承担 2050 年约 **65%** 的减排量'])
pw = (CW - 0.6) / 2; ph = 14.35 - y0
label(s, MARGIN, y0, '欧盟 SAF 最低掺混比例')
picture_fit(s, J('j7_easa_saf_chart.png'), MARGIN, y0 + 1.2, pw, ph - 1.9, mode='contain')
textbox(s, MARGIN, y0 + ph - 0.65, pw, 0.6, ['红色为 SAF 总比例，橙色为其中的合成燃料子目标'], 12, True, TEXT, align=PP_ALIGN.CENTER)
x2 = MARGIN + pw + 0.6
panel(s, x2, y0, pw, ph, '我国 SAF 进展')
picture_fit(s, J('j3_saf_c919.jpg'), x2 + 0.4, y0 + 1.3, pw - 0.8, 3.7)
body(s, x2 + 0.4, y0 + 5.2, pw - 0.8, ph - 5.3, ['2024 年启动首批试点：**12** 个航班、**4** 座机场',
     '“十四五”累计消费目标 **5 万吨**', '2025 年底绿色航空燃料产能 **170 万吨/年**'], bullet=True, sa=4)
textbox(s, MARGIN, 14.5, CW, 1.6, ['SAF（Sustainable Aviation Fuel，可持续航空燃料）：用废油脂、农林废弃物、粪污等非化石原料生产的航空煤油，可按比例掺入传统航煤使用',
        'CO₂ 当量：把甲烷、氧化亚氮等温室气体按增温效应折算成二氧化碳的量'], 12, False, TEXT, space_after=0)
conclusion(s, '体会：粪污能源化的减排效益，需要在全生命周期边界内核算', y=16.4)
source(s, '中华人民共和国气候变化第一次双年透明度报告；IATA SAF Fact Sheet, 2026；EASA European Aviation Environmental Report 2025；民航局《民航绿色发展“十五五”实施方案》')

# =============== 体会 ===============
s = new_slide('', '课程学习体会'); S[19] = s
items = [('一、多方协同推进', '设施、装备和生产工艺需要〖相互衔接〗，共同推进机械化与智能化；新技术进入生产要经过一段磨合，应当给它改进的时间。'),
         ('二、兼顾投入与产出', '设施生产在追求产量和品质的同时，还要兼顾〖能源、空间和成本〗；增加的投入能否收回，决定了一项技术能否推广。'),
         ('三、持续探索', '研究成果做出来以后，还要〖继续挖掘用途、提升性能〗，并与不同领域的研究者合作，拓展成果的应用方向。'),
         ('四、服务生产实际', '农业工程问题涉及生物、环境、设施和装备，要放在〖整个生产系统〗中考虑，研究最终要服务于生产实际。')]
y = 2.8
for (hd, tx), n in zip(items, [2, 2, 1, 1]):
    hgt = 1.15 + n * 0.86
    tb = rich_box(s, MARGIN + 0.3, y, CW - 0.6, hgt, [(hd, 22, GREEN, 3), (tx, 19, TEXT, 0)])
    for pp in tb.text_frame.paragraphs[1:]: pp.line_spacing = 1.25
    y += hgt + 0.85
r = rect(s, MARGIN, 16.4, CW, 1.45, fill=GREEN_L, radius=0.25)
shape_text(r, ['未来，我希望在自己的研究中从生产实际出发，用系统的眼光看待问题，让研究成果用到生产中去。'], 20, True, GREEN)

# ---------- order ----------
order = [cover, toc, divs[0], S[4], S[6], S[7], S[5], divs[1], S[9], S[10], S[11], S[12],
         divs[2], S[14], S[15], S[16], S[17], divs[3], S[19], end]
lst = prs.slides._sldIdLst
idmap = {prs.part.related_part(x.get(qn('r:id'))): x for x in list(lst)}
for x in list(lst): lst.remove(x)
for o in order: lst.append(idmap[o.part])
prs.part.drop_rel([k for k, v in prs.part.rels.items() if v.target_part is sample.part][0])

md = open(TEXTMD, encoding='utf-8').read()
notes = {int(m.group(1)): m.group(2) for m in re.finditer(r'## 第 (\d+) 页[^\n]*\n(?:.*\n)*?备注：(.*)', md)}
for i, sl in enumerate(prs.slides, 1):
    if i in notes: sl.notes_slide.notes_text_frame.text = notes[i]
prs.save(OUT)
print('saved', len(prs.slides), 'notes', len(notes))
