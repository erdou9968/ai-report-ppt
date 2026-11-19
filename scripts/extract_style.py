# -*- coding: utf-8 -*-
"""从一份参照 PPT 里统计样式，输出 markdown 报告，用来写样式表。

用法：python extract_style.py 参照.pptx [报告.md] [--pages 4-15]
统计内容：页面尺寸；字体；字号（按字数加权）及各字号的加粗比例和常用颜色；
行距和段后；形状种类、圆角、线宽、虚线、渐变、透明度；填充色和线色；
每页的字数、段落数、图片数、图片面积占比、表格数；标题位置。
"""
import sys, re, collections
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

CM = 360000


def cm(v):
    return round(v / CM, 2) if v is not None else None


def walk(shapes):
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(s.shapes)
        else:
            yield s


def run_color(run):
    try:
        if run.font.color and run.font.color.type is not None and run.font.color.rgb is not None:
            return str(run.font.color.rgb)
    except Exception:
        pass
    try:
        return 'theme:' + str(run.font.color.theme_color).split('.')[-1].split(' ')[0]
    except Exception:
        return '继承'


def run_font(run):
    rPr = run._r.find(qn('a:rPr'))
    ea = None
    if rPr is not None:
        e = rPr.find(qn('a:ea'))
        if e is not None:
            ea = e.get('typeface')
    return run.font.name or ea or '继承'


def iter_text_frames(shape):
    if shape.has_text_frame:
        yield shape.text_frame
    if getattr(shape, 'has_table', False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                yield cell.text_frame


def parse_pages(arg, n):
    if not arg:
        return set(range(1, n + 1))
    out = set()
    for part in arg.split(','):
        if '-' in part:
            a, b = part.split('-'); out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return out


def main():
    args = [a for a in sys.argv[1:]]
    pages_arg = None
    if '--pages' in args:
        i = args.index('--pages'); pages_arg = args[i + 1]; del args[i:i + 2]
    src = args[0]
    out = args[1] if len(args) > 1 else None
    prs = Presentation(src)
    W, H = prs.slide_width, prs.slide_height
    pages = parse_pages(pages_arg, len(prs.slides))

    fonts = collections.Counter()
    sizes = collections.Counter()
    size_bold = collections.Counter()
    size_color = collections.defaultdict(collections.Counter)
    run_colors = collections.Counter()
    spacing = collections.Counter()
    space_after = collections.Counter()
    shape_kinds = collections.Counter()
    radii = collections.Counter()
    line_w = collections.Counter()
    fills = collections.Counter()
    lines = collections.Counter()
    dash = grad = alpha = 0
    per_slide = []
    titles = []

    for idx, slide in enumerate(prs.slides, 1):
        if idx not in pages:
            continue
        chars = paras = pics = tables = 0
        pic_area = 0
        biggest_top = None
        for s in walk(slide.shapes):
            st = s.shape_type
            if st == MSO_SHAPE_TYPE.PICTURE:
                pics += 1
                if s.width and s.height:
                    pic_area += s.width * s.height
            if getattr(s, 'has_table', False) and s.has_table:
                tables += 1
            if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
                try:
                    kind = str(s.auto_shape_type).split('.')[-1].split(' ')[0]
                except Exception:
                    kind = 'AUTO'
                shape_kinds[kind] += 1
                try:
                    if len(s.adjustments):
                        radii[round(s.adjustments[0], 2)] += 1
                except Exception:
                    pass
                spPr = s._element.spPr
                if spPr.find(qn('a:gradFill')) is not None:
                    grad += 1
                try:
                    if s.fill.type == 1:
                        fills[str(s.fill.fore_color.rgb)] += 1
                except Exception:
                    pass
                ln = spPr.find(qn('a:ln'))
                if ln is not None:
                    if ln.get('w'):
                        line_w[round(int(ln.get('w')) / 12700, 2)] += 1
                    if ln.find(qn('a:prstDash')) is not None:
                        dash += 1
                    sf = ln.find(qn('a:solidFill'))
                    if sf is not None and len(sf):
                        lines[sf[0].get('val') or sf[0].tag.split('}')[-1]] += 1
                        if sf[0].find(qn('a:alpha')) is not None:
                            alpha += 1
            elif st == MSO_SHAPE_TYPE.TEXT_BOX:
                shape_kinds['文本框'] += 1
            for tf in iter_text_frames(s):
                for p in tf.paragraphs:
                    t = ''.join(r.text for r in p.runs)
                    if not t.strip():
                        continue
                    paras += 1
                    if p.line_spacing is not None:
                        spacing[round(p.line_spacing, 2) if isinstance(p.line_spacing, float) else str(p.line_spacing)] += len(t)
                    if p.space_after is not None:
                        space_after[round(Emu(p.space_after).pt, 1)] += 1
                    for r in p.runs:
                        n = len(r.text.strip())
                        if not n:
                            continue
                        chars += n
                        sz = r.font.size.pt if r.font.size else '继承'
                        fonts[run_font(r)] += n
                        sizes[sz] += n
                        if r.font.bold:
                            size_bold[sz] += n
                        c = run_color(r)
                        size_color[sz][c] += n
                        run_colors[c] += n
                        if isinstance(sz, float) and s.top is not None and s.top < H * 0.12:
                            if biggest_top is None or sz > biggest_top[0]:
                                biggest_top = (sz, cm(s.left), cm(s.top), s.text_frame.text.strip()[:20] if s.has_text_frame else r.text[:20])
        per_slide.append((idx, chars, paras, pics, round(100 * pic_area / (W * H)), tables))
        if biggest_top:
            titles.append(biggest_top)

    L = []
    L.append('# 样式提取报告：%s\n' % src.replace('\\', '/').split('/')[-1])
    L.append('页面：%.2f × %.2f cm（宽高比 %.3f；16:9 = 1.778）；统计页：%s\n' % (cm(W), cm(H), W / H, pages_arg or '全部'))
    tot = sum(sizes.values()) or 1
    L.append('## 字体（按字数）')
    for f, n in fonts.most_common(6):
        L.append('- %s：%d%%' % (f, round(100 * n / tot)))
    L.append('\n## 字号层级（按字数占比；加粗比例；最常用颜色）')
    L.append('| 字号 pt | 字数占比 | 加粗 | 常用颜色 |')
    L.append('|---|---|---|---|')
    for sz, n in sorted(sizes.items(), key=lambda x: (isinstance(x[0], str), -x[0] if not isinstance(x[0], str) else 0)):
        if n / tot < 0.005:
            continue
        cols = '、'.join('%s(%d%%)' % (c, round(100 * k / n)) for c, k in size_color[sz].most_common(3))
        L.append('| %s | %d%% | %d%% | %s |' % (sz, round(100 * n / tot), round(100 * size_bold[sz] / n), cols))
    L.append('\n## 文字颜色（按字数）')
    for c, n in run_colors.most_common(8):
        L.append('- %s：%d%%' % (c, round(100 * n / tot)))
    L.append('\n## 段落')
    L.append('- 行距（按字数）：' + '，'.join('%s：%d%%' % (k, round(100 * v / tot)) for k, v in spacing.most_common(5)) + '（未列出的为继承，通常 1.0）')
    L.append('- 段后（pt，按段数）：' + '，'.join('%s：%d' % kv for kv in space_after.most_common(5)))
    L.append('\n## 形状')
    L.append('- 种类：' + '，'.join('%s %d' % kv for kv in shape_kinds.most_common(10)))
    L.append('- 圆角参数（第一调整值，越小越方）：' + '，'.join('%s：%d' % kv for kv in radii.most_common(6)))
    L.append('- 线宽（磅）：' + '，'.join('%s：%d' % kv for kv in line_w.most_common(6)))
    L.append('- 虚线 %d 个；渐变填充 %d 个；线条带透明度 %d 个' % (dash, grad, alpha))
    L.append('- 填充色：' + '，'.join('%s：%d' % kv for kv in fills.most_common(8)))
    L.append('- 线条色：' + '，'.join('%s：%d' % kv for kv in lines.most_common(6)))
    if titles:
        tsz = collections.Counter(t[0] for t in titles).most_common(1)[0][0]
        lefts = collections.Counter(t[1] for t in titles).most_common(1)[0][0]
        L.append('\n## 页面标题')
        L.append('- 最常见：%s pt，左边距约 %s cm；例：%s' % (tsz, lefts, '；'.join(t[3] for t in titles[:3])))
    L.append('\n## 逐页密度')
    L.append('| 页 | 字数 | 段落 | 图片 | 图片面积占比 | 表格 |')
    L.append('|---|---|---|---|---|---|')
    for row in per_slide:
        L.append('| %d | %d | %d | %d | %d%% | %d |' % row)
    content = [r for r in per_slide if r[1] > 60]
    if content:
        L.append('\n内容页（字数 > 60）平均：字数 %d，图片 %.1f 张，图片面积 %d%%。' % (
            sum(r[1] for r in content) / len(content), sum(r[3] for r in content) / len(content),
            sum(r[4] for r in content) / len(content)))
    L.append('\n说明：图片面积按图片外框算，裁切后的实际可见面积会小一些；“继承”表示字号或字体写在母版或主题里，需要打开母版核对。')
    text = '\n'.join(L)
    if out:
        open(out, 'w', encoding='utf-8').write(text)
        print('写入', out)
    else:
        sys.stdout.reconfigure(encoding='utf-8')
        print(text)


if __name__ == '__main__':
    main()
