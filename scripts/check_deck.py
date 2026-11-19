# -*- coding: utf-8 -*-
"""交给用户之前的自动检查。查的是我们第一版最常被指出的问题。

用法：python check_deck.py 成品.pptx [渲染图文件夹] [--min-font 12] [--out 报告.md]
渲染图文件夹：render.ps1 导出的 s01.png、s02.png……；给了才查留白。

检查项
1 形状超出页面
2 小字：正文字号低于 --min-font（来源、图注可到 10 pt，低于 10 pt 一律报）
3 文本框装不下（估算行数 × 行高 > 框高）或框里空得多（框高比文字多出 1 cm 以上）
4 连续 3 页以上没有图片（封面、目录、分隔页、致谢除外；这几类页也不查超出页面和留白）
8 模板水印、店铺素材残留（按形状名）
5 AI 味和偷懒的词句：不是……而是、赋能、一句话看懂、同左、同上、竖线分隔等
6 大片留白（需要渲染图）：内容区里最大一块连续空白占比
7 占位框没清（名字或文字里有“待放”“待补”“TODO”“XX”）
结果只是提示，最后仍要看渲染图。
"""
import sys, os, re, math, glob
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

CM = 360000
BANNED = [
    (r'不是[^。；，]{1,20}而是', '“不是……而是”对偶'),
    (r'不仅[^。；]{1,20}更', '“不仅……更”'),
    (r'赋能|助力|至关重要|全方位|多维度|深入探讨|综上所述|总而言之|值得注意的是|需要指出的是', '空洞套话'),
    (r'一句话看懂|你不知道的|竟然|揭秘', '新闻或营销腔标题'),
    (r'同左|同上', '表格写“同左/同上”，应合并单元格'),
    (r'\s[｜|]\s|[一-鿿]｜[一-鿿]', '竖线分隔'),
    (r'待放|待补|待核实|TODO|XXX|〔待', '占位或待核实没清'),
]


def walk(shapes):
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(s.shapes)
        else:
            yield s


def est_lines(text, width_cm, size_pt):
    em = sum(0.55 if ord(ch) < 128 else 1.0 for ch in text)
    per_line = max(1.0, (width_cm - 0.3) / (size_pt * 0.0353 * 1.02))
    return max(1, math.ceil(em / per_line))


def need_height(tf, width_cm):
    h = 0.0
    for p in tf.paragraphs:
        t = ''.join(r.text for r in p.runs)
        sz = max([r.font.size.pt for r in p.runs if r.font.size] or [18])
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        sa = Emu(p.space_after).pt if p.space_after is not None else 0
        n = est_lines(t, width_cm, sz) if t.strip() else 1
        h += n * sz * 0.0353 * 1.2 * ls + sa * 0.0353
    try:
        h += (tf.margin_top + tf.margin_bottom) / CM
    except Exception:
        h += 0.25
    return h


def blank_ratio(png):
    """内容区（去掉页眉 12%、页底 6%）切成 32×18 格，找最大连续空白块。"""
    from PIL import Image
    import numpy as np
    im = np.asarray(Image.open(png).convert('L'), dtype=float)
    H, W = im.shape
    top, bot = int(H * 0.12), int(H * 0.94)
    area = im[top:bot, int(W * 0.02):int(W * 0.98)]
    gh, gw = 18, 32
    ch, cw = area.shape[0] // gh, area.shape[1] // gw
    blank = [[False] * gw for _ in range(gh)]
    for i in range(gh):
        for j in range(gw):
            c = area[i * ch:(i + 1) * ch, j * cw:(j + 1) * cw]
            blank[i][j] = c.std() < 3 and c.mean() > 235
    # 最大全空白矩形（直方图法）
    best, best_pos = 0, None
    hist = [0] * gw
    for i in range(gh):
        for j in range(gw):
            hist[j] = hist[j] + 1 if blank[i][j] else 0
        stack = []
        for j in range(gw + 1):
            cur = hist[j] if j < gw else 0
            start = j
            while stack and stack[-1][1] >= cur:
                s, hgt = stack.pop()
                if hgt * (j - s) > best:
                    best, best_pos = hgt * (j - s), (i - hgt + 1, s, hgt, j - s)
                start = s
            stack.append((start, cur))
    total = gh * gw
    frac = best / total
    where = ''
    if best_pos:
        r0, c0, hh, ww = best_pos
        cy, cx = r0 + hh / 2, c0 + ww / 2
        where = ('上' if cy < gh / 3 else '下' if cy > 2 * gh / 3 else '中') + ('左' if cx < gw / 3 else '右' if cx > 2 * gw / 3 else '中')
    all_blank = sum(map(sum, blank)) / total
    return frac, where, all_blank


def main():
    args = sys.argv[1:]
    min_font, out = 12.0, None
    if '--min-font' in args:
        i = args.index('--min-font'); min_font = float(args[i + 1]); del args[i:i + 2]
    if '--out' in args:
        i = args.index('--out'); out = args[i + 1]; del args[i:i + 2]
    prs = Presentation(args[0])
    render = args[1] if len(args) > 1 else None
    W, H = prs.slide_width, prs.slide_height
    pngs = sorted(glob.glob(os.path.join(render, '*.png'))) if render else []

    rows, no_pic_run, issues_total = [], [], 0
    n = len(prs.slides)
    watermark = set()
    for idx, slide in enumerate(prs.slides, 1):
        msgs = []
        pre = sum(len(sh.text_frame.text.strip()) for sh in walk(slide.shapes) if sh.has_text_frame)
        is_frame = idx in (1, n) or pre < 60      # 封面、目录、分隔页、致谢
        pics, chars, small, tiny, total_chars = 0, 0, 0, 0, 0
        texts = []
        for s in walk(slide.shapes):
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pics += 1
            if re.search(r'淘宝|倒卖|水印|watermark|千图|包图|稻壳', s.name or '', re.I):
                watermark.add(idx)
            if s.left is not None and s.width is not None and not is_frame:
                if s.left < -CM * 0.05 or s.top < -CM * 0.05 or s.left + s.width > W + CM * 0.05 or s.top + s.height > H + CM * 0.05:
                    msgs.append('超出页面：%s' % s.name)
            if re.search(r'待放|placeholder|TODO', s.name or '', re.I):
                msgs.append('占位框没清：%s' % s.name)
            tfs = []
            if s.has_text_frame:
                tfs.append((s.text_frame, s.width, s.height, s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX))
            if getattr(s, 'has_table', False) and s.has_table:
                for row in s.table.rows:
                    for cell in row.cells:
                        tfs.append((cell.text_frame, None, None, False))
            for tf, w, h, is_tb in tfs:
                t = '\n'.join(''.join(r.text for r in p.runs) for p in tf.paragraphs)
                if not t.strip():
                    continue
                texts.append(t)
                for p in tf.paragraphs:
                    for r in p.runs:
                        k = len(r.text.strip())
                        total_chars += k
                        if r.font.size:
                            if r.font.size.pt < 10:
                                tiny += k
                            elif r.font.size.pt < min_font:
                                small += k
                if w and h and is_tb and h > CM * 0.4:
                    need = need_height(tf, w / CM)
                    have = h / CM
                    try:
                        auto = tf.auto_size
                    except Exception:
                        auto = None
                    if need > have * 1.25 + 0.3:
                        msgs.append('可能装不下（看渲染图确认）：“%s…” 需约 %.1f cm，框高 %.1f cm' % (t.strip()[:12], need, have))
                    elif have - need > 1.0 and (tf.vertical_anchor is None or str(tf.vertical_anchor).startswith('TOP')):
                        msgs.append('框里空：“%s…” 文字约 %.1f cm，框高 %.1f cm' % (t.strip()[:12], need, have))
        if tiny:
            msgs.append('有 %d 字小于 10 pt' % tiny)
        if total_chars and small / total_chars > 0.25:
            msgs.append('小字（<%g pt）占 %d%%' % (min_font, round(100 * small / total_chars)))
        joined = '\n'.join(texts)
        for pat, why in BANNED:
            m = re.search(pat, joined)
            if m:
                msgs.append('%s：“%s”' % (why, m.group(0).strip()))
        if not is_frame and pics == 0:
            no_pic_run.append(idx)
        else:
            if len(no_pic_run) >= 3:
                msgs.append('第 %d–%d 页连续没有图片' % (no_pic_run[0], no_pic_run[-1]))
            no_pic_run = []
        if pngs and idx <= len(pngs) and not is_frame:
            frac, where, allb = blank_ratio(pngs[idx - 1])
            if frac > 0.06:
                msgs.append('大块留白：最大空白块占内容区 %d%%（%s）' % (round(100 * frac), where))
        issues_total += len(msgs)
        rows.append((idx, total_chars, pics, msgs))

    L = ['# 自动检查：%s\n' % os.path.basename(args[0]),
         '共 %d 页，提示 %d 条。提示不等于错误，逐条对照渲染图判断。\n' % (n, issues_total)
         + ('\n模板水印或店铺素材残留（按形状名），在第 %s 页：交付前删掉或改名。\n' % '、'.join(map(str, sorted(watermark))) if watermark else ''),
         '| 页 | 字数 | 图片 | 提示 |', '|---|---|---|---|']
    for idx, chars, pics, msgs in rows:
        L.append('| %d | %d | %d | %s |' % (idx, chars, pics, '<br>'.join(msgs) if msgs else '—'))
    text = '\n'.join(L)
    if out:
        open(out, 'w', encoding='utf-8').write(text); print('写入', out, '提示', issues_total)
    else:
        sys.stdout.reconfigure(encoding='utf-8'); print(text)


if __name__ == '__main__':
    main()
