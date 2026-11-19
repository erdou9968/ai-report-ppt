# -*- coding: utf-8 -*-
"""把渲染图（文件夹里的 s01.png…）或一份 PDF 拼成总览图，方便一眼看全篇或发给用户。

用法：python make_sheet.py 渲染图文件夹或文件.pdf 输出.jpg [每行张数=4] [标题] [页码如 4,5,9]
没有 PowerPoint 时：soffice --headless --convert-to pdf 成品.pptx，再把 PDF 传给本脚本。
"""
import sys, os, glob, tempfile
from PIL import Image, ImageDraw, ImageFont


def load_pages(src):
    if src.lower().endswith('.pdf'):
        import pymupdf
        tmp = tempfile.mkdtemp()
        for i, p in enumerate(pymupdf.open(src), 1):
            p.get_pixmap(dpi=60).save(os.path.join(tmp, 's%02d.png' % i))
        src = tmp
    return sorted(glob.glob(os.path.join(src, '*.png')))


def font(bold, size):
    for f in (['C:/Windows/Fonts/msyhbd.ttc'] if bold else []) + ['C:/Windows/Fonts/msyh.ttc',
              '/System/Library/Fonts/PingFang.ttc', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc']:
        if os.path.exists(f):
            return ImageFont.truetype(f, size)
    return ImageFont.load_default()


def main():
    a = sys.argv[1:]
    files = load_pages(a[0]); out = a[1]
    cols = int(a[2]) if len(a) > 2 else 4
    title = a[3] if len(a) > 3 else ''
    nums = list(range(1, len(files) + 1))
    if len(a) > 4:
        nums = [int(x) for x in a[4].split(',')]
    W = 640
    ims = []
    for k in nums:
        im = Image.open(files[k - 1]).convert('RGB')
        ims.append((k, im.resize((W, int(W * im.height / im.width)), Image.LANCZOS)))
    H = ims[0][1].height
    gap, top = 14, (60 if title else 14)
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * W + (cols + 1) * gap, top + rows * (H + gap + 26)), 'white')
    d = ImageDraw.Draw(sheet)
    if title:
        d.text((gap, 12), title, fill='black', font=font(True, 30))
    for i, (k, im) in enumerate(ims):
        r, c = divmod(i, cols); x = gap + c * (W + gap); y = top + r * (H + gap + 26)
        sheet.paste(im, (x, y)); d.rectangle([x - 1, y - 1, x + W, y + H], outline='#999999')
        d.text((x, y + H + 2), '第 %d 页' % k, fill='#444444', font=font(False, 18))
    sheet.save(out, quality=88)
    print(out, sheet.size)


if __name__ == '__main__':
    main()
