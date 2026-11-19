# -*- coding: utf-8 -*-
"""Downsample large images inside a pptx (keeps format and file names)."""
import sys, zipfile, io, os
from PIL import Image

src, dst = sys.argv[1], sys.argv[2]
MAX_SIDE = 1800
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
saved = 0
for item in zin.infolist():
    data = zin.read(item.filename)
    low = item.filename.lower()
    if low.startswith('ppt/media/') and low.endswith(('.jpg', '.jpeg', '.png')) and len(data) > 300_000:
        try:
            im = Image.open(io.BytesIO(data))
            fmt = 'PNG' if low.endswith('.png') else 'JPEG'
            if max(im.size) > MAX_SIDE:
                im.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
            buf = io.BytesIO()
            if fmt == 'JPEG':
                im.convert('RGB').save(buf, 'JPEG', quality=85, optimize=True)
            else:
                if im.mode in ('RGBA', 'LA', 'P'):
                    im.save(buf, 'PNG', optimize=True)
                else:   # opaque png photo -> keep png but quantise-free optimise
                    im.save(buf, 'PNG', optimize=True)
            if buf.tell() < len(data):
                saved += len(data) - buf.tell()
                data = buf.getvalue()
        except Exception as e:
            print('skip', item.filename, e)
    zout.writestr(item, data)
zout.close()
print('saved MB', round(saved / 1e6, 1), 'new size MB', round(os.path.getsize(dst) / 1e6, 1))
