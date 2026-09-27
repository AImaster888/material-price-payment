# -*- coding: utf-8 -*-
"""把 Excel 匯出的單頁 PDF 轉成 PNG，並裁掉四周白邊。用法：python pdf_to_png.py 輸入.pdf 輸出.png"""
import sys

import fitz
from PIL import Image, ImageChops

src, dst = sys.argv[1], sys.argv[2]
page = fitz.open(src)[0]
pix = page.get_pixmap(dpi=160)
img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
bbox = ImageChops.difference(img, Image.new('RGB', img.size, 'white')).getbbox()
if bbox:
    pad = 12
    img = img.crop((max(bbox[0] - pad, 0), max(bbox[1] - pad, 0),
                    min(bbox[2] + pad, img.width), min(bbox[3] + pad, img.height)))
img.save(dst)
print('saved: {0} ({1}x{2})'.format(dst, img.width, img.height))
