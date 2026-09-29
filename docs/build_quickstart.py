# -*- coding: utf-8 -*-
"""
把 快速上手-原稿.html 裡引用的 img/*.png 包進 HTML，產生可以單獨傳給別人的 快速上手.html。

原稿用相對路徑引用圖片，只傳 HTML 給別人時對方沒有 img 資料夾，圖片會看不到。
用法：python docs\build_quickstart.py（在任何目錄執行都可以）
"""
import base64
import re
from pathlib import Path

DOCS = Path(__file__).parent
SRC = DOCS / '快速上手-原稿.html'
OUT = DOCS / '快速上手.html'
BANNER = '<!-- 此檔由 build_quickstart.py 自動產生（圖片已內嵌），要改內容請改 快速上手-原稿.html -->\n'


def inline(match):
    img = DOCS / match.group(1)
    data = base64.b64encode(img.read_bytes()).decode('ascii')
    return 'src="data:image/png;base64,{0}"'.format(data)


html = SRC.read_text(encoding='utf-8')
html, n = re.subn(r'src="(img/[^"]+\.png)"', inline, html)
doctype, rest = html.split('\n', 1)  # 註解放在 DOCTYPE 之後，放前面舊瀏覽器會進怪異模式
OUT.write_text(doctype + '\n' + BANNER + rest, encoding='utf-8')
print('已產生 {0}（內嵌 {1} 張圖，{2:,} KB）'.format(OUT.name, n, OUT.stat().st_size // 1024))
