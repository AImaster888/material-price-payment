# -*- coding: utf-8 -*-
"""把 字卡.html 每一張 section 截成 1920x1080 PNG，存到 字卡\\。
清單類字卡（c04 五項檢查、c12 回顧）另外產生逐項反白版本 cXX_hlN.png。
用法：python docs\\影片素材\\render_cards.py
"""
from pathlib import Path

from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
OUT = HERE / '字卡'
HL = {'c04': 5, 'c12': 4}  # 要做反白版本的字卡：項目數

OUT.mkdir(exist_ok=True)
with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={'width': 1920, 'height': 1080})
    page.goto((HERE / '字卡.html').as_uri())
    page.evaluate('document.fonts.ready')
    for cid in page.eval_on_selector_all('section.card', 'els => els.map(e => e.id)'):
        page.evaluate("id => { document.querySelectorAll('.card').forEach(c => c.classList.toggle('on', c.id === id));"
                      " document.body.removeAttribute('data-hl') }", cid)
        page.screenshot(path=str(OUT / (cid + '.png')))
        for n in range(1, HL.get(cid, 0) + 1):
            page.evaluate("n => document.body.setAttribute('data-hl', n)", str(n))
            page.screenshot(path=str(OUT / ('%s_hl%d.png' % (cid, n))))
        print('saved:', cid)
    browser.close()
