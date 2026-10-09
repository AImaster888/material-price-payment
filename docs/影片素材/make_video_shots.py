# -*- coding: utf-8 -*-
"""教學影片素材截圖（對應 docs\\教學影片腳本.md 的素材清單）。

用法：在 repo 根目錄執行  python docs\\影片素材\\make_video_shots.py
需要 Python 套件 playwright（用內建 chromium，headless）。
素材只取自 examples\\ 練習用假資料的核算報告、docs\\快速上手-原稿.html、GitHub 公開頁面。
"""
import html
import re
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
EX = ROOT / 'examples'
QUICK = ROOT / 'docs' / '快速上手-原稿.html'
REPORT_DETAIL = EX / '練習用-第一期估驗計價-請款明細表-核算報告.md'
REPORT_TOTAL = EX / '練習用-第一期估驗計價-計價總表-核算報告.md'
REPO_URL = 'https://github.com/AImaster888/material-price-payment'
SCALE = 2  # 截圖放大倍率，影片 1080p 放大也清楚

CSS = """
body{margin:0;background:#fff;font-family:"Microsoft JhengHei","Noto Sans TC",sans-serif;color:#22271f}
.card{padding:28px 36px;width:fit-content;max-width:1400px}
.file{font-size:15px;color:#6c716a;margin-bottom:10px}
h2{font-size:26px;margin:0 0 10px;color:#2f5233;border-bottom:2px solid #dcdfd6;padding-bottom:6px}
p{font-size:19px;margin:8px 0}
blockquote{margin:10px 0;padding:10px 16px;background:#fff8e1;border-left:5px solid #f0d98c;font-size:20px}
table{border-collapse:collapse;font-size:19px;margin-top:10px}
th,td{border:1px solid #c9ccc3;padding:8px 14px;white-space:nowrap}
th{background:#e7efe7}
td.n{text-align:right;font-variant-numeric:tabular-nums}
tr.hi td{background:#fdeaea;font-weight:700;color:#9c2b2b}
tr.hi2 td{background:#e3f1e5;font-weight:700;color:#186a2e}
.concl p{font-size:22px}
"""


def md_inline(s):
    s = html.escape(s)
    s = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', s)
    return re.sub(r'`(.+?)`', r'<code>\1</code>', s)


def md_to_html(lines, highlight):
    """極簡 markdown：## 標題、段落、表格、> 引言、粗體。highlight = {關鍵字: css class}。"""
    out, i = [], 0
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.startswith('## '):
            out.append('<h2>%s</h2>' % md_inline(ln[3:]))
        elif ln.startswith('> '):
            out.append('<blockquote>%s</blockquote>' % md_inline(ln[2:]))
        elif ln.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].startswith('|'):
                rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')])
                i += 1
            aligns = rows[1]
            t = ['<table><tr>%s</tr>' % ''.join('<th>%s</th>' % md_inline(c) for c in rows[0])]
            for r in rows[2:]:
                cls = next((v for k, v in highlight.items() if any(k in c for c in r)), '')
                tds = ''.join('<td%s>%s</td>' % (' class="n"' if a.endswith(':') else '', md_inline(c))
                              for c, a in zip(r, aligns))
                t.append('<tr class="%s">%s</tr>' % (cls, tds))
            out.append(''.join(t) + '</table>')
            continue
        elif ln and ln != '---':
            out.append('<p>%s</p>' % md_inline(ln))
        i += 1
    return '\n'.join(out)


def report_section(path, title):
    """取出報告中「## title」那一節（到下一個 ## 或 --- 為止）。"""
    lines = path.read_text(encoding='utf-8').splitlines()
    start = lines.index('## ' + title)
    end = next((j for j in range(start + 1, len(lines)) if lines[j].startswith('## ') or lines[j] == '---'),
               len(lines))
    return lines[start:end]


def report_conclusion(path):
    lines = path.read_text(encoding='utf-8').splitlines()
    return lines[lines.index('---') + 1:]


def shot_html(page, body, fname):
    page.set_content('<html><head><meta charset="utf-8"><style>%s</style></head><body>'
                     '<div class="card">%s</div></body></html>' % (CSS, body))
    page.locator('.card').screenshot(path=str(OUT / fname))
    print('saved:', fname)


def shot_report(page, path, title, fname, highlight=None, tag_class=''):
    body = '<div class="file">📄 %s</div>' % html.escape(path.name)
    body += '<div class="%s">%s</div>' % (tag_class, md_to_html(report_section(path, title), highlight or {}))
    shot_html(page, body, fname)


def shot_quick(page, selector, fname, mark=None, open_details=True):
    page.goto(QUICK.as_uri())
    page.evaluate("document.querySelectorAll('details').forEach(d => d.open = %s)" % ('true' if open_details else 'false'))
    if mark:  # 用紅框標出要觀眾注意的元素（只找這一段裡面的）
        page.locator(selector).first.locator(mark).first.evaluate("e => { e.style.outline = '4px solid #d62828'; e.style.outlineOffset = '3px' }")
    page.locator(selector).first.screenshot(path=str(OUT / fname))
    print('saved:', fname)


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1280, 'height': 800}, device_scale_factor=SCALE)

        # 1  GitHub 首頁、Code → Download ZIP
        page.goto(REPO_URL, wait_until='networkidle')
        page.screenshot(path=str(OUT / '01-GitHub首頁.png'))
        print('saved: 01-GitHub首頁.png')
        page.get_by_role('button', name=re.compile(r'^Code')).first.click()
        zip_item = page.get_by_text('Download ZIP').first
        zip_item.wait_for()
        zip_item.evaluate("e => { const t = e.closest('a,li') || e; t.style.outline = '4px solid #d62828' }")
        page.screenshot(path=str(OUT / '02-GitHub-Download-ZIP.png'))
        print('saved: 02-GitHub-Download-ZIP.png')

        # 3  快速上手手冊各步驟（第 0 步安裝、第 2 步複製、第 3 步結論表、第 5 步清單、練習對答案）
        page.set_viewport_size({'width': 1000, 'height': 800})
        shot_quick(page, 'section.step >> nth=0', '03-手冊-第0步安裝.png')
        shot_quick(page, 'section.step >> nth=2', '04-手冊-第2步複製這段話.png', mark='.say .copy')
        shot_quick(page, 'section.step >> nth=3', '05-手冊-第3步看結論.png')
        shot_quick(page, '#practice', '06-手冊-練習對答案.png', open_details=False)
        shot_quick(page, 'section.step >> nth=5', '07-手冊-第5步檢核清單.png')

        # 7  核算報告各段（練習檔）
        page.set_viewport_size({'width': 1500, 'height': 800})
        shot_report(page, REPORT_DETAIL, '跨期銜接比對', '08-報告-A-跨期銜接比對.png',
                    {'物價調整': 'hi', '總計': 'hi'})
        shot_report(page, REPORT_TOTAL, '前期累計 + 本期完成 = 累計至本期（前一期）',
                    '09-報告-B-前期加本期不等於累計.png', {'交通工程費': 'hi'})
        shot_report(page, REPORT_TOTAL, '畫面顯示文字比對', '10-報告-C-畫面顯示比對.png',
                    {'格式不同': 'hi2'})
        shot_report(page, REPORT_TOTAL, '單價 x 本期數量 = 複價（前一期）', '11-報告-D-略過未檢查.png')
        body = '<div class="file">📄 %s</div><div class="concl">%s</div>' % (
            html.escape(REPORT_TOTAL.name), md_to_html(report_conclusion(REPORT_TOTAL), {}))
        shot_html(page, body, '12-報告-計價總表結論.png')
        body = '<div class="file">📄 %s</div><div class="concl">%s</div>' % (
            html.escape(REPORT_DETAIL.name), md_to_html(report_conclusion(REPORT_DETAIL), {}))
        shot_html(page, body, '13-報告-請款明細表結論.png')

        browser.close()


if __name__ == '__main__':
    main()
