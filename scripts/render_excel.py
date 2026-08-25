# -*- coding: utf-8 -*-
"""
render_excel.py — 把 Excel 儲存格「實際運算出來的值」轉成「畫面上會顯示的文字」。

背景（為什麼需要這個檔案）：
    跨檔案核對數字時，很容易只比較 openpyxl 讀到的底層 float 值，數學上相等就判定
    「相符」。但審查文件的人看的是螢幕/紙本上顯示的文字，不是儲存格的完整精度。
    兩份檔案的儲存格顯示格式（小數位數、千分位、零值顯示方式）很可能不一樣——同一個
    數字在不同格式設定下，畫面顯示出來可能完全不像同一回事（例如 0.367188 在一份檔案
    顯示成「0.37」，另一份顯示成「0.367188」）。

    這個模組把 openpyxl 讀到的 (value, number_format) 轉成「Excel 會顯示的文字」，
    讓比對程式可以做「畫面文字是否一致」的比對，而不是只比底層數值。這是一個獨立、
    盡力而為的實作，涵蓋工程估驗計價文件常見的格式（千分位、固定小數位、會計格式的
    零值顯示為 "-"、負數括號），不是完整的 Excel number_format 解析器。
"""
import re


def _decimals_in_section(section):
    """從格式字串的一個區段抓小數位數，例如 '#,##0.00' -> 2。"""
    m = re.search(r'0\.(0+)', section)
    return len(m.group(1)) if m else 0


def _zero_shows_dash(section):
    """會計格式常見寫法：零值區段是 '"-"' 或類似，畫面顯示成一條槓。"""
    return '"-"' in section or "'-'" in section


def render_display(value, number_format):
    """
    依照 number_format 把 value 轉成「Excel 畫面上會顯示的文字」。

    參數：
        value: openpyxl 用 data_only=True 讀到的儲存格值（int/float/None/str）
        number_format: openpyxl 用 data_only=False 讀到的 cell.number_format

    回傳：字串，模擬 Excel 顯示結果。無法判斷小數位數時預設當整數（0位）處理。
    """
    if value is None:
        value = 0
    if not isinstance(value, (int, float)):
        return str(value)

    sections = (number_format or '#,##0').split(';')
    pos = sections[0] if len(sections) > 0 else '#,##0'
    neg = sections[1] if len(sections) > 1 else None
    zero = sections[2] if len(sections) > 2 else None

    if value == 0 and zero is not None and _zero_shows_dash(zero):
        return '-'

    section = pos
    negative = value < 0
    if negative and neg is not None:
        section = neg
        value = abs(value)

    # 百分比格式：值本身是分數（0.05），畫面顯示成 "5.00%"
    if '%' in section:
        nd = _decimals_in_section(section)
        text = '{0:,.{1}f}%'.format(round(value * 100, nd), nd)
        return ('-' + text) if negative else text

    nd = _decimals_in_section(section)
    rounded = round(value, nd)
    text = '{0:,.{1}f}'.format(rounded, nd)
    if negative:
        text = '(' + text + ')' if neg is None else '-' + text
    return text


def compare_display(v1, f1, v2, f2):
    """回傳 (顯示文字1, 顯示文字2, 是否一致)。"""
    d1 = render_display(v1, f1)
    d2 = render_display(v2, f2)
    return d1, d2, (d1 == d2)


if __name__ == '__main__':
    # 快速自我測試 — 對應本工具實際踩過的案例
    cases = [
        (0.3671875, '#,##0.00', '0.37'),
        (0.3671875, '#,##0.000000', '0.367188'),
        (0, '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"_-;_-@_-', '-'),
        (1015000, '#,##0', '1,015,000'),
        (0.05, '0.0%', '5.0%'),
    ]
    ok = True
    for value, fmt, expected in cases:
        got = render_display(value, fmt)
        status = 'OK' if got == expected else 'FAIL'
        if got != expected:
            ok = False
        print('{0}: render_display({1!r}, {2!r}) = {3!r} (expected {4!r})'.format(
            status, value, fmt, got, expected))
    raise SystemExit(0 if ok else 1)
