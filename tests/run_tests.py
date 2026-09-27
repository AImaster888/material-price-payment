# -*- coding: utf-8 -*-
"""
run_tests.py — 用合成的假資料（不含任何真實案件內容）驗證 verify_payment.py 的核算邏輯。

跑法：python tests/run_tests.py
全綠才算過，任何一項失敗都會印出來並讓 exit code = 1。
"""
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))
from verify_payment import (SheetLayout, check_extended_price, check_within_period,  # noqa: E402
                             check_subitem_sum, check_cross_period, check_display_text_cross,
                             summarize, md_report_path, DISPLAY_CAUSE_FORMAT, DISPLAY_CAUSE_VALUE)

FAILURES = []


def check(name, condition, detail=''):
    status = 'PASS' if condition else 'FAIL'
    print('[{0}] {1}{2}'.format(status, name, (' — ' + detail) if detail and not condition else ''))
    if not condition:
        FAILURES.append(name)


def make_sheet(wb, name, rows, header_row=4):
    """rows: list of dict，鍵是 code/unit/qty0/price0/amt0(原契約)/pq/pa(前期)/cq/ca(本期)/tq/ta(累計)"""
    ws = wb.create_sheet(name)
    ws['A' + str(header_row)] = '項目'
    ws['B' + str(header_row)] = '工程項目'
    ws['C' + str(header_row)] = '原契約'
    ws['G' + str(header_row)] = '前期累計'
    ws['I' + str(header_row)] = '本期完成'
    ws['K' + str(header_row)] = '累計至本期'
    sub = header_row + 1
    ws['C' + str(sub)] = '單位'
    ws['D' + str(sub)] = '數量'
    ws['E' + str(sub)] = '單價'
    ws['F' + str(sub)] = '複價'
    ws['G' + str(sub)] = '數量'
    ws['H' + str(sub)] = '複價'
    ws['I' + str(sub)] = '數量'
    ws['J' + str(sub)] = '複價'
    ws['K' + str(sub)] = '數量'
    ws['L' + str(sub)] = '複價'

    r = sub + 1
    for row in rows:
        ws.cell(row=r, column=1, value=row.get('code'))
        ws.cell(row=r, column=2, value=row.get('name'))
        if 'unit' in row:
            ws.cell(row=r, column=3, value=row['unit'])
        if 'price' in row:
            ws.cell(row=r, column=5, value=row['price'])
        if 'pq' in row:
            ws.cell(row=r, column=7, value=row['pq'])
        if 'pa' in row:
            ws.cell(row=r, column=8, value=row['pa'])
        if 'cq' in row:
            ws.cell(row=r, column=9, value=row['cq'])
        if 'ca_formula' in row:
            ws.cell(row=r, column=10, value=row['ca_formula'])
        elif 'ca' in row:
            ws.cell(row=r, column=10, value=row['ca'])
        if 'tq' in row:
            ws.cell(row=r, column=11, value=row['tq'])
        if 'ta' in row:
            ws.cell(row=r, column=12, value=row['ta'])
        if 'fmt_j' in row:
            ws.cell(row=r, column=10).number_format = row['fmt_j']
        r += 1
    return ws


def build_workbook(rows):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    make_sheet(wb, '明細表', rows)
    return wb


def save_and_reload(wb):
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    wb.save(tmp.name)
    wf = openpyxl.load_workbook(tmp.name, data_only=False)['明細表']
    wv = openpyxl.load_workbook(tmp.name, data_only=True)['明細表']
    Path(tmp.name).unlink()
    return wf, wv


# ------------------------------------------------------------------
# extended_price：單價 x 本期數量 = 本期複價
# ------------------------------------------------------------------
rows = [
    dict(code='壹.一', name='正確項目', price=100, cq=5, ca=500),
    dict(code='壹.二', name='錯誤項目', price=100, cq=5, ca=999),
]
wf, wv = save_and_reload(build_workbook(rows))
layout = SheetLayout(wf)
results, skip = check_extended_price(wf, wv, layout, group='本期完成')
check('extended_price 偵測到2項', len(results) == 2)
check('extended_price 正確項目判定為 ok', results[0]['ok'] is True)
check('extended_price 錯誤項目判定為不 ok', results[1]['ok'] is False)

# 四捨五入要跟 Excel 的 ROUND 一致（不是 Python round() 的銀行家捨入）：
# 0.5 x 2469 = 1234.5，Excel 給 1235；用 round() 會給 1234 而誤判成異常
rows = [
    dict(code='壹.一', name='剛好.5進位', price=0.5, cq=2469, ca=1235),
    dict(code='壹.二', name='剛好.5進位2', price=0.5, cq=2467, ca=1234),
]
wf, wv = save_and_reload(build_workbook(rows))
results, skip = check_extended_price(wf, wv, SheetLayout(wf), group='本期完成')
check('extended_price .5 依 Excel 規則進位，不誤判',
      results[0]['ok'] is True and results[1]['ok'] is True,
      str([(x['expected'], x['actual']) for x in results]))

# ------------------------------------------------------------------
# within_period：前期 + 本期 = 累計
# ------------------------------------------------------------------
rows = [
    dict(code='壹.一', name='正確', pa=100, ca=50, ta=150),
    dict(code='壹.二', name='錯誤', pa=100, ca=50, ta=999),
    dict(code='壹.三', name='式列無數量', pa=None, ca=None, ta=None),  # 不該被判為異常
]
wf, wv = save_and_reload(build_workbook(rows))
layout = SheetLayout(wf)
results, skip = check_within_period(wv, layout)
check('within_period 偵測到3項', len(results) == 3)
check('within_period 正確項目 ok', results[0]['amt_ok'] is True)
check('within_period 錯誤項目不 ok', results[1]['amt_ok'] is False)
check('within_period 全空列不誤判為異常（數量）', results[2]['qty_ok'] is True)
check('within_period 全空列不誤判為異常（複價）', results[2]['amt_ok'] is True, str(results[2]))

# 大項標題列：有名稱、三個複價欄都空的，不該被判成「累計欄漏填」
rows = [
    dict(code='壹', name='大項標題列（無金額）'),
    dict(code='壹.一', name='正常', pa=100, ca=50, ta=150),
]
wf, wv = save_and_reload(build_workbook(rows))
results, skip = check_within_period(wv, SheetLayout(wf))
check('within_period 大項標題列不誤判為異常',
      results[0]['qty_ok'] is True and results[0]['amt_ok'] is True, str(results[0]))

# ------------------------------------------------------------------
# subitem_sum：小計列 SUM 公式 vs 子項加總
# openpyxl 不會真的計算公式，沒辦法用它自然產生「公式在、但快取值跟子項和不一致」
# 的真實檔案，所以這裡用最小的假物件（只實作 .cell(row,column).value 介面）直接
# 測 check_subitem_sum 的比對邏輯，不牽涉檔案存取。
# ------------------------------------------------------------------
class _FakeCell:
    def __init__(self, value):
        self.value = value


class _FakeSheet:
    def __init__(self, data):
        self.data = data  # {(row, col): value}

    def cell(self, row, column):
        return _FakeCell(self.data.get((row, column)))


class _FakeLayout:
    def __init__(self, code_col, name_col, first_data_row, last_data_row, groups):
        self.code_col = code_col
        self.name_col = name_col
        self.first_data_row = first_data_row
        self.last_data_row = last_data_row
        self.groups = groups

    def col(self, group, subcol):
        return self.groups.get(group, {}).get(subcol)


fake_data_wf = {
    (6, 1): '壹', (6, 2): '小計正確', (6, 10): '=SUM(J7:J8)',
    (7, 1): '壹.一', (7, 2): '子項1', (7, 10): 300,
    (8, 1): '壹.二', (8, 2): '子項2', (8, 10): 200,
    (9, 1): '貳', (9, 2): '小計錯誤', (9, 10): '=SUM(J10:J11)',
    (10, 1): '貳.一', (10, 2): '子項1', (10, 10): 100,
    (11, 1): '貳.二', (11, 2): '子項2', (11, 10): 100,
}
fake_data_wv = dict(fake_data_wf)
fake_data_wv[(6, 10)] = 500   # 快取值：跟子項(300+200)相符
fake_data_wv[(9, 10)] = 999   # 快取值：跟子項(100+100=200)不相符——模擬公式範圍拉錯或資料被覆蓋

fake_wf = _FakeSheet(fake_data_wf)
fake_wv = _FakeSheet(fake_data_wv)
fake_layout = _FakeLayout(1, 2, 6, 11, {'本期完成': {'複價': 10}})
results, skip = check_subitem_sum(fake_wf, fake_wv, fake_layout, group='本期完成')
check('subitem_sum 偵測到2個 SUM 小計', len(results) == 2, str(results))
if len(results) == 2:
    check('subitem_sum 正確小計判定為 ok', results[0]['ok'] is True, str(results[0]))
    check('subitem_sum 錯誤小計判定為不 ok', results[1]['ok'] is False, str(results[1]))

# 絕對位址 $ 與小寫 sum 也要認得（Excel 另存或手工改公式很常出現）
fake_abs_wf = {
    (6, 1): '壹', (6, 2): '小計($絕對位址)', (6, 10): '=SUM($J$7:$J$8)',
    (7, 10): 300, (8, 10): 200,
    (9, 1): '貳', (9, 2): '小計(小寫sum)', (9, 10): '=sum(J10:J11)',
    (10, 10): 100, (11, 10): 100,
}
fake_abs_wv = dict(fake_abs_wf)
fake_abs_wv[(6, 10)] = 500
fake_abs_wv[(9, 10)] = 200
results = check_subitem_sum(_FakeSheet(fake_abs_wf), _FakeSheet(fake_abs_wv), fake_layout,
                            group='本期完成')[0]
check('subitem_sum 認得 $ 絕對位址與小寫 sum', len(results) == 2, str(results))
check('subitem_sum $ 絕對位址的小計判定正確',
      all(x['ok'] for x in results), str(results))

# ------------------------------------------------------------------
# cross_period：A檔累計至本期 vs B檔前期累計（含無代碼、用名稱配對的總計列）
# ------------------------------------------------------------------
rows_a = [
    dict(code='壹.一', name='一致項目', ta=100),
    dict(code='壹.二', name='不一致項目', ta=50),
    dict(code=None, name='總計', ta=150),
]
rows_b = [
    dict(code='壹.一', name='一致項目', pa=100),
    dict(code='壹.二', name='不一致項目', pa=999),
    dict(code=None, name='總計', pa=150),
]
wf_a, wv_a = save_and_reload(build_workbook(rows_a))
wf_b, wv_b = save_and_reload(build_workbook(rows_b))
layout_a, layout_b = SheetLayout(wf_a), SheetLayout(wf_b)
results, skip = check_cross_period(wv_a, layout_a, wv_b, layout_b)
by_code = {x['code']: x for x in results}
check('cross_period 一致項目判定為相符', by_code['壹.一']['amt_ok'] is True)
check('cross_period 不一致項目判定為不符', by_code['壹.二']['amt_ok'] is False)
none_rows = [x for x in results if x['code'] is None]
check('cross_period 無代碼列用名稱配對成功', len(none_rows) == 1 and none_rows[0]['matched'] is True,
      str(none_rows))
if none_rows:
    check('cross_period 無代碼列(總計)判定為相符', none_rows[0]['amt_ok'] is True)

# 項次代碼型別不同（一份存成數字 1、另一份存成文字 "1"）仍要配得起來
wf_a, wv_a = save_and_reload(build_workbook([dict(code=1, name='項目一', ta=100)]))
wf_b, wv_b = save_and_reload(build_workbook([dict(code='1', name='項目一', pa=100)]))
results = check_cross_period(wv_a, SheetLayout(wf_a), wv_b, SheetLayout(wf_b))[0]
check('cross_period 代碼型別不同仍配對成功',
      results[0]['matched'] is True and results[0]['amt_ok'] is True, str(results[0]))

# 本期有、前一期沒有的列：前期累計有金額 -> 異常；前期累計是 0（本期新增）-> 不算異常
wf_a, wv_a = save_and_reload(build_workbook([dict(code='壹.一', name='舊項目', ta=100)]))
wf_b, wv_b = save_and_reload(build_workbook([
    dict(code='壹.一', name='舊項目', pa=100),
    dict(code='壹.二', name='本期新增(前期累計0)', pa=0),
    dict(code='壹.三', name='前一期沒有卻有前期累計', pa=500),
]))
results = check_cross_period(wv_a, SheetLayout(wf_a), wv_b, SheetLayout(wf_b))[0]
reverse = [x for x in results if x['a_row'] is None]
check('cross_period 反向掃到「前一期沒有、本期前期累計卻有金額」的列',
      len(reverse) == 1 and reverse[0]['code'] == '壹.三', str(reverse))
check('cross_period 本期新增(前期累計0)不誤報', all(x['code'] != '壹.二' for x in reverse))

# 本期有兩列同樣的項次 -> 不能亂配，要判為無法配對
wf_a, wv_a = save_and_reload(build_workbook([dict(code='壹.一', name='項目', ta=100)]))
wf_b, wv_b = save_and_reload(build_workbook([
    dict(code='壹.一', name='項目', pa=100),
    dict(code='壹.一', name='項目(重複代碼)', pa=999),
]))
results = check_cross_period(wv_a, SheetLayout(wf_a), wv_b, SheetLayout(wf_b))[0]
dup_row = [x for x in results if x['a_row'] is not None][0]
check('cross_period 本期重複代碼判為無法配對',
      dup_row['matched'] is False and '無法確定' in dup_row['note'], str(dup_row))

# ------------------------------------------------------------------
# display_text：底層數值相同，但顯示格式不同 -> 應判定為不一致
# ------------------------------------------------------------------
rows_a = [dict(code='壹.一', name='格式不同', ta=0.3671875)]
rows_b = [dict(code='壹.一', name='格式不同', pa=0.3671875)]
wb_a = build_workbook(rows_a)
wb_a['明細表']['L6'].number_format = '#,##0.00'
wf_a, wv_a = save_and_reload(wb_a)
wb_b = build_workbook(rows_b)
wb_b['明細表']['H6'].number_format = '#,##0.000000'
wf_b, wv_b = save_and_reload(wb_b)
layout_a, layout_b = SheetLayout(wf_a), SheetLayout(wf_b)
results, skip = check_display_text_cross(wv_a, wf_a, layout_a, wv_b, wf_b, layout_b)
check('display_text 偵測到1項', len(results) == 1, str(results))
if results:
    check('display_text 底層相同但格式不同 -> 顯示不一致',
          results[0]['amt_match'] is False, str(results[0]))
    check('display_text 顯示文字確實不同（非誤判）',
          results[0]['amt_disp_a'] != results[0]['amt_disp_b'])
    check('display_text 底層相同 -> 原因標為「格式不同（數字相同）」',
          results[0]['cause'] == DISPLAY_CAUSE_FORMAT, str(results[0]))

# display_text：數字真的不同（例如累計數量被誤植成固定值）-> 原因要標「數字不同」，不可說成格式問題
rows_a = [dict(code='壹.二', name='數字不同', tq=30, ta=100)]
rows_b = [dict(code='壹.二', name='數字不同', pq=0.3, pa=100)]
wf_a, wv_a = save_and_reload(build_workbook(rows_a))
wf_b, wv_b = save_and_reload(build_workbook(rows_b))
results = check_display_text_cross(wv_a, wf_a, SheetLayout(wf_a), wv_b, wf_b, SheetLayout(wf_b))[0]
check('display_text 底層不同 -> 原因標為「數字不同」',
      len(results) == 1 and results[0]['cause'] == DISPLAY_CAUSE_VALUE, str(results))

# display_text：完全一致的列不標原因
rows_same = [dict(code='壹.三', name='一致', tq=1, ta=100)]
rows_same_b = [dict(code='壹.三', name='一致', pq=1, pa=100)]
wf_a, wv_a = save_and_reload(build_workbook(rows_same))
wf_b, wv_b = save_and_reload(build_workbook(rows_same_b))
results = check_display_text_cross(wv_a, wf_a, SheetLayout(wf_a), wv_b, wf_b, SheetLayout(wf_b))[0]
check('display_text 顯示一致的列 cause 為空', results and results[0]['cause'] == '', str(results))

# ------------------------------------------------------------------
# --md 報告檔名要帶分頁名稱：同一組檔案先核請款明細表、再核計價總表，不可互相覆蓋
# ------------------------------------------------------------------
p1 = md_report_path('D:/案件/1150301-第一期.xlsx', '請款明細表')
p2 = md_report_path('D:/案件/1150301-第一期.xlsx', '計價總表')
check('md 報告檔名含分頁名稱', p1.name == '1150301-第一期-請款明細表-核算報告.md', str(p1))
check('不同分頁的 md 報告檔名不同（不會互相覆蓋）', p1 != p2)
check('md 報告放在前一期檔案同目錄', p1.parent == Path('D:/案件'), str(p1.parent))

# ------------------------------------------------------------------
# SheetLayout 資料列掃描：可以停，但不可以「靜靜地」停
# ------------------------------------------------------------------
rows = [dict(code='壹.%d' % i, name='項目%d' % i, pa=1, ca=1, ta=2) for i in range(1, 451)]
wf, wv = save_and_reload(build_workbook(rows))
layout = SheetLayout(wf)
scanned = len(check_within_period(wv, layout)[0])
check('掃描沒有 400 列上限（450 列要全掃到）', scanned == 450, '實際只掃到 {0} 列'.format(scanned))
check('掃描停止原因有記錄', bool(layout.scan_stop_reason), str(layout.scan_stop_reason))
check('describe() 會印出掃描停止原因', '掃描停止原因' in layout.describe())

# 「本頁總計」是中途的分頁彙總列，不能當成資料區結尾
rows = [dict(code='壹.一', name='前段', pa=1, ca=1, ta=2),
        dict(code=None, name='本頁總計', pa=1, ca=1, ta=2),
        dict(code='貳.一', name='後段', pa=1, ca=1, ta=2)]
wf, wv = save_and_reload(build_workbook(rows))
scanned = len(check_within_period(wv, SheetLayout(wf))[0])
check('「本頁總計」不截斷資料區', scanned == 3, '只掃到 {0} 列'.format(scanned))

# 「總計」仍然視為資料區的最後一列
rows = [dict(code='壹.一', name='前段', pa=1, ca=1, ta=2),
        dict(code=None, name='總計', pa=1, ca=1, ta=2),
        dict(code=None, name='施工承攬廠商（編製）：')]
wf, wv = save_and_reload(build_workbook(rows))
scanned = len(check_within_period(wv, SheetLayout(wf))[0])
check('「總計」仍視為資料區最後一列', scanned == 2, '掃到 {0} 列'.format(scanned))

# ------------------------------------------------------------------
# 結論/exit code：有檢查被略過時，不能算成「全部相符」
# ------------------------------------------------------------------
bad, skipped = summarize({
    'extended_price': [], 'extended_price_skip': '缺少必要欄位（單價/數量/複價），略過此檢查',
    'within_period': [dict(qty_ok=True, amt_ok=True)], 'within_period_skip': None,
})
check('summarize 會回報被略過的檢查（exit code 才會是 2）',
      bad == 0 and len(skipped) == 1, str((bad, skipped)))

# ------------------------------------------------------------------
print()
if FAILURES:
    print('=== {0} 項測試失敗 ==='.format(len(FAILURES)))
    for f in FAILURES:
        print('  -', f)
    sys.exit(1)
else:
    print('=== 全部測試通過 ===')
    sys.exit(0)
