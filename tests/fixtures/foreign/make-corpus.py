#!/usr/bin/env python
"""
Corpus of workbooks that Excel did NOT write (round 5, module H1): openpyxl / XlsxWriter / pandas flavours.

    python make-corpus.py <outdir>          (needs: pip install openpyxl XlsxWriter pandas)

The node half (ExcelJS, SheetJS, xlsx-populate and the hand-made "Google Sheets / LibreOffice / POI / Open XML SDK /
Mac Excel 2011 / ERP converter / Strict OOXML ..." packages) is make-corpus.mjs. Every file is a student list in the
shape of the real ones: a header row, 20-40 data rows, an ID column (学籍番号 / 学生ID), 氏名, 大学, 点数, dates,
formulas, often a 合計 row, freeze panes, filters, merged titles ... The data is deterministic (no random), so a
regenerated corpus is byte-stable apart from the writers' own timestamps.

partner.xlsx (written here with XlsxWriter) is the "other" file of every pipeline run: 学籍番号 / 氏名 / 面接評価 /
評価日 / メモ for most of the roster plus a few people who are not in it (they become appended rows in a full join).
"""
import datetime as dt
import os
import struct
import sys
import zlib

import openpyxl
import pandas as pd
import xlsxwriter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'corpus')
os.makedirs(OUT, exist_ok=True)

SURN = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤', '吉田', '山田', '佐々木', '山口', '松本', '井上', '木村', '林', '斎藤', '清水']
GIVEN = ['太郎', '花子', '一郎', '美咲', '健太', '結衣', '大輔', '陽子', '翔', '愛', '拓也', '直美', '誠', '由美', '亮', '恵']
UNIV = ['慶應義塾大学', '早稲田大学', '東京大学', '一橋大学', '上智大学', '明治大学', '青山学院大学', '立教大学']
FAC = ['総合政策学部', '環境情報学部', '経済学部', '法学部', '商学部', '理工学部']


def person(i):
    return SURN[i % len(SURN)] + GIVEN[(i * 7 + 3) % len(GIVEN)]


def sid(i):
    return 'S%07d' % (2024000 + i)


def roster(n=30):
    rows = []
    for i in range(1, n + 1):
        rows.append(dict(
            id=sid(i), name=person(i), univ=UNIV[i % len(UNIV)], fac=FAC[(i * 5) % len(FAC)],
            score=50 + (i * 37) % 51, date=dt.date(2024, 4, 1) + dt.timedelta(days=(i * 11) % 200),
            memo='' if i % 4 else 'メモ%d' % i))
    return rows


R = roster(30)
HEAD = ['学籍番号', '氏名', '大学', '学部', '点数', '入学日', '備考']
thin = Side(style='thin', color='FF808080')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HFILL = PatternFill('solid', start_color='FFDDEBF7', end_color='FFDDEBF7')


def path(name):
    return os.path.join(OUT, name)


def fill_basic(ws, top=1, left=1, rows=R, with_formula=True, total=True, styled=True):
    """header at (top, left), data below; a per-row formula column and a SUM row two rows below the data"""
    heads = HEAD + (['点数x1.1'] if with_formula else [])
    for j, h in enumerate(heads):
        c = ws.cell(top, left + j, h)
        if styled:
            c.font = Font(bold=True, name='Meiryo UI', size=10)
            c.fill = HFILL
            c.border = BOX
            c.alignment = Alignment(horizontal='center')
    for k, r in enumerate(rows):
        row = top + 1 + k
        vals = [r['id'], r['name'], r['univ'], r['fac'], r['score'], r['date'], r['memo'] or None]
        for j, v in enumerate(vals):
            c = ws.cell(row, left + j, v)
            if styled:
                c.border = BOX
            if j == 5:
                c.number_format = 'yyyy/mm/dd'
            if j == 4:
                c.number_format = '#,##0'
        if with_formula:
            col = get_column_letter(left + 4)
            c = ws.cell(row, left + 7, f'={col}{row}*1.1')
            c.number_format = '0.0'
            if styled:
                c.border = BOX
    last = top + len(rows)
    if total:
        col = get_column_letter(left + 4)
        ws.cell(last + 2, left, '合計')
        ws.cell(last + 2, left + 4, f'=SUM({col}{top + 1}:{col}{last})')
        if with_formula:
            c2 = get_column_letter(left + 7)
            ws.cell(last + 2, left + 7, f'=SUM({c2}{top + 1}:{c2}{last})')
    return last


# ---------------------------------------------------------------------------------------------------- partner
def make_partner():
    wb = xlsxwriter.Workbook(path('partner.xlsx'))
    ws = wb.add_worksheet('面接評価')
    hdr = wb.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1})
    dfmt = wb.add_format({'num_format': 'yyyy/m/d'})
    for j, h in enumerate(['学籍番号', '氏名', '面接評価', '評価日', 'メモ']):
        ws.write(0, j, h, hdr)
    rows = []
    for i in range(3, 33, 1):            # roster 3..30 (1 and 2 are missing) + 31, 32 which the roster does not have
        rows.append(i)
    for k, i in enumerate(rows):
        ws.write(k + 1, 0, sid(i))
        ws.write(k + 1, 1, person(i))
        ws.write(k + 1, 2, 'ABC'[i % 3])
        ws.write_datetime(k + 1, 3, dt.datetime(2024, 9, 1) + dt.timedelta(days=i), dfmt)
        ws.write(k + 1, 4, '良好' if i % 2 else None)
    ws.set_column(0, 1, 14)
    wb.close()


# ---------------------------------------------------------------------------------------------------- openpyxl
def openpyxl_files():
    # 1. styled list: freeze, filter, CF, DV, widths, formulas, total row
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    last = fill_basic(ws)
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = f'A1:H{last}'
    ws.conditional_formatting.add(f'E2:E{last}', CellIsRule(operator='lessThan', formula=['60'], fill=PatternFill('solid', bgColor='FFFFC7CE')))
    ws.conditional_formatting.add(f'H2:H{last}', ColorScaleRule(start_type='min', start_color='FFF8696B', end_type='max', end_color='FF63BE7B'))
    ws.conditional_formatting.add(f'A2:H{last}', FormulaRule(formula=['$E2>=95'], fill=PatternFill('solid', bgColor='FFC6EFCE')))
    dv = DataValidation(type='list', formula1='"慶應義塾大学,早稲田大学,東京大学"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'C2:C{last}')
    for col, w in zip('ABCDEFGH', [12, 14, 18, 16, 8, 12, 12, 10]):
        ws.column_dimensions[col].width = w
    wb.save(path('py-01-openpyxl-styled.xlsx'))

    # 2. title rows + merged title, header at B4 (odd start position), blank row + totals
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '成績一覧'
    ws['B1'] = '2024年度 前期 成績一覧'
    ws.merge_cells('B1:I1')
    ws['B1'].font = Font(bold=True, size=14)
    ws['B1'].alignment = Alignment(horizontal='center')
    ws['B2'] = '作成日: 2024/10/01'
    fill_basic(ws, top=4, left=2)
    wb.save(path('py-02-openpyxl-title-odd-origin.xlsx'))

    # 3. two-row header: group row (merged) over the real header row
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    ws['A1'] = '基本情報'
    ws.merge_cells('A1:D1')
    ws['E1'] = '成績'
    ws.merge_cells('E1:F1')
    ws['G1'] = 'その他'
    for j, h in enumerate(HEAD):
        c = ws.cell(2, j + 1, h)
        c.font = Font(bold=True)
    for k, r in enumerate(R):
        for j, v in enumerate([r['id'], r['name'], r['univ'], r['fac'], r['score'], r['date'], r['memo'] or None]):
            c = ws.cell(3 + k, j + 1, v)
            if j == 5:
                c.number_format = 'yyyy/m/d'
    wb.save(path('py-03-openpyxl-two-row-header.xlsx'))

    # 4. a real Table (ListObject) with a calculated column in structured-reference form
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    for j, h in enumerate(HEAD + ['加点']):
        ws.cell(1, j + 1, h)
    for k, r in enumerate(R):
        for j, v in enumerate([r['id'], r['name'], r['univ'], r['fac'], r['score'], r['date'], r['memo'] or None]):
            c = ws.cell(2 + k, j + 1, v)
            if j == 5:
                c.number_format = 'yyyy/mm/dd'
        ws.cell(2 + k, 8, f'=tblRoster[[#This Row],[点数]]+5')
    t = Table(displayName='tblRoster', ref=f'A1:H{1 + len(R)}')
    t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(t)
    ws2 = wb.create_sheet('集計')
    ws2['A1'] = '平均点'
    ws2['B1'] = '=AVERAGE(tblRoster[点数])'
    ws2['A2'] = '人数'
    ws2['B2'] = '=COUNTA(tblRoster[学籍番号])'
    wb.save(path('py-04-openpyxl-table.xlsx'))

    # 5. hidden sheet (first), hidden rows / columns, grouped columns, row heights
    wb = openpyxl.Workbook()
    ms = wb.active
    ms.title = 'マスタ'
    ms['A1'] = '大学'
    for i, u in enumerate(UNIV):
        ms.cell(2 + i, 1, u)
    ms.sheet_state = 'hidden'
    ws = wb.create_sheet('名簿')
    wb.active = 1
    last = fill_basic(ws)
    ws.column_dimensions['D'].hidden = True
    ws.column_dimensions.group('F', 'G', hidden=False, outline_level=1)
    ws.row_dimensions[5].hidden = True
    ws.row_dimensions[7].hidden = True
    ws.row_dimensions[1].height = 30
    wb.save(path('py-05-openpyxl-hidden.xlsx'))

    # 6. 1904 date system
    wb = openpyxl.Workbook()
    wb.epoch = CALENDAR_MAC_1904
    ws = wb.active
    ws.title = '名簿'
    fill_basic(ws, total=False)
    wb.save(path('py-06-openpyxl-date1904.xlsx'))

    # 7. full-width IDs / names with full-width spaces and digits
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    for j, h in enumerate(['学籍番号', '氏　名', '点数', 'ＴＥＬ']):
        ws.cell(1, j + 1, h)
    for k, r in enumerate(R):
        ws.cell(2 + k, 1, sid(k + 1).translate({ord(str(d)): ord('０') + d for d in range(10)}).replace('S', 'Ｓ'))
        ws.cell(2 + k, 2, r['name'][:2] + '　' + r['name'][2:])
        ws.cell(2 + k, 3, '{:d}'.format(r['score']).translate({ord(str(d)): ord('０') + d for d in range(10)}))
        ws.cell(2 + k, 4, '０３－１２３４－%04d' % (5000 + k))
    wb.save(path('py-07-openpyxl-fullwidth.xlsx'))

    # 8. very long text + rich text + line breaks
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    for j, h in enumerate(['学籍番号', '氏名', '自己PR', '備考']):
        ws.cell(1, j + 1, h)
    for k, r in enumerate(R[:12]):
        ws.cell(2 + k, 1, r['id'])
        ws.cell(2 + k, 2, r['name'])
    ws.cell(2, 3, 'あ' * 32000)
    ws.cell(3, 3, ('とても長い自己PRの文章です。' * 400)[:9000])
    ws.cell(4, 3, '1行目\n2行目\n3行目').alignment = Alignment(wrap_text=True)
    ws.cell(5, 4, CellRichText('通常 ', TextBlock(InlineFont(b=True, color='FFFF0000'), '太字赤'), ' と 普通'))
    ws.cell(6, 4, 'a&b <c> "d" \'e\'')
    ws.cell(7, 4, '_x0041_ literal escape')
    wb.save(path('py-08-openpyxl-longtext-rich.xlsx'))

    # 9. thousands of columns
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    ws.cell(1, 1, '学籍番号')
    ws.cell(1, 2, '氏名')
    for j in range(3, 2501):
        ws.cell(1, j, f'項目{j - 2}')
    for k in range(20):
        ws.cell(2 + k, 1, sid(k + 1))
        ws.cell(2 + k, 2, person(k + 1))
        for j in range(3, 2501):
            ws.cell(2 + k, j, (k * j) % 97)
    ws.column_dimensions['C'].width = 20
    wb.save(path('py-09-openpyxl-2500-columns.xlsx'))

    # 10. ISO dates (t="d" cells)
    wb = openpyxl.Workbook(iso_dates=True)
    ws = wb.active
    ws.title = '名簿'
    fill_basic(ws, total=False)
    wb.save(path('py-10-openpyxl-iso-dates.xlsx'))

    # 11. charts on another sheet + defined names + hyperlinks + comments + DV from a named range
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    last = fill_basic(ws)
    ws['J1'] = 'リンク'
    ws['J2'] = 'HP'
    ws['J2'].hyperlink = 'https://example.com/a?b=1&c=2'
    ws['J3'] = '内部'
    ws['J3'].hyperlink = "#'名簿'!A1"
    ws['B2'].comment = Comment('確認済み', '担当')
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names['点数列'] = DefinedName('点数列', attr_text=f"'名簿'!$E$2:$E${last}")
    wb.defined_names['氏名列'] = DefinedName('氏名列', attr_text=f"'名簿'!$B$2:$B${last}")
    dv = DataValidation(type='list', formula1='=氏名列', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add('K2:K10')
    cs = wb.create_sheet('グラフ')
    ch = BarChart()
    ch.title = '点数'
    ch.add_data(Reference(ws, min_col=5, min_row=1, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=2, max_row=last))
    cs.add_chart(ch, 'B2')
    lc = LineChart()
    lc.add_data(Reference(ws, min_col=8, min_row=1, max_row=last), titles_from_data=True)
    cs.add_chart(lc, 'B20')
    cs['A1'] = '=SUM(点数列)'
    cs['A2'] = "=COUNTA(名簿!B2:B31)"
    wb.save(path('py-11-openpyxl-chart-names-links.xlsx'))

    # 12. sheet + workbook protection
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    fill_basic(ws)
    ws.protection.sheet = True
    ws.protection.password = 'x'
    ws.protection.formatColumns = False
    wb.security.lockStructure = True
    wb.save(path('py-12-openpyxl-protected.xlsx'))

    # 13. data in the middle of a sheet with blank rows inside + stray formatted cells far away
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    fill_basic(ws, top=3, left=1, rows=R[:10], total=False)
    for k, r in enumerate(R[10:20]):
        row = 3 + 1 + 10 + 2 + k      # two blank rows, then a second block
        ws.cell(row, 1, r['id'])
        ws.cell(row, 2, r['name'])
        ws.cell(row, 5, r['score'])
    ws['Z500'].fill = HFILL
    ws.row_dimensions[900].height = 40
    wb.save(path('py-13-openpyxl-blank-rows-stray.xlsx'))

    # 14. numbers stored as text, leading zeros, errors, booleans, percent / currency / time formats
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    for j, h in enumerate(['社員番号', '氏名', '入社年月', '達成率', '給与', '勤務開始', '在籍', '備考']):
        ws.cell(1, j + 1, h)
    for k in range(25):
        ws.cell(2 + k, 1, '%06d' % (100 + k))
        ws.cell(2 + k, 2, person(k + 1))
        ws.cell(2 + k, 3, dt.date(2020 + k % 5, 1 + k % 12, 1)).number_format = 'yyyy"年"m"月"'
        ws.cell(2 + k, 4, 0.5 + k / 100).number_format = '0.0%'
        ws.cell(2 + k, 5, 250000 + k * 1000).number_format = '"¥"#,##0'
        ws.cell(2 + k, 6, dt.time(9, k % 60)).number_format = 'h:mm'
        ws.cell(2 + k, 7, k % 2 == 0)
    ws['H2'] = '=1/0'
    ws['H3'] = '=NA()'
    ws['H4'] = '=IFERROR(1/0,"n/a")'
    wb.save(path('py-14-openpyxl-formats-errors.xlsx'))


    # 15a. a chart (default-namespace drawing / chart XML) anchored ON the data sheet next to the list, plus two notes (VML with ns0: prefixes)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    last = fill_basic(ws, total=False)
    ch = BarChart()
    ch.title = '点数'
    ch.add_data(Reference(ws, min_col=5, min_row=1, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=2, max_row=last))
    ws.add_chart(ch, 'K2')
    ws['C3'].comment = Comment('要確認', '担当')
    ws['G6'].comment = Comment('電話済み', '担当')
    wb.save(path('py-27-openpyxl-chart-on-data-sheet.xlsx'))

    # 15. as many columns as a sheet can have (XFD): importing anything next to it cannot fit
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    ws.cell(1, 1, '学籍番号')
    ws.cell(1, 2, '氏名')
    for j in range(3, 16385):
        ws.cell(1, j, f'項目{j - 2}')
    for k in range(5):
        ws.cell(2 + k, 1, sid(k + 1))
        ws.cell(2 + k, 2, person(k + 1))
        ws.cell(2 + k, 3, k)
        ws.cell(2 + k, 16384, k)
    wb.save(path('py-26-openpyxl-xfd-columns.xlsx'))


# ---------------------------------------------------------------------------------------------------- XlsxWriter
def xlsxwriter_files():
    wb = xlsxwriter.Workbook(path('py-15-xlsxwriter-basic.xlsx'))
    ws = wb.add_worksheet('名簿')
    h = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'font_name': 'ＭＳ Ｐゴシック'})
    d = wb.add_format({'num_format': 'yyyy/m/d', 'border': 1})
    n = wb.add_format({'num_format': '#,##0', 'border': 1})
    tx = wb.add_format({'border': 1})
    ws.merge_range('A1:H1', '学生名簿（XlsxWriter）', wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'}))
    for j, t in enumerate(HEAD + ['加点']):
        ws.write(2, j, t, h)
    for k, r in enumerate(R):
        row = 3 + k
        ws.write(row, 0, r['id'], tx)
        ws.write(row, 1, r['name'], tx)
        ws.write(row, 2, r['univ'], tx)
        ws.write(row, 3, r['fac'], tx)
        ws.write_number(row, 4, r['score'], n)
        ws.write_datetime(row, 5, dt.datetime.combine(r['date'], dt.time()), d)
        ws.write(row, 6, r['memo'], tx)
        ws.write_formula(row, 7, f'=E{row + 1}+5', n, r['score'] + 5)
    last = 3 + len(R)
    ws.write(last + 1, 0, '合計')
    ws.write_formula(last + 1, 4, f'=SUM(E4:E{last})', n, sum(r['score'] for r in R))
    ws.freeze_panes(3, 2)
    ws.autofilter(2, 0, last - 1, 7)
    ws.conditional_format(f'E4:E{last}', {'type': '3_color_scale'})
    ws.conditional_format(f'H4:H{last}', {'type': 'cell', 'criteria': '>', 'value': 90, 'format': wb.add_format({'bg_color': '#FFC7CE'})})
    ws.data_validation(f'C4:C{last}', {'validate': 'list', 'source': UNIV})
    ws.write_comment('B4', '確認済み')
    ws.write_comment('F8', '日付を確認')
    ws.set_column('A:A', 12)
    ws.set_column('B:D', 16)
    ws.set_column('E:H', 10)
    ws.set_row(2, 24)
    wb.define_name('点数', f"='名簿'!$E$4:$E${last}")
    ws2 = wb.add_worksheet('集計')
    ws2.write_formula('A1', '=SUM(点数)', None, sum(r['score'] for r in R))
    ws2.write_formula('A2', f"=AVERAGE('名簿'!E4:E{last})", None, 0)
    wb.close()

    # constant_memory: rows are flushed as they are written and every string is an INLINE string
    wb = xlsxwriter.Workbook(path('py-16-xlsxwriter-constant-memory.xlsx'), {'constant_memory': True})
    ws = wb.add_worksheet('名簿')
    for j, t in enumerate(HEAD):
        ws.write(0, j, t)
    for k, r in enumerate(R):
        ws.write(1 + k, 0, r['id'])
        ws.write(1 + k, 1, r['name'])
        ws.write(1 + k, 2, r['univ'])
        ws.write(1 + k, 3, r['fac'])
        ws.write_number(1 + k, 4, r['score'])
        ws.write_datetime(1 + k, 5, dt.datetime.combine(r['date'], dt.time()), wb.add_format({'num_format': 'yyyy-mm-dd'}))
        ws.write(1 + k, 6, r['memo'])
    wb.close()

    # dynamic / legacy array formulas, _xlfn functions (+ one written WITHOUT the prefix -> #NAME? in Excel)
    wb = xlsxwriter.Workbook(path('py-17-xlsxwriter-arrays.xlsx'))
    ws = wb.add_worksheet('名簿')
    for j, t in enumerate(['学籍番号', '氏名', '点数']):
        ws.write(0, j, t)
    for k, r in enumerate(R):
        ws.write(1 + k, 0, r['id'])
        ws.write(1 + k, 1, r['name'])
        ws.write_number(1 + k, 2, r['score'])
    ws.write_dynamic_array_formula('E2:E4', '=SORT(C2:C31,1,-1)', None, 0)
    ws.write_array_formula('G2:G4', '{=C2:C4*2}', None, 0)
    ws.write_formula('I2', '=_xlfn.STDEV.S(C2:C31)', None, 0)
    ws.write_formula('I3', '=_xlfn.IFS(C2>90,"A",C2>70,"B",TRUE,"C")', None, '')
    ws.write_formula('I4', '=STDEV.S(C2:C31)', None, 0)   # no prefix: Excel shows #NAME? until re-entered
    ws.write_formula('J2', '=TEXT(C2,"0.0")&"点"', None, '')
    wb.close()

    # chart + image
    png = b'\x89PNG\r\n\x1a\n' + _chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0)) + _chunk(b'IDAT', zlib.compress(b'\x00\xff\x00\x00\xff\x00\x00\x00\x00\xff\x00\xff\x00\xff\xff\x00')) + _chunk(b'IEND', b'')
    with open(os.path.join(OUT, '_pixel.png'), 'wb') as f:
        f.write(png)
    wb = xlsxwriter.Workbook(path('py-18-xlsxwriter-chart-image.xlsx'))
    ws = wb.add_worksheet('名簿')
    for j, t in enumerate(['学籍番号', '氏名', '点数']):
        ws.write(0, j, t)
    for k, r in enumerate(R):
        ws.write(1 + k, 0, r['id'])
        ws.write(1 + k, 1, r['name'])
        ws.write_number(1 + k, 2, r['score'])
    ch = wb.add_chart({'type': 'column'})
    ch.add_series({'name': '点数', 'categories': "='名簿'!$B$2:$B$31", 'values': "='名簿'!$C$2:$C$31"})
    ws.insert_chart('F2', ch)
    ws.insert_image('F20', os.path.join(OUT, '_pixel.png'))
    wb.close()
    os.remove(os.path.join(OUT, '_pixel.png'))

    # tables from XlsxWriter (column formulas, totals row)
    wb = xlsxwriter.Workbook(path('py-19-xlsxwriter-table.xlsx'))
    ws = wb.add_worksheet('名簿')
    data = [[r['id'], r['name'], r['univ'], r['score']] for r in R]
    ws.add_table(2, 1, 2 + len(data) + 1, 5, {
        'data': [d + [None] for d in data], 'name': 'Roster', 'total_row': True, 'style': 'Table Style Light 9',
        'columns': [{'header': '学籍番号', 'total_string': '合計'}, {'header': '氏名'}, {'header': '大学'},
                    {'header': '点数', 'total_function': 'sum'},
                    {'header': '加点', 'formula': '=[@点数]+10', 'total_function': 'average'}]})
    ws.write(0, 1, '名簿（テーブル）')
    wb.close()


def _chunk(tag, data):
    return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff)


# ---------------------------------------------------------------------------------------------------- pandas
def pandas_files():
    df = pd.DataFrame([{ '学籍番号': r['id'], '氏名': r['name'], '大学': r['univ'], '点数': r['score'], '入学日': pd.Timestamp(r['date']), '備考': r['memo'] or None } for r in R])
    df.to_excel(path('py-20-pandas-default-index.xlsx'), sheet_name='名簿')   # index column A, styled header
    df.to_excel(path('py-21-pandas-noindex-xlsxwriter.xlsx'), sheet_name='名簿', index=False, engine='xlsxwriter')
    # MultiIndex columns -> two header rows with merged cells + an index column
    cols = pd.MultiIndex.from_tuples([('基本', '学籍番号'), ('基本', '氏名'), ('成績', '前期'), ('成績', '後期'), ('成績', '合計')])
    m = pd.DataFrame([[r['id'], r['name'], r['score'], r['score'] - 3, 2 * r['score'] - 3] for r in R], columns=cols)
    m.to_excel(path('py-22-pandas-multiindex-header.xlsx'), sheet_name='成績', engine='openpyxl')
    m.to_excel(path('py-23-pandas-multiindex-xlsxwriter.xlsx'), sheet_name='成績', engine='xlsxwriter')
    # datetimes with time + NaN + timezone-free + mixed types
    t = pd.DataFrame({'社員番号': ['A%03d' % i for i in range(25)], '氏名': [person(i) for i in range(25)],
                      '打刻': pd.date_range('2024-04-01 08:30', periods=25, freq='37min'),
                      '金額': [None if i % 6 == 0 else i * 1234.5 for i in range(25)],
                      '区分': pd.Categorical(['a', 'b', 'c', 'a', 'b'] * 5)})
    t.to_excel(path('py-24-pandas-datetime-nan.xlsx'), sheet_name='勤怠', index=False)
    # one-sheet-per-group workbook (ExcelWriter with several sheets, sheet names with spaces / special chars)
    with pd.ExcelWriter(path('py-25-pandas-many-sheets.xlsx')) as w:
        df.to_excel(w, sheet_name='全員', index=False)
        for u, g in df.groupby('大学'):
            g.to_excel(w, sheet_name=u[:10], index=False)
        df.head(3).to_excel(w, sheet_name="O'Brien & 名簿", index=False)


def streaming_and_styler():
    """openpyxl write_only (streamed rows, no dimension) and pandas Styler.to_excel (per-cell styles: gradient fills, bold, borders)"""
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet('名簿')
    ws.freeze_panes = 'A2'
    ws.append(HEAD)
    for r in R:
        ws.append([r['id'], r['name'], r['univ'], r['fac'], r['score'], r['date'], r['memo'] or None])
    wb.save(path('py-30-openpyxl-write-only.xlsx'))
    df = pd.DataFrame([{'学籍番号': r['id'], '氏名': r['name'], '点数': r['score'], '入学日': pd.Timestamp(r['date'])} for r in R])
    styler = df.style.apply(lambda col: ['background-color: #ffeeaa' if (isinstance(v, int) and v >= 90) else '' for v in col], subset=['点数']).format({'入学日': '{:%Y/%m/%d}'})
    styler.to_excel(path('py-31-pandas-styler.xlsx'), sheet_name='名簿', index=False, engine='openpyxl')


def other_writers():
    """pyexcelerate and pylightxl: two more pure-python writers (their own sheet XML: inline strings / shared strings, no theme)"""
    try:
        from pyexcelerate import Workbook as PWorkbook
        wb = PWorkbook()
        data = [HEAD] + [[r['id'], r['name'], r['univ'], r['fac'], r['score'], r['date'], r['memo']] for r in R]
        ws = wb.new_sheet('名簿', data=data)
        ws.set_col_style(6, __import__('pyexcelerate').Style(format=__import__('pyexcelerate').Format('yyyy/mm/dd')))
        wb.save(path('py-28-pyexcelerate.xlsx'))
    except Exception as e:  # pragma: no cover - optional dependency
        print('pyexcelerate skipped:', e)
    try:
        import pylightxl as xl
        db = xl.Database()
        db.add_ws('名簿')
        ws = db.ws('名簿')
        for j, h in enumerate(HEAD):
            ws.update_index(1, j + 1, h)
        for i, r in enumerate(R):
            for j, v in enumerate([r['id'], r['name'], r['univ'], r['fac'], r['score'], r['date'].isoformat(), r['memo']]):
                ws.update_index(i + 2, j + 1, v)
        ws.update_address('I2', '=E2*2')
        xl.writexl(db, path('py-29-pylightxl.xlsx'))
    except Exception as e:  # pragma: no cover - optional dependency
        print('pylightxl skipped:', e)


if __name__ == '__main__':
    streaming_and_styler()
    other_writers()
    make_partner()
    openpyxl_files()
    xlsxwriter_files()
    pandas_files()
    print('written:', len([f for f in os.listdir(OUT) if f.endswith('.xlsx')]), 'files ->', OUT)
