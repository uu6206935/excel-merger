# -*- coding: utf-8 -*-
"""
Corpus generator for the formula-import findings / oracle (task F1).

    python gen_corpus.py <outDir>

Writes <outDir>/nocache/*.xlsx   workbooks written by openpyxl / XlsxWriter: formulas WITHOUT cached results
       <outDir>/base.xlsx        the base workbook (sheet 名簿, 30 students)
       <outDir>/manifest.json    one entry per source file (what the COM step and the Node tools need)

The COM step (author.ps1) opens every nocache file in a hidden Excel, applies the `comPost` steps (shared formulas
need Excel), recalculates and saves <outDir>/cached/<same name>.xlsx (Excel-written cached results, shared formulas ...).

Persons: base S001..S030; every source holds S001..S027 (S028..S030 are base-only) + S101..S104 (source-only, appended
with join 'full'). Orders: same (ID order, extras last), rev (reversed), shuf (seeded shuffle).
Source layout (sheet 評価, header row 1, data rows 2..32): A 学生ID, B 氏名, C 評価1, D 評価2, E 評価3, F.. formulas.
"""
import json
import os
import random
import sys

import openpyxl
import xlsxwriter
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

SURNAMES = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤']
GIVEN = ['太郎', '花子', '一郎', '美咲', '健太', '結衣', '大輔', '陽菜', '翔太', '愛']
UNIS = ['東京大学', '京都大学', '慶應義塾大学', '早稲田大学', '大阪大学']
SHEET = '評価'
ORDERS = ['same', 'rev', 'shuf']
BASE_IDS = ['S%03d' % i for i in range(1, 31)]
SRC_IDS = ['S%03d' % i for i in range(1, 28)] + ['S%03d' % i for i in range(101, 105)]
LAST = 32  # last data row of a standard source (31 persons, header row 1)


def person_name(sid):
    n = int(sid[1:])
    return SURNAMES[n % 10] + GIVEN[(n // 10 + n * 3) % 10] + ('' if n < 100 else str(n % 100))


def person_uni(sid):
    return UNIS[int(sid[1:]) % 5]


ALL_IDS = BASE_IDS + ['S%03d' % i for i in range(101, 105)]
# every score column holds DIFFERENT values (a lookup / MATCH by score must find one person: with ties the answer would depend on the
# row order and a reordered merge could not be compared with the source)
_PERM = [random.Random(1000 + k).sample(range(45, 100), len(ALL_IDS)) for k in range(3)]


def score(sid, k):
    return _PERM[k][ALL_IDS.index(sid)]


def ordered(order):
    ids = list(SRC_IDS)
    if order == 'rev':
        ids.reverse()
    elif order == 'shuf':
        random.Random(42).shuffle(ids)
    return ids


def write_base(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '名簿'
    ws.append(['学生ID', '氏名', '大学', '学年'])
    for sid in BASE_IDS:
        ws.append([sid, person_name(sid), person_uni(sid), 1 + int(sid[1:]) % 4])
    for c in ws[1]:
        c.font = Font(bold=True)
    wb.save(path)


def std_rows(ws, order, header_row=1, extra_cols=0):
    """header + data (A..E) of the standard source; returns the list of person ids in sheet order"""
    heads = ['学生ID', '氏名', '評価1', '評価2', '評価3']
    for i, h in enumerate(heads):
        ws.cell(row=header_row, column=i + 1, value=h).font = Font(bold=True)
    ids = ordered(order)
    for k, sid in enumerate(ids):
        r = header_row + 1 + k
        ws.cell(row=r, column=1, value=sid)
        ws.cell(row=r, column=2, value=person_name(sid))
        for j in range(3):
            ws.cell(row=r, column=3 + j, value=score(sid, j))
    return ids


def put_formulas(ws, cols, first=2, last=LAST):
    """cols: list of (header, template) ; template uses {r} {p} (previous row) {n} (next row); a callable gets r"""
    for i, (h, t) in enumerate(cols):
        c = 6 + i
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
        for r in range(first, last + 1):
            f = t(r) if callable(t) else t.format(r=r, p=r - 1, n=r + 1)
            if f is None:
                continue
            ws.cell(row=r, column=c, value=f)


def item_samerow(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('合計', '=C{r}+D{r}+E{r}'),
        ('平均', '=AVERAGE(C{r}:E{r})'),
        ('判定', '=IF(F{r}>=210,"合格","要再試験")'),
        ('丸め', '=ROUND(G{r},1)'),
        ('空返し', '=IF(F{r}>240,F{r},"")'),
        ('日付', '=DATE(2026,4,1)+E{r}'),
        ('文字', '=TEXT(G{r},"0.0")&"点"'),
        ('エラー', '=IF(MOD(C{r},7)=0,NA(),C{r})'),
    ])
    for r in range(2, LAST + 1):
        ws.cell(row=r, column=11).number_format = 'yyyy/m/d'


def item_shared(wb, order):
    ws = wb.active
    std_rows(ws, order)
    for i, h in enumerate(['合計', '倍', '判定']):
        ws.cell(row=1, column=6 + i, value=h).font = Font(bold=True)
    # the formulas are assigned to whole ranges by Excel (comPost) so that they are stored as shared formulas


def item_keycol(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('様', '=A{r}&"様"'),
        ('重複', '=IF(COUNTIF($A$2:$A$32,A{r})>1,"重複","")'),
        ('氏名評価', '=B{r}&"（"&C{r}&"）"'),
        ('IDで引く', '=VLOOKUP(A{r},$A$2:$E$32,3,FALSE)'),
        ('姓', '=LEFT(B{r},1)'),
        ('ID長', '=LEN(A{r})'),
        ('氏名で引く', '=INDEX($D$2:$D$32,MATCH(B{r},$B$2:$B$32,0))'),
    ])


def item_lookups(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('INDEXMATCH', '=INDEX($D$2:$D$32,MATCH(A{r},$A$2:$A$32,0))'),
        ('SUMIF', '=SUMIF($C$2:$C$32,">"&C{r},$D$2:$D$32)'),
        ('COUNTIF', '=COUNTIF($C$2:$C$32,">="&C{r})'),
        ('RANK', '=RANK(C{r},$C$2:$C$32)'),
        ('LARGE3', '=LARGE($C$2:$C$32,3)'),
        ('XLOOKUP', '=_xlfn.XLOOKUP(C{r},$C$2:$C$32,$D$2:$D$32,"-")'),
        ('SUMPRODUCT', '=SUMPRODUCT(($C$2:$C$32>C{r})*1)'),
        ('MATCH位置', '=MATCH(C{r},$C$2:$C$32,0)'),
        ('VLOOKUP2', '=VLOOKUP(C{r},$C$2:$E$32,2,FALSE)'),
        ('同ID数', '=COUNTIFS($C$2:$C$32,C{r},$D$2:$D$32,D{r})'),
    ])


def item_ranges(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('合計全体', '=SUM(C$2:C$32)'),
        ('平均D', '=AVERAGE(D2:D32)'),
        ('平均D見出し込み', '=AVERAGE(D1:D32)'),
        ('最大E列', '=MAX(E:E)'),
        ('平均D1000', '=AVERAGE(D2:D1000)'),
        ('件数', '=COUNT(C2:C100)'),
        ('合計1行少ない', '=SUM(C2:C31)'),
        ('合計上1行少ない', '=SUM(C3:C32)'),
        ('絶対行', '=C{r}-C$2'),
        ('列全体', '=C{r}/SUM(C:C)'),
    ])


def item_running(wb, order):
    ws = wb.active
    std_rows(ws, order)
    ws['Q1'] = 'レート'
    ws['P1'] = 1.1
    put_formulas(ws, [
        ('累計', '=SUM(C$2:C{r})'),
        ('前行差', lambda r: '=C2' if r == 2 else '=C%d-C%d' % (r, r - 1)),
        ('構成比', '=C{r}/SUM(C$2:C$32)'),
        ('レート', '=C{r}*$P$1'),
        ('偏差', '=C{r}-AVERAGE(C$2:C$32)'),
        ('行番号', '=ROW()-1'),
        ('行数', '=ROWS($C$2:$C$32)'),
        ('連番', '=COUNTA($A$2:A{r})'),
        ('次行', lambda r: '=C%d' % (r + 1)),
    ])


def item_othersheet(wb, order):
    ws = wb.active
    std_rows(ws, order)
    m = wb.create_sheet('Master')
    m.append(['学生ID', '区分'])
    for sid in SRC_IDS + BASE_IDS[27:]:
        m.append([sid, '区分' + 'ABC'[int(sid[1:]) % 3]])
    s2 = wb.create_sheet('Sheet2')
    s2['B3'] = 1.5
    put_formulas(ws, [
        ('区分', '=VLOOKUP(A{r},Master!$A$2:$B$100,2,FALSE)'),
        ('区分2', '=INDEX(Master!B:B,MATCH(A{r},Master!A:A,0))'),
        ('係数', '=Sheet2!$B$3*C{r}'),
        ('自シート修飾', "='評価'!C{r}*2"),
        ('IFERROR', '=IFERROR(VLOOKUP(A{r},Master!$A:$B,2,0),"未登録")'),
        ('自シート範囲', "=SUM('評価'!C$2:C$32)"),
    ])


def item_names(wb, order):
    ws = wb.active
    std_rows(ws, order)
    m = wb.create_sheet('Master')
    m.append(['学生ID', '区分'])
    for sid in SRC_IDS:
        m.append([sid, '区分' + 'ABC'[int(sid[1:]) % 3]])
    ws['P1'] = 2
    wb.defined_names['税率'] = DefinedName('税率', attr_text='0.08')
    wb.defined_names['評価列'] = DefinedName('評価列', attr_text="'評価'!$C$2:$C$32")
    wb.defined_names['係数'] = DefinedName('係数', attr_text="'評価'!$P$1")
    wb.defined_names['区分表'] = DefinedName('区分表', attr_text='Master!$A$2:$B$100')
    wb.defined_names['動的'] = DefinedName('動的', attr_text="OFFSET('評価'!$C$1,1,0,31,1)")
    wb.defined_names['ラベル'] = DefinedName('ラベル', attr_text='"点"')
    ws.defined_names['シート内'] = DefinedName('シート内', attr_text="'評価'!$D$2:$D$32")
    put_formulas(ws, [
        ('税込', '=C{r}*税率'),
        ('列合計', '=SUM(評価列)'),
        ('係数掛け', '=C{r}*係数'),
        ('区分表', '=VLOOKUP(A{r},区分表,2,FALSE)'),
        ('動的合計', '=SUM(動的)'),
        ('構成比', '=C{r}/SUM(評価列)'),
        ('順位', '=RANK(C{r},評価列)'),
        ('ラベル付', '=C{r}&ラベル'),
        ('シート内平均', '=AVERAGE(シート内)'),
    ])


def item_tables(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('合計', '=Table1[[#This Row],[評価1]]+Table1[[#This Row],[評価2]]'),
        ('E合計', '=SUM(Table1[評価3])'),
        ('倍', '=Table1[[#This Row],[評価1]]*2'),
        ('順位', '=RANK(Table1[[#This Row],[評価1]],Table1[評価1])'),
        ('人数', '=COUNTA(Table1[[#Data],[氏名]])'),
        ('引当', '=INDEX(Table1[評価2],MATCH(Table1[[#This Row],[学生ID]],Table1[学生ID],0))'),
        ('見出し', '=Table1[[#Headers],[評価1]]'),
        ('行数', '=ROWS(Table1[])'),
        ('範囲和', '=SUM(Table1[[#This Row],[評価1]:[評価3]])'),
    ])
    t = Table(displayName='Table1', ref='A1:N%d' % LAST)
    t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(t)


def item_values(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('優可', '=IF(C{r}>=80,"優","可")'),
        ('日付', '=DATE(2026,4,1)+C{r}'),
        ('時刻', '=TIME(MOD(C{r},24),0,0)'),
        ('真偽', '=C{r}>=D{r}'),
        ('ゼロ割', '=C{r}/(D{r}-D{r})'),
        ('NA', '=NA()'),
        ('空文字', '=""'),
    ])
    for r in range(2, LAST + 1):
        ws.cell(row=r, column=7).number_format = 'yyyy/m/d'
        ws.cell(row=r, column=8).number_format = 'h:mm'


def item_footer(wb, order):
    ws = wb.active
    std_rows(ws, order)
    ws.cell(row=34, column=2, value='合計')
    ws.cell(row=35, column=2, value='平均')
    for c in 'CDE':
        ws['%s34' % c] = '=SUM(%s2:%s32)' % (c, c)
        ws['%s35' % c] = '=AVERAGE(%s2:%s32)' % (c, c)
    put_formulas(ws, [
        ('合計', '=C{r}+D{r}'),
        ('構成比', '=C{r}/$C$34'),
        ('平均との差', '=C{r}-$C$35'),
        ('範囲和', '=SUM(C$2:C$32)'),
    ])
    ws['F34'] = '=SUM(F2:F32)'


def item_hyperlink(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('リンク', '=HYPERLINK("https://example.com/"&A{r},B{r})'),
        ('内部リンク', '=HYPERLINK("#\'評価\'!C"&ROW(),"移動")'),
        ('CONCAT', '=_xlfn.CONCAT(B{r},"-",C{r})'),
    ])


def item_twoheader(wb, order):
    ws = wb.active
    ids = ordered(order)
    ws.cell(row=1, column=3, value='面接')
    ws.cell(row=1, column=5, value='筆記')
    heads = ['学生ID', '氏名', '1次', '2次', '点数']
    for i, h in enumerate(heads):
        ws.cell(row=2, column=i + 1, value=h).font = Font(bold=True)
    for k, sid in enumerate(ids):
        r = 3 + k
        ws.cell(row=r, column=1, value=sid)
        ws.cell(row=r, column=2, value=person_name(sid))
        for j in range(3):
            ws.cell(row=r, column=3 + j, value=score(sid, j))
    last = 2 + len(ids)
    for i, (h, t) in enumerate([('合計', '=C{r}+D{r}+E{r}'), ('平均', '=AVERAGE(C$3:C$%d)' % last), ('ヘッダ参照', '=C$2&":"&C{r}')]):
        ws.cell(row=2, column=6 + i, value=h).font = Font(bold=True)
        for r in range(3, last + 1):
            ws.cell(row=r, column=6 + i, value=t.format(r=r))


def item_unsafe(wb, order):
    ws = wb.active
    std_rows(ws, order)
    put_formulas(ws, [
        ('外部', "='C:\\data\\[other.xlsx]Sheet1'!A1"),
        ('3D', '=SUM(評価:Sheet2!B3)'),
        ('INDIRECT', '=INDIRECT("C"&ROW())'),
        ('OFFSET', '=OFFSET(C{r},0,1)'),
        ('ADDRESS', '=ADDRESS(ROW(),COLUMN())'),
        ('通常', '=C{r}+1'),
        ('全角', '=C{r}＋1'),
    ])
    wb.create_sheet('Sheet2')['B3'] = 5


STD_ITEMS = [
    ('f01_samerow', item_samerow, '同じ行だけを読む四則・IF・AVERAGE・日付・エラー'),
    ('f02_shared', item_shared, '共有数式（Excel で FillDown）と、それを読む数式'),
    ('f03_keycol', item_keycol, '取り込まない列（学生ID・氏名）を読む数式'),
    ('f04_lookups', item_lookups, 'VLOOKUP / INDEX-MATCH / SUMIF / COUNTIF / RANK / LARGE / XLOOKUP'),
    ('f05_ranges', item_ranges, '表全体・見出し込み・D:D・D2:D1000・1行ずれの範囲'),
    ('f06_running', item_running, '累計・前行差・構成比・脇のセル・行番号'),
    ('f07_othersheet', item_othersheet, '別シート参照・自シート修飾'),
    ('f08_names', item_names, '定義名（定数・範囲・脇のセル・別シート・OFFSET・シート名スコープ）'),
    ('f09_tables', item_tables, 'テーブルと構造化参照'),
    ('f11_values', item_values, '文字・日付・時刻・真偽・エラー・空文字'),
    ('f12_footer', item_footer, '表の下の合計行・平均行を読む数式'),
    ('f13_hyperlink', item_hyperlink, 'HYPERLINK / CONCAT'),
    ('f16_unsafe', item_unsafe, '外部ブック・3D・INDIRECT・OFFSET・ADDRESS・全角演算子'),
]


def write_dynamic(path, order):
    """XlsxWriter: legacy CSE single-cell arrays, dynamic-array single cell + a spilling UNIQUE with a spill reference"""
    wb = xlsxwriter.Workbook(path, {'use_future_functions': True})
    ws = wb.add_worksheet(SHEET)
    ids = ordered(order)
    for i, h in enumerate(['学生ID', '氏名', '評価1', '評価2', '評価3', '重複排除', '件数', 'FILTER和', '先頭5件']):
        ws.write(0, i, h)
    for k, sid in enumerate(ids):
        r = 1 + k
        ws.write(r, 0, sid)
        ws.write(r, 1, person_name(sid))
        for j in range(3):
            ws.write(r, 2 + j, score(sid, j))
        ws.write_array_formula(r, 6, r, 6, '{=SUM(IF($C$2:$C$32>C%d,1,0))}' % (r + 1))
        ws.write_dynamic_array_formula(r, 7, r, 7, '=SUM(FILTER($D$2:$D$32,$C$2:$C$32>C%d))' % (r + 1))
    ws.write_dynamic_array_formula(1, 5, 1, 5, '=UNIQUE(E2:E32)')
    ws.write_dynamic_array_formula(1, 8, 1, 8, '=TAKE(SORT(C2:C32,,-1),5)')
    ws.write_dynamic_array_formula(1, 9, 1, 9, '=SUM(_xlfn.ANCHORARRAY(F2))')
    wb.close()


SAMPLES = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'samples', '01_学生名簿.xlsx'))


def samples_people():
    """(学生ID, 氏名) of the 60 students of samples/01_学生名簿.xlsx (header row 3, rows 4..63; the ID is a 6-digit text)"""
    wb = openpyxl.load_workbook(SAMPLES, read_only=False)
    ws = wb['名簿']
    return [(str(ws.cell(r, 1).value), str(ws.cell(r, 2).value)) for r in range(4, 64)]


def write_samples_item(path, order):
    """A source for the app's own sample base (merged title, header row 3, names, charts, pivot, footer): 57 of its 60 students + 3 strangers."""
    people = samples_people()
    persons = people[:57] + [('009991', '架空　一'), ('009992', '架空　二'), ('009993', '架空　三')]
    if order == 'rev':
        persons.reverse()
    elif order == 'shuf':
        random.Random(7).shuffle(persons)
    n = len(persons)
    last = n + 1
    perm = [random.Random(2000 + k).sample(range(40, 100), n) for k in range(3)]
    idx = {p[0]: i for i, p in enumerate(people[:57] + [('009991', ''), ('009992', ''), ('009993', '')])}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    for i, h in enumerate(['学生ID', '氏名', '評価1', '評価2', '評価3']):
        ws.cell(row=1, column=i + 1, value=h).font = Font(bold=True)
    wb.defined_names['税率'] = DefinedName('税率', attr_text='0.1')
    cols = [
        ('合計', '=C{r}+D{r}+E{r}'),
        ('様', '=B{r}&"様"'),
        ('偏差', '=ROUND((C{r}-AVERAGE(C$2:C$%d))/STDEV(C$2:C$%d)*10+50,1)' % (last, last)),
        ('順位', '=RANK(C{r},$C$2:$C$%d)' % last),
        ('前行差', lambda r: '=C2' if r == 2 else '=C%d-C%d' % (r, r - 1)),
        ('累計', '=SUM(C$2:C{r})'),
        ('引当', '=INDEX($D$2:$D$%d,MATCH(A{r},$A$2:$A$%d,0))' % (last, last)),
        ('構成比', '=C{r}/SUM(C:C)'),
        ('税込', '=C{r}*税率'),
        ('重複', '=IF(COUNTIF($A$2:$A$%d,A{r})>1,"重複","")' % last),
    ]
    for k, (pid, name) in enumerate(persons):
        r = 2 + k
        ws.cell(row=r, column=1, value=pid)
        ws.cell(row=r, column=2, value=name)
        for j in range(3):
            ws.cell(row=r, column=3 + j, value=perm[j][idx[pid]])
    put_formulas(ws, cols, first=2, last=last)
    wb.save(path)
    return last


def main(out):
    os.makedirs(os.path.join(out, 'nocache'), exist_ok=True)
    write_base(os.path.join(out, 'base.xlsx'))
    import shutil
    shutil.copyfile(SAMPLES, os.path.join(out, 'base_samples.xlsx'))
    manifest = []
    for order in ORDERS:
        fname = 'g01_samples__%s.xlsx' % order
        last = write_samples_item(os.path.join(out, 'nocache', fname), order)
        manifest.append({'name': 'g01_samples', 'order': order, 'file': fname, 'desc': 'サンプルの名簿（結合セル・見出し3行目・名前・グラフ・ピボット・合計行）を基準にした取り込み', 'sheet': SHEET, 'headerRow': 1, 'firstDataRow': 2, 'lastDataRow': last, 'comPost': [], 'base': 'base_samples.xlsx'})
    for name, fn, desc in STD_ITEMS + [('f14_twoheader', item_twoheader, '2行見出し'), ('f10_arrays', None, '配列・動的配列・スピル')]:
        for order in ORDERS:
            fname = '%s__%s.xlsx' % (name, order)
            path = os.path.join(out, 'nocache', fname)
            entry = {'name': name, 'order': order, 'file': fname, 'desc': desc, 'sheet': SHEET, 'headerRow': 1, 'firstDataRow': 2, 'lastDataRow': LAST, 'comPost': []}
            if name == 'f10_arrays':
                write_dynamic(path, order)
            else:
                wb = openpyxl.Workbook()
                wb.active.title = SHEET
                fn(wb, order)
                wb.save(path)
            if name == 'f14_twoheader':
                entry.update(headerRow=2, firstDataRow=3, lastDataRow=33)
            if name == 'f02_shared':
                entry['comPost'] = [
                    {'op': 'formula', 'range': 'F2:F32', 'formula': '=C2+D2'},
                    {'op': 'formula', 'range': 'G2:G32', 'formula': '=F2*2'},
                    {'op': 'formula', 'range': 'H2:H32', 'formula': '=IF(G2>200,"A","B")'},
                ]
            if name == 'f01_samerow' and order == 'same':
                entry['comSaveXlsm'] = True
            manifest.append(entry)
    with open(os.path.join(out, 'manifest.json'), 'w', encoding='utf-8') as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=1)
    print('wrote %d source files' % len(manifest))


if __name__ == '__main__':
    main(sys.argv[1])
