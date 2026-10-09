#!/usr/bin/env python
"""
The 50,000-row x 30-column file of the round-5 memory / time check (module H1), written by XlsxWriter (shared strings) and by the
constant_memory mode (inline strings), plus a 50,000-row partner with the same IDs:

    python make-big.py <outdir>

Columns: 学籍番号, 氏名, 6 text columns, 8 numbers, 4 dates, 4 percent / currency columns, 3 per-row formulas, 合計 (SUM over the row) and
a last flag column = 30. Rows with the same person repeat with different IDs; nothing is random.
"""
import datetime as dt
import os
import sys

import xlsxwriter

OUT = sys.argv[1] if len(sys.argv) > 1 else '.'
os.makedirs(OUT, exist_ok=True)
N = int(os.environ.get('BIG_ROWS', '50000'))
SURN = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤', '吉田', '山田', '佐々木', '山口', '松本', '井上', '木村', '林', '斎藤', '清水']
GIVEN = ['太郎', '花子', '一郎', '美咲', '健太', '結衣', '大輔', '陽子', '翔', '愛', '拓也', '直美', '誠', '由美', '亮', '恵']
UNIV = ['慶應義塾大学', '早稲田大学', '東京大学', '一橋大学', '上智大学', '明治大学', '青山学院大学', '立教大学']


def person(i):
    return SURN[i % 20] + GIVEN[(i * 7 + 3) % 16] + ('' if i < 320 else str(i // 320))


def sid(i):
    return 'S%07d' % (2024000 + i)


def write(path, constant_memory):
    wb = xlsxwriter.Workbook(path, {'constant_memory': constant_memory})
    ws = wb.add_worksheet('名簿')
    head = wb.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1})
    date = wb.add_format({'num_format': 'yyyy/m/d'})
    money = wb.add_format({'num_format': '"¥"#,##0'})
    pct = wb.add_format({'num_format': '0.0%'})
    heads = ['学籍番号', '氏名'] + [f'区分{i}' for i in range(1, 7)] + [f'点数{i}' for i in range(1, 9)] + [f'日付{i}' for i in range(1, 5)] + ['給与', '手当', '達成率', '係数'] + ['加点1', '加点2', '判定', '合計', '有効', '備考']
    assert len(heads) == 30
    for j, h in enumerate(heads):
        ws.write(0, j, h, head)
    for i in range(1, N + 1):
        r = i  # 0-based row index = i (row 1 is the header)
        ws.write(r, 0, sid(i))
        ws.write(r, 1, person(i))
        for k in range(6):
            ws.write(r, 2 + k, UNIV[(i + k) % 8] if k < 2 else f'区分{k}-{(i * (k + 3)) % 17}')
        for k in range(8):
            ws.write_number(r, 8 + k, (i * (k + 5)) % 101)
        for k in range(4):
            ws.write_datetime(r, 16 + k, dt.datetime(2020 + k, 1, 1) + dt.timedelta(days=(i * 13) % 365), date)
        ws.write_number(r, 20, 250000 + (i % 50) * 1000, money)
        ws.write_number(r, 21, 12000 + (i % 7) * 500, money)
        ws.write_number(r, 22, ((i * 3) % 100) / 100, pct)
        ws.write_number(r, 23, 1 + (i % 9) / 10)
        row = r + 1
        ws.write_formula(r, 24, f'=I{row}+5', None, (i * 5) % 101 + 5)
        ws.write_formula(r, 25, f'=J{row}*X{row}', None, 0)
        ws.write_formula(r, 26, f'=IF(I{row}>=60,"合","否")', None, '合')
        ws.write_formula(r, 27, f'=SUM(I{row}:P{row})', None, 0)
        ws.write_boolean(r, 28, i % 2 == 0)
        ws.write(r, 29, '' if i % 5 else '要確認')
    ws.set_column(0, 1, 12)
    ws.freeze_panes(1, 2)
    wb.close()


def write_partner(path):
    wb = xlsxwriter.Workbook(path, {'constant_memory': True})
    ws = wb.add_worksheet('面接評価')
    for j, h in enumerate(['学籍番号', '氏名', '面接評価', '評価日', 'メモ']):
        ws.write(0, j, h)
    date = wb.add_format({'num_format': 'yyyy/m/d'})
    # every 3rd roster row is missing, 100 extra people at the end
    r = 1
    for i in range(1, N + 100 + 1):
        if i <= N and i % 3 == 0:
            continue
        ws.write(r, 0, sid(i))
        ws.write(r, 1, person(i))
        ws.write(r, 2, 'ABC'[i % 3])
        ws.write_datetime(r, 3, dt.datetime(2024, 9, 1) + dt.timedelta(days=i % 60), date)
        ws.write(r, 4, '良好' if i % 2 else '')
        r += 1
    wb.close()


if __name__ == '__main__':
    write(os.path.join(OUT, f'big-{N}x30-sst.xlsx'), False)
    write(os.path.join(OUT, f'big-{N}x30-inline.xlsx'), True)
    write_partner(os.path.join(OUT, f'big-partner-{N}.xlsx'))
    print('written to', OUT)
