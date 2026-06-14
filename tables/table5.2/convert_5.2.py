# -*- coding: utf-8 -*-
"""
材料設備檢(試)驗管制總表轉換工具
=============================
將 詳細價目表.xlsx 作為母本，填入 表5.2.docx 模板，
自動分頁、合併欄位，產出完整之材料設備檢(試)驗管制總表。

修正歷程：
  v1.0  2026/06/13  初始版本（比照表5.1架構，10欄位）
  v2.0  2026/06/14  重寫資料列建立方式，改用 etree.SubElement 避免範本屬性導致 Word 頁數異常

作者：OpenCode Assistant / cfwuarch
版本：v2.0
最後更新：2026/06/14

相依套件：
  - openpyxl>=3.0.0,<4.0.0
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0
  - pandas>=1.5.0,<3.0.0
  - Pillow>=9.0.0,<12.0.0

使用方法：
  python -X utf8 tables/table5.2/convert_5.2.py -p data/價目表.xlsx -t tables/table5.2/表5.2.docx -o output/輸出.docx

功能說明：
  1. 讀取詳細價目表（excel），依項次壹.三.1 之後、排除指定單位
  2. 以 Pillow 測量各材料名稱字寬及規定抽樣頻率文字寬，計算實際所需列高
  3. 動態填滿每頁可用空間，確保不跨頁
  4. 標題列合併規則：C0/C3/C5/C7/C8（R0+R1 垂直合併）
  5. 資料列合併規則：C0/C3/C5/C7/C8（奇偶列垂直合併）
  6. ⑤→窗（修正 PDF 轉檔錯誤）
  7. C5 規定抽樣頻率：每頁第一組為「進場時至少一次」，其餘為「至少一次」
  8. 每頁之間插入分頁符號

參數說明：
  -p, --price   詳細價目表 Excel 路徑（預設：../../data/02_成德-詳細價目表.xlsx）
  -t, --template  表5.2模板 docx 路徑（預設：./表5.2.docx）
  -o, --output   輸出 docx 路徑（預設：../../output/表5.2_完成.docx）
  --exclude-units  排除單位（預設：工）
  --test-num  測試流水號，輸出檔名自動插入 test_N
  --max-pairs  每頁最多資料組數（預設：20）
  --max-pages  最多頁數（預設：0=不限）
"""
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from copy import deepcopy
from PIL import ImageFont, ImageDraw, Image
import argparse
import os
import sys
import math


# ── 常數 ──
LINE_H_TWIP = 240
CELL_TOP_TWIP = 15
MIN_ROW_H_TWIP = 226
C1_WIDTH_TWIP = 2338
C5_WIDTH_TWIP = 506
TITLE_H_TWIP = 971
PAGE_H_TWIP = 13958
DATA_AVAIL_TWIP = 10600


def clean_text(text):
    import re
    text = text.replace('\n', '').replace('\r', '')
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[\s]+([、，,。．])', r'\1', text)
    return text.strip()
COL_W = [380, 2338, 1214, 671, 804, 506, 877, 623, 621, 838]

_FONT_CACHE = None


def _get_font():
    global _FONT_CACHE
    if _FONT_CACHE is None:
        try:
            _FONT_CACHE = ImageFont.truetype('kaiu', 10)
        except Exception:
            _FONT_CACHE = ImageFont.load_default()
    return _FONT_CACHE


def calc_name_lines(name_text):
    font = _get_font()
    img = Image.new('RGB', (1, 1))
    draw = ImageDraw.Draw(img)
    segments = name_text.split(chr(10))
    total_lines = 0
    for seg in segments:
        if not seg.strip():
            total_lines += 1
            continue
        bbox = draw.textbbox((0, 0), seg, font=font)
        w_px = bbox[2] - bbox[0]
        w_twip = w_px * 20
        lines = max(1, math.ceil(w_twip / C1_WIDTH_TWIP))
        total_lines += lines
    return total_lines


def calc_text_lines(text, col_width):
    font = _get_font()
    img = Image.new('RGB', (1, 1))
    draw = ImageDraw.Draw(img)
    segments = text.split(chr(10))
    total_lines = 0
    for seg in segments:
        if not seg.strip():
            total_lines += 1
            continue
        bbox = draw.textbbox((0, 0), seg, font=font)
        w_px = bbox[2] - bbox[0]
        w_twip = w_px * 20
        lines = max(1, math.ceil(w_twip / col_width))
        total_lines += lines
    return total_lines


def calc_row_height(n_lines):
    return max(MIN_ROW_H_TWIP, n_lines * LINE_H_TWIP + CELL_TOP_TWIP)


def load_price_sheet(path, exclude_units=None):
    if exclude_units is None:
        exclude_units = {'工'}
    df = pd.read_excel(path, sheet_name='Table 1')
    c0 = df.columns[0]
    c1 = df.columns[1]
    cu = df.columns[6]
    cq = df.columns[7]

    items = []
    started = False
    for _, row in df.iterrows():
        item = clean_text(str(row[c0]))
        name = clean_text(str(row[c1]))
        unit = clean_text(str(row[cu])) if pd.notna(row[cu]) else ''
        qty = row[cq]

        name = name.replace('⑤', '窗')
        unit = unit.replace('⑤', '窗')

        if '壹.三.1' in item:
            started = True
        if not started:
            continue
        if not item.startswith('壹.'):
            continue
        if not unit or unit == 'nan':
            continue
        if unit in exclude_units:
            continue
        if any(k in name for k in ['小計', '合計', '總價']):
            continue
        if pd.isna(qty):
            continue

        if isinstance(qty, float) and qty == int(qty):
            qty_str = str(int(qty))
        else:
            qty_str = str(qty)
        items.append((item, name, qty_str + unit))

    return items


def make_cell(parent, text, ci, merge_restart=None):
    text = clean_text(text or '')
    tc = etree.SubElement(parent, qn('w:tc'))
    tcPr = etree.SubElement(tc, qn('w:tcPr'))
    tcW = etree.SubElement(tcPr, qn('w:tcW'))
    tcW.set(qn('w:w'), str(COL_W[ci]))
    tcW.set(qn('w:type'), 'dxa')
    if merge_restart is not None:
        vm = etree.SubElement(tcPr, qn('w:vMerge'))
        if merge_restart:
            vm.set(qn('w:val'), 'restart')
    vAlign = etree.SubElement(tcPr, qn('w:vAlign'))
    vAlign.set(qn('w:val'), 'center')
    p = etree.SubElement(tc, qn('w:p'))
    pPr = etree.SubElement(p, qn('w:pPr'))
    spacing = etree.SubElement(pPr, qn('w:spacing'))
    spacing.set(qn('w:line'), str(LINE_H_TWIP))
    spacing.set(qn('w:lineRule'), 'exact')
    if text:
        r = etree.SubElement(p, qn('w:r'))
        rPr = etree.SubElement(r, qn('w:rPr'))
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:eastAsia'), '標楷體')
        sz = etree.SubElement(rPr, qn('w:sz'))
        sz.set(qn('w:val'), '20')
        szCs = etree.SubElement(rPr, qn('w:szCs'))
        szCs.set(qn('w:val'), '20')
        t = etree.SubElement(r, qn('w:t'))
        t.text = text
        t.set(qn('xml:space'), 'preserve')
    return tc


def add_page_break(body):
    pb_p = etree.SubElement(body, qn('w:p'))
    pb_r = etree.SubElement(pb_p, qn('w:r'))
    pb_br = etree.SubElement(pb_r, qn('w:br'))
    pb_br.set(qn('w:type'), 'page')


def convert(price_path, template_path, output_path,
            exclude_units=None, max_pairs=20, max_pages=0):
    items = load_price_sheet(price_path, exclude_units)
    print(f'價目表載入：{len(items)} 項')

    pair_heights = []
    odd_heights = []
    even_heights = []
    c5_texts = []
    first_item = True
    for item, name, qty_unit in items:
        odd_lines = calc_name_lines(item)
        even_lines = calc_name_lines(name)
        h_odd_c1 = calc_row_height(odd_lines)
        h_even = calc_row_height(even_lines)

        c5_txt = '進場時至少一次' if first_item else '至少一次'
        c5_lines = calc_text_lines(c5_txt, C5_WIDTH_TWIP)
        h_odd = max(h_odd_c1, calc_row_height(c5_lines))

        pair_heights.append(h_odd + h_even)
        odd_heights.append(h_odd)
        even_heights.append(h_even)
        c5_texts.append(c5_txt)
        first_item = False

    doc = Document(template_path)
    body = doc.element.body
    t0 = doc.tables[0]

    # 備份標題列與 sectPr
    title_trs = t0._tbl.findall(qn('w:tr'))[:2]
    title_xml = [deepcopy(tr) for tr in title_trs]
    sect = body.find(qn('w:sectPr'))
    sect_xml = deepcopy(sect) if sect is not None else None

    # 清除 body
    for child in list(body):
        body.remove(child)

    DAT_MERGE = {0, 3, 5, 7, 8}
    seq = 0
    idx = 0
    page_num = 0

    while idx < len(items):
        if 0 < max_pages <= page_num:
            break
        if page_num > 0:
            add_page_break(body)

        # 計算此頁可放多少組
        acc_h = 0
        pairs_this_page = 0
        for pi in range(len(items) - idx):
            h = pair_heights[idx + pi]
            if acc_h + h > DATA_AVAIL_TWIP and pairs_this_page > 0:
                break
            acc_h += h
            pairs_this_page += 1

        if pairs_this_page == 0:
            break

        # 建立表格
        tbl = etree.SubElement(body, qn('w:tbl'))

        # tblPr
        tblPr = etree.SubElement(tbl, qn('w:tblPr'))
        tblW = etree.SubElement(tblPr, qn('w:tblW'))
        tblW.set(qn('w:w'), '4771')
        tblW.set(qn('w:type'), 'pct')

        # tblGrid
        tblGrid = etree.SubElement(tbl, qn('w:tblGrid'))
        for w in COL_W:
            gc = etree.SubElement(tblGrid, qn('w:gridCol'))
            gc.set(qn('w:w'), str(w))

        # 標題列
        for tr in title_xml:
            tbl.append(deepcopy(tr))

        # 資料列
        first_pair = True
        for pi in range(pairs_this_page):
            if idx >= len(items):
                break
            seq += 1
            item, name, qty_unit = items[idx]

            c5_text = c5_texts[idx]
            odd_text = [str(seq), item, qty_unit, '', '', c5_text, '', '', '', '']
            even_text = ['', name, '', '', '', '', '', '', '', '']

            # 奇列
            tr1 = etree.SubElement(tbl, qn('w:tr'))
            trPr1 = etree.SubElement(tr1, qn('w:trPr'))
            th1 = etree.SubElement(trPr1, qn('w:trHeight'))
            th1.set(qn('w:val'), str(odd_heights[idx]))
            th1.set(qn('w:hRule'), 'exact')
            for ci, txt in enumerate(odd_text):
                make_cell(tr1, txt, ci,
                          True if ci in DAT_MERGE else None)

            # 偶列
            tr2 = etree.SubElement(tbl, qn('w:tr'))
            trPr2 = etree.SubElement(tr2, qn('w:trPr'))
            th2 = etree.SubElement(trPr2, qn('w:trHeight'))
            th2.set(qn('w:val'), str(even_heights[idx]))
            th2.set(qn('w:hRule'), 'exact')
            for ci, txt in enumerate(even_text):
                make_cell(tr2, txt, ci,
                          False if ci in DAT_MERGE else None)

            first_pair = False
            idx += 1

        page_num += 1

    # 補回頁面設定
    body.append(etree.SubElement(body, qn('w:p')))
    if sect_xml is not None:
        body.append(sect_xml)

    doc.save(output_path)
    print(f'已完成：{seq} 項，共 {page_num} 頁')
    return page_num, seq


def main():
    parser = argparse.ArgumentParser(
        description='材料設備檢(試)驗管制總表轉換工具 v2.0')
    parser.add_argument('-p', '--price', default='../../data/02_成德-詳細價目表.xlsx',
                        help='詳細價目表 Excel 路徑')
    parser.add_argument('-t', '--template', default='./表5.2.docx',
                        help='表5.2 模板 docx 路徑')
    parser.add_argument('-o', '--output', default='../../output/表5.2_完成.docx',
                        help='輸出 docx 路徑')
    parser.add_argument('--test-num', type=int,
                        help='測試流水號，指定後輸出檔名自動插入 _test_N')
    parser.add_argument('--exclude-units', nargs='*', default=['工'],
                        help='排除單位（預設：工）')
    parser.add_argument('--max-pairs', type=int, default=20,
                        help='每頁最多資料組數（預設：20）')
    parser.add_argument('--max-pages', type=int, default=0,
                        help='最多頁數（預設：0=不限）')

    args = parser.parse_args()

    base = os.path.dirname(os.path.abspath(__file__))
    price_path = os.path.normpath(os.path.join(base, args.price))
    template_path = os.path.normpath(os.path.join(base, args.template))
    output_path = os.path.normpath(os.path.join(base, args.output))

    if args.test_num is not None:
        root, ext = os.path.splitext(output_path)
        output_path = f'{root}_test_{args.test_num}{ext}'

    for f, label in [(price_path, '價目表'), (template_path, '模板')]:
        if not os.path.isfile(f):
            print(f'錯誤：{label} {f} 不存在')
            sys.exit(1)

    convert(price_path, template_path, output_path,
            exclude_units=set(args.exclude_units),
            max_pairs=args.max_pairs,
            max_pages=args.max_pages)


if __name__ == '__main__':
    main()
