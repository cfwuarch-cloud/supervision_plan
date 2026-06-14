# -*- coding: utf-8 -*-
"""
材料送審管制總表轉換工具 v2.0
=============================
將 詳細價目表.xlsx 作為母本，填入 表5.1.docx 模板，
自動分頁、合併欄位，產出完整之材料送審管制總表。

修正歷程：
  v2.0  2026/06/14  重寫資料列建立方式（etree.SubElement），
                    同步表5.2 v2.0 架構：文字清洗、黑實線框、
                    exact 行高、表頭三行、通用表格元件

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
  python -X utf8 tables/table5.1/convert_5.1_v2.py --exclude-units 式 工

功能說明：
  1. 讀取詳細價目表（excel），依項次壹.三.1 之後、排除指定單位
  2. 以 Pillow 測量各材料名稱字寬，計算實際所需列高
  3. 動態填滿每頁可用空間，確保不跨頁
  4. 資料列使用 etree.SubElement 重建，避免 deepcopy 範本屬性
  5. 黑實線框、exact 行高 + 段落 exact 行高，文字不裁切
  6. ⑤→窗（修正 PDF 轉檔錯誤）
  7. 每頁之間插入分頁符號，每頁頂部含表頭三行

參數說明：
  -p, --price   詳細價目表 Excel 路徑（預設：../../data/02_成德-詳細價目表.xlsx）
  -t, --template  表5.1模板 docx 路徑（預設：./表5.1.docx）
  -o, --output   輸出 docx 路徑（預設：../../output/表5.1_完成.docx）
  --exclude-units  排除單位（預設：式 工）
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

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
from common.docx_table import add_cell


# ── 常數 ──
LINE_H_TWIP = 240
CELL_TOP_TWIP = 0
MIN_ROW_H_TWIP = 226
C1_WIDTH_TWIP = 2899
TITLE_H_TWIP = 1900
PAGE_H_TWIP = 13958
DATA_AVAIL_TWIP = 11170
HEADER_H_TWIP = 2600  # 表頭三列(1900) + 三段段落(~700)

COL_W = [288, 2899, 857, 499, 1185, 547, 438, 488, 394, 455, 404, 434, 335, 865, 438]


def clean_text(text):
    import re
    text = text.replace('\n', '').replace('\r', '')
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[\s]+([、，,。．])', r'\1', text)
    return text.strip()


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


def calc_row_height(n_lines):
    return max(MIN_ROW_H_TWIP, n_lines * LINE_H_TWIP + CELL_TOP_TWIP)


def load_price_sheet(path, exclude_units=None):
    if exclude_units is None:
        exclude_units = {'式', '工'}
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


def add_page_break(body):
    pb_p = etree.SubElement(body, qn('w:p'))
    pb_r = etree.SubElement(pb_p, qn('w:r'))
    pb_br = etree.SubElement(pb_r, qn('w:br'))
    pb_br.set(qn('w:type'), 'page')


PROJECT_NAME = '臺南市政府社會局委託辦理北區成德公設民營托嬰中心室內裝修統包工程'


def add_header(body, page_num, total_pages):
    """加入表格上方三行表頭"""
    p1 = etree.SubElement(body, qn('w:p'))
    pPr1 = etree.SubElement(p1, qn('w:pPr'))
    jc1 = etree.SubElement(pPr1, qn('w:jc'))
    jc1.set(qn('w:val'), 'center')
    sp1 = etree.SubElement(pPr1, qn('w:spacing'))
    sp1.set(qn('w:line'), '480')
    sp1.set(qn('w:lineRule'), 'exact')
    r1 = etree.SubElement(p1, qn('w:r'))
    rPr1 = etree.SubElement(r1, qn('w:rPr'))
    rFonts1 = etree.SubElement(rPr1, qn('w:rFonts'))
    rFonts1.set(qn('w:ascii'), 'Arial')
    rFonts1.set(qn('w:hAnsi'), 'Arial')
    rFonts1.set(qn('w:eastAsia'), '標楷體')
    etree.SubElement(rPr1, qn('w:b'))
    sz1 = etree.SubElement(rPr1, qn('w:sz'))
    sz1.set(qn('w:val'), '28')
    szCs1 = etree.SubElement(rPr1, qn('w:szCs'))
    szCs1.set(qn('w:val'), '28')
    spc1 = etree.SubElement(rPr1, qn('w:spacing'))
    spc1.set(qn('w:val'), '6')
    t1 = etree.SubElement(r1, qn('w:t'))
    t1.text = f'表5-1 材料設備送審管制總表-{page_num}'

    p2 = etree.SubElement(body, qn('w:p'))
    pPr2 = etree.SubElement(p2, qn('w:pPr'))
    jc2 = etree.SubElement(pPr2, qn('w:jc'))
    jc2.set(qn('w:val'), 'left')
    r2 = etree.SubElement(p2, qn('w:r'))
    rPr2 = etree.SubElement(r2, qn('w:rPr'))
    rFonts2 = etree.SubElement(rPr2, qn('w:rFonts'))
    rFonts2.set(qn('w:ascii'), 'Arial')
    rFonts2.set(qn('w:hAnsi'), 'Arial')
    rFonts2.set(qn('w:eastAsia'), '標楷體')
    sz2 = etree.SubElement(rPr2, qn('w:sz'))
    sz2.set(qn('w:val'), '24')
    szCs2 = etree.SubElement(rPr2, qn('w:szCs'))
    szCs2.set(qn('w:val'), '24')
    t2 = etree.SubElement(r2, qn('w:t'))
    t2.text = f'工程名稱：{PROJECT_NAME}'

    p3 = etree.SubElement(body, qn('w:p'))
    pPr3 = etree.SubElement(p3, qn('w:pPr'))
    jc3 = etree.SubElement(pPr3, qn('w:jc'))
    jc3.set(qn('w:val'), 'left')
    sp3 = etree.SubElement(pPr3, qn('w:spacing'))
    sp3.set(qn('w:line'), '240')
    sp3.set(qn('w:lineRule'), 'atLeast')
    r3 = etree.SubElement(p3, qn('w:r'))
    rPr3 = etree.SubElement(r3, qn('w:rPr'))
    rFonts3 = etree.SubElement(rPr3, qn('w:rFonts'))
    rFonts3.set(qn('w:ascii'), 'Times New Roman')
    rFonts3.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts3.set(qn('w:eastAsia'), '標楷體')
    sz3 = etree.SubElement(rPr3, qn('w:sz'))
    sz3.set(qn('w:val'), '24')
    szCs3 = etree.SubElement(rPr3, qn('w:szCs'))
    szCs3.set(qn('w:val'), '24')
    t3 = etree.SubElement(r3, qn('w:t'))
    t3.text = f'(監造單位使用)                              ' \
              f'第{page_num}頁共{total_pages}頁   表單編號：E51-{page_num}'


def convert(price_path, template_path, output_path,
            exclude_units=None, max_pairs=20, max_pages=0):
    items = load_price_sheet(price_path, exclude_units)
    print(f'價目表載入：{len(items)} 項')

    pair_heights = []
    odd_heights = []
    even_heights = []
    for item, name, qty_unit in items:
        odd_lines = calc_name_lines(item)
        even_lines = calc_name_lines(name)
        h_odd = calc_row_height(odd_lines)
        h_even = calc_row_height(even_lines)
        pair_heights.append(h_odd + h_even)
        odd_heights.append(h_odd)
        even_heights.append(h_even)

    doc = Document(template_path)
    body = doc.element.body
    t0 = doc.tables[0]

    title_trs = t0._tbl.findall(qn('w:tr'))[:2]
    title_xml = [deepcopy(tr) for tr in title_trs]
    sect = body.find(qn('w:sectPr'))
    sect_xml = deepcopy(sect) if sect is not None else None

    for child in list(body):
        body.remove(child)

    DAT_MERGE = {0, 2, 3, 6, 7, 8, 9, 10, 11, 12, 14}
    seq = 0
    idx = 0

    page_plan = []
    tmp_idx = 0
    while tmp_idx < len(items):
        if 0 < max_pages <= len(page_plan):
            break
        acc_h = 0
        pairs = 0
        for pi in range(len(items) - tmp_idx):
            h = pair_heights[tmp_idx + pi]
            if acc_h + h > DATA_AVAIL_TWIP - HEADER_H_TWIP and pairs > 0:
                break
            acc_h += h
            pairs += 1
        if pairs == 0:
            break
        page_plan.append(pairs)
        tmp_idx += pairs

    total_pages = len(page_plan)

    for page_num, pairs_this_page in enumerate(page_plan):
        if page_num > 0:
            add_page_break(body)

        add_header(body, page_num + 1, total_pages)

        tbl = etree.SubElement(body, qn('w:tbl'))
        tblPr = etree.SubElement(tbl, qn('w:tblPr'))
        tblW = etree.SubElement(tblPr, qn('w:tblW'))
        tblW.set(qn('w:w'), '5289')
        tblW.set(qn('w:type'), 'pct')
        tblGrid = etree.SubElement(tbl, qn('w:tblGrid'))
        for w in COL_W:
            gc = etree.SubElement(tblGrid, qn('w:gridCol'))
            gc.set(qn('w:w'), str(w))

        for tr in title_xml:
            tbl.append(deepcopy(tr))

        for pi in range(pairs_this_page):
            if idx >= len(items):
                break
            seq += 1
            item, name, qty_unit = items[idx]

            odd_text = [str(seq), item, qty_unit, '', '', '', '', '', '', '', '', '', '', '', '']
            even_text = ['', name, '', '', '', '', '', '', '', '', '', '', '', '', '']

            tr1 = etree.SubElement(tbl, qn('w:tr'))
            trPr1 = etree.SubElement(tr1, qn('w:trPr'))
            th1 = etree.SubElement(trPr1, qn('w:trHeight'))
            th1.set(qn('w:val'), str(odd_heights[idx]))
            th1.set(qn('w:hRule'), 'atLeast')
            for ci, txt in enumerate(odd_text):
                add_cell(tr1, txt, COL_W[ci],
                         center=ci != 1, bold=False,
                         merge_restart=True if ci in DAT_MERGE else None,
                         line_twip=LINE_H_TWIP)

            tr2 = etree.SubElement(tbl, qn('w:tr'))
            trPr2 = etree.SubElement(tr2, qn('w:trPr'))
            th2 = etree.SubElement(trPr2, qn('w:trHeight'))
            th2.set(qn('w:val'), str(even_heights[idx]))
            th2.set(qn('w:hRule'), 'atLeast')
            for ci, txt in enumerate(even_text):
                add_cell(tr2, txt, COL_W[ci],
                         center=ci != 1, bold=False,
                         merge_restart=False if ci in DAT_MERGE else None,
                         line_twip=LINE_H_TWIP)

            idx += 1

    page_num = total_pages
    body.append(etree.SubElement(body, qn('w:p')))
    if sect_xml is not None:
        body.append(sect_xml)

    doc.save(output_path)
    print(f'已完成：{seq} 項，共 {page_num} 頁')
    return page_num, seq


def main():
    parser = argparse.ArgumentParser(
        description='材料送審管制總表轉換工具 v2.0')
    parser.add_argument('-p', '--price', default='../../data/02_成德-詳細價目表.xlsx',
                        help='詳細價目表 Excel 路徑')
    parser.add_argument('-t', '--template', default='./表5.1.docx',
                        help='表5.1 模板 docx 路徑')
    parser.add_argument('-o', '--output', default='../../output/表5.1_完成.docx',
                        help='輸出 docx 路徑')
    parser.add_argument('--test-num', type=int,
                        help='測試流水號，指定後輸出檔名自動插入 _test_N')
    parser.add_argument('--exclude-units', nargs='*', default=['式', '工'],
                        help='排除單位（預設：式 工）')
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
