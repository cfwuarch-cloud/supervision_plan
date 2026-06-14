# -*- coding: utf-8 -*-
"""
材料設備檢(試)驗管制總表轉換工具
=============================
將 詳細價目表.xlsx 作為母本，填入 表5.2.docx 模板，
自動分頁、合併欄位，產出完整之材料設備檢(試)驗管制總表。

修正歷程：
  v1.0  2026/06/13  初始版本（比照表5.1架構，10欄位）

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/13

相依套件：
  - openpyxl>=3.0.0,<4.0.0
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0
  - pandas>=1.5.0,<3.0.0
  - Pillow>=9.0.0,<12.0.0

使用方法：
  python -X utf8 supervision_plan/table5.2_converter/convert_5.2.py -p 價目表.xlsx -t 表5.2.docx -o 輸出.docx

功能說明：
  1. 讀取詳細價目表（excel），依項次壹.三.1 之後、排除單位「工」
  2. 以 Pillow 測量各材料名稱字寬，計算實際所需列高
  3. 動態填滿每頁可用空間（總高 11500 twip = 標題 971 + 資料 10529），確保不跨頁
  4. 標題列合併規則：C0/C3/C5/C7/C8（R0+R1 垂直合併）
  5. 資料列合併規則：C0/C3/C5/C7/C8（奇偶列垂直合併）
  6. ⑤→窗（修正 PDF 轉檔錯誤）
  7. 每頁之間插入分頁符號

參數說明：
  -p, --price   詳細價目表 Excel 路徑（預設：../02_成德-詳細價目表.xlsx）
  -t, --template  表5.2模板 docx 路徑（預設：../表5.2.docx）
  -o, --output   輸出 docx 路徑（預設：../表5.2_完成.docx）
  --exclude-units  排除單位（預設：工）
  --test-num  測試流水號，輸出檔名自動插入 test_N
  --max-pairs  每頁最多資料組數（預設：20）
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


# ── 行高校準常數（比照表5.1） ──
LINE_H_TWIP = 240        # 每行文字高（twip）
CELL_TOP_TWIP = 15       # 儲存格上邊距（tcMar top）
MIN_ROW_H_TWIP = 226     # 資料列最小高度
C1_WIDTH_TWIP = 2338     # 名稱欄寬（gridCol 第 1 欄，4.12cm）
TITLE_H_TWIP = 971       # 標題列總高（模板實測 R0=463 + R1=508）
PAGE_H_TWIP = 13958      # 頁面可用高度（16838 − 上 1440 − 下 1440）
DATA_AVAIL_TWIP = 10529  # 資料列可用高度（11500 − 標題 971）

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
        exclude_units = {'工'}
    df = pd.read_excel(path, sheet_name='Table 1')
    c0 = df.columns[0]
    c1 = df.columns[1]
    cu = df.columns[6]
    cq = df.columns[7]

    items = []
    started = False
    for _, row in df.iterrows():
        item = str(row[c0]).strip()
        name = str(row[c1]).strip()
        unit = str(row[cu]).strip() if pd.notna(row[cu]) else ''
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


def set_tr_height(tr, h_twip):
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = etree.SubElement(tr, qn('w:trPr'))
    old = trPr.find(qn('w:trHeight'))
    if old is not None:
        trPr.remove(old)
    he = etree.SubElement(trPr, qn('w:trHeight'))
    he.set(qn('w:val'), str(h_twip))
    he.set(qn('w:hRule'), 'exact')


def set_tc_text(tc, text):
    p = tc.find(qn('w:p'))
    if p is None:
        p = etree.SubElement(tc, qn('w:p'))
    r = p.find(qn('w:r'))
    if r is None:
        r = etree.SubElement(p, qn('w:r'))
    t = r.find(qn('w:t'))
    if t is None:
        t = etree.SubElement(r, qn('w:t'))
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(r, qn('w:rPr'))
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), '標楷體')
    sz = rPr.find(qn('w:sz'))
    if sz is None:
        sz = etree.SubElement(rPr, qn('w:sz'))
    sz.set(qn('w:val'), '20')
    szCs = rPr.find(qn('w:szCs'))
    if szCs is None:
        szCs = etree.SubElement(rPr, qn('w:szCs'))
    szCs.set(qn('w:val'), '20')


def set_vmerge(tc, restart=True):
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn('w:tcPr'))
        tc.remove(tcPr)
        tc.insert(0, tcPr)
    vm = tcPr.find(qn('w:vMerge'))
    if vm is None:
        vm = etree.SubElement(tcPr, qn('w:vMerge'))
    if restart:
        vm.set(qn('w:val'), 'restart')
    else:
        if qn('w:val') in vm.attrib:
            del vm.attrib[qn('w:val')]


def remove_all_vmerge(table):
    for row in table.rows:
        tr = row._tr
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                vm = tcPr.find(qn('w:vMerge'))
                if vm is not None:
                    tcPr.remove(vm)


def add_empty_rows(table, n_pairs, template_row_index=2):
    if n_pairs <= 0:
        return
    template_tr = table.rows[template_row_index]._tr
    for _ in range(n_pairs * 2):
        table._tbl.append(deepcopy(template_tr))


def apply_header_merges(t0, merge_cols):
    """
    對標題列 R0/R1 套用垂直合併。
    表5.2 無 gridSpan 跨欄，所有 tc 一一對應。
    """
    r0_tcs = list(t0.rows[0]._tr.findall(qn('w:tc')))
    r1_tcs = list(t0.rows[1]._tr.findall(qn('w:tc')))
    for ci in merge_cols:
        set_vmerge(r0_tcs[ci], restart=True)
        set_vmerge(r1_tcs[ci], restart=False)


def fill_data_pair(trs, ri, seq, item, name, qty_unit, merge_cols,
                   odd_h=None, even_h=None):
    """
    填入一組資料（奇偶列），對 merge_cols 套用垂直合併。
    表5.2 資料結構：
      C0 = 項次（奇偶合併）
      C1 = 奇：項次編號、偶：材料名稱
      C2-C9 依應用需求填入
    """
    r1_tr = trs[ri]
    r2_tr = trs[ri + 1]
    r1_tcs = r1_tr.findall(qn('w:tc'))
    r2_tcs = r2_tr.findall(qn('w:tc'))

    # C0：項次（流水號）—奇偶合併
    set_tc_text(r1_tcs[0], str(seq))
    set_vmerge(r1_tcs[0], restart=True)
    set_vmerge(r2_tcs[0], restart=False)

    # C1：奇=項次編碼、偶=材料名稱（不併）
    set_tc_text(r1_tcs[1], item)
    set_tc_text(r2_tcs[1], name)

    # C3、C5、C7、C8：合併
    for ci in merge_cols:
        if ci == 0:
            continue
        set_vmerge(r1_tcs[ci], restart=True)
        set_vmerge(r2_tcs[ci], restart=False)
        set_tc_text(r1_tcs[ci], '')
        set_tc_text(r2_tcs[ci], '')

    # C2、C4、C6、C9：不併，填入空字串
    for ci in [2, 4, 6, 9]:
        set_tc_text(r1_tcs[ci], '')
        set_tc_text(r2_tcs[ci], '')

    # 數量+單位（奇列 C2）
    set_tc_text(r1_tcs[2], qty_unit)

    if odd_h is not None:
        set_tr_height(r1_tr, odd_h)
    if even_h is not None:
        set_tr_height(r2_tr, even_h)


def add_page_break(body):
    pb_p = etree.SubElement(body, qn('w:p'))
    pb_r = etree.SubElement(pb_p, qn('w:r'))
    pb_br = etree.SubElement(pb_r, qn('w:br'))
    pb_br.set(qn('w:type'), 'page')


def convert(price_path, template_path, output_path,
            exclude_units=None, max_pairs=20):
    """
    主轉換程式（動態分頁）。
    """
    # 讀取資料
    items = load_price_sheet(price_path, exclude_units)
    print(f'價目表載入：{len(items)} 項')

    # 預先計算每組高度
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

    # 載入模板
    doc = Document(template_path)
    t0 = doc.tables[0]

    # 清除既有 vMerge
    remove_all_vmerge(t0)

    # 標題合併：C0, C3, C5, C7, C8
    HDR_MERGE = {0, 3, 5, 7, 8}
    apply_header_merges(t0, HDR_MERGE)

    # 資料合併欄位
    DAT_MERGE = {0, 3, 5, 7, 8}

    # 擴增空白資料列
    cur_pairs = (len(t0.rows) - 2) // 2
    need = 50 - cur_pairs
    if need > 0:
        add_empty_rows(t0, need)

    # 儲存主模板 XML
    master_xml = deepcopy(t0._tbl)

    # 清空文件，保留 sectPr
    body = doc.element.body
    old_sectPr = body.find(qn('w:sectPr'))
    for child in list(body):
        body.remove(child)

    # 多頁動態分頁
    seq = 0
    idx = 0
    page_num = 0

    while idx < len(items):
        if page_num > 0:
            add_page_break(body)

        cur_tbl = deepcopy(master_xml)
        body.append(cur_tbl)
        cur_trs = cur_tbl.findall(qn('w:tr'))
        total_pairs = (len(cur_trs) - 2) // 2

        # 計算此頁可放多少對
        acc_h = 0
        pairs_this_page = 0
        for pi in range(total_pairs):
            if idx + pi >= len(items):
                break
            h = pair_heights[idx + pi]
            if acc_h + h > DATA_AVAIL_TWIP and pairs_this_page > 0:
                break
            acc_h += h
            pairs_this_page += 1

        for pi in range(pairs_this_page):
            if idx >= len(items):
                break
            seq += 1
            item, name, qty_unit = items[idx]
            ri = 2 + pi * 2
            fill_data_pair(
                cur_trs, ri, seq, item, name, qty_unit, DAT_MERGE,
                odd_h=odd_heights[idx], even_h=even_heights[idx],
            )
            idx += 1

        # 刪除未填入的多餘資料列
        used_rows = 2 + pairs_this_page * 2
        all_rows = cur_tbl.findall(qn('w:tr'))
        for extra_tr in all_rows[used_rows:]:
            cur_tbl.remove(extra_tr)

        page_num += 1

    # 補回頁面設定
    if old_sectPr is not None:
        body.append(deepcopy(old_sectPr))

    doc.save(output_path)
    print(f'已完成：{seq} 項，共 {page_num} 頁')
    return page_num, seq


def main():
    parser = argparse.ArgumentParser(
        description='材料設備檢(試)驗管制總表轉換工具 v1.0')
    parser.add_argument('-p', '--price', default='./02_成德-詳細價目表.xlsx',
                        help='詳細價目表 Excel 路徑')
    parser.add_argument('-t', '--template', default='./表5.2.docx',
                        help='表5.2 模板 docx 路徑')
    parser.add_argument('-o', '--output', default='../Vbaoffice/表5.2_完成.docx',
                        help='輸出 docx 路徑')
    parser.add_argument('--test-num', type=int,
                        help='測試流水號，指定後輸出檔名自動插入 _test_N')
    parser.add_argument('--exclude-units', nargs='*', default=['工'],
                        help='排除單位（預設：工）')
    parser.add_argument('--max-pairs', type=int, default=20,
                        help='每頁最多資料組數（預設：20）')

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
            max_pairs=args.max_pairs)


if __name__ == '__main__':
    main()
