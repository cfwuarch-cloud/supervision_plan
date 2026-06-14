# -*- coding: utf-8 -*-
"""
材料設備送審管制總表轉換工具
=============================
將 詳細價目表.xlsx 作為母本，填入 表5.1.docx 模板，
自動分頁、合併欄位，產出完整之材料設備送審管制總表。

修正歷程：
  v1.0  2026/06/12  初始版本（固定每頁 10 對）
  v2.0  2026/06/13  動態分頁：Pillow 字寬測量 + Word 實測行高校準（174 twip/行）

作者：OpenCode Assistant / cfwuarch
版本：v2.0
最後更新：2026/06/13

相依套件：
  - openpyxl>=3.0.0,<4.0.0
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0
  - pandas>=1.5.0,<3.0.0
  - Pillow>=9.0.0,<12.0.0

使用方法：
  python -X utf8 supervision_plan/table5.1_converter/convert_5.1.py -p 價目表.xlsx -t 表5.1.docx -o 輸出.docx

功能說明：
  1. 讀取詳細價目表（excel），依項次壹.三.1 之後、排除單位「式」「工」
  2. 以 Pillow 測量各材料名稱字寬，依校準行高（174 twip）計算實際所需列高
  3. 動態填滿每頁可用空間（12058 twip），確保不跨頁（每頁最多 20 對）
  4. 標題列合併規則：C0/C2/C3/C6/C14（R1+R2 垂直合併）
  5. 資料列合併規則：C0/C2/C3/C6/C7~C12/C14（奇偶列垂直合併）
  6. ⑤→窗（修正 PDF 轉檔錯誤）
  7. 每頁之間插入分頁符號

參數說明：
  -p, --price   詳細價目表 Excel 路徑（預設：../02_成德-詳細價目表.xlsx）
  -t, --template  表5模板 docx 路徑（預設：../表5.1.docx）
  -o, --output   輸出 docx 路徑（預設：../表5_完成10.docx）
  --exclude-units  排除單位（預設：式 工）
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


# ── 行高校準常數（自 Word 實測 表5_行高校準測試.docx 回推） ──
# A (20字/2行, 0.62cm=352twip) → 176 twip/行
# B (40字/3行, 0.92cm=522twip) → 174 twip/行
# C (含\n/7行, 2.13cm=1208twip) → 173 twip/行
# 取整：LINE_H_TWIP = 174（另加 4 twip 安全餘裕）
LINE_H_TWIP = 240        # 每行文字高（twip），保守取 10pt 單行行距
CELL_TOP_TWIP = 15       # 儲存格上邊距（tcMar top）
MIN_ROW_H_TWIP = 226     # 資料列最小高度（模板 trHeight atLeast）
C1_WIDTH_TWIP = 2899     # 名稱欄寬（gridCol 第 1 欄）
TITLE_H_TWIP = 1900      # 標題列總高（819 + 1081）
PAGE_H_TWIP = 13958      # 頁面可用高度（16838 − 上 1440 − 下 1440）
DATA_AVAIL_TWIP = 9600   # 資料列可用高度（使用者實測 11500 − 標題 1900）

# 快取 Pillow 字型（標楷體 10pt）
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
    """
    以 Pillow 測量材料名稱字寬，計算在 C1 欄寬內需折多少行。

    支援多行文字（\\n, chr(10)），各段獨立計算後加總。
    """
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
        w_twip = w_px * 20  # Pillow 72 DPI: 1px = 20twip
        lines = max(1, math.ceil(w_twip / C1_WIDTH_TWIP))
        total_lines += lines
    return total_lines


def calc_row_height(n_lines):
    """計算一列所需高度（包含儲存格上邊距）。"""
    return max(MIN_ROW_H_TWIP, n_lines * LINE_H_TWIP + CELL_TOP_TWIP)


def calc_pair_height(item_text, name_text):
    """
    計算一組資料列（奇偶列）的總列高。

    奇數列 = 項次（通常 1 行），偶數列 = 名稱（依內容動態）。
    每列高度 = max(最小列高, 行數 × 行高 + 儲存格上邊距)
    """
    odd_lines = calc_name_lines(item_text)
    even_lines = calc_name_lines(name_text)
    h_odd = calc_row_height(odd_lines)
    h_even = calc_row_height(even_lines)
    return h_odd + h_even


def load_price_sheet(path, exclude_units=None):
    """
    讀取詳細價目表，回傳 (item, name, qty_unit) 列表。

    規則：
    - 從壹.三.1 之後開始抓取
    - 項次須以「壹.」開頭
    - 排除單位在 exclude_units 中的項目
    - 排除名稱含小計/合計/總價之項目
    - ⑤→窗（修正 PDF 轉檔錯誤）
    """
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
    """
    設定資料列的 trHeight（twip），hRule=atLeast。

    範本使用 val=226 (twip) + hRule=None (預設 atLeast)，
    此處直接寫入 twip 值以保持一致。
    """
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = etree.SubElement(tr, qn('w:trPr'))
        tr.insert(0, trPr)
    trH = trPr.find(qn('w:trHeight'))
    if trH is None:
        trH = etree.SubElement(trPr, qn('w:trHeight'))
    trH.set(qn('w:val'), str(h_twip))
    trH.set(qn('w:hRule'), 'exact')


def set_tc_text(tc, text):
    """
    清除 tc 所有段落後寫入指定文字，並設定字型為標楷體 10pt。

    段落行距繼承自範本預設（與校準測試檔一致），不在此指定。
    列屬性（trHeight atLeast）由範本 trHeight 控制。
    """
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    p = etree.SubElement(tc, qn('w:p'))
    # 段落屬性：行距 = 固定 240 twip（12pt，標準單行行距）
    pPr = etree.SubElement(p, qn('w:pPr'))
    spacing = etree.SubElement(pPr, qn('w:spacing'))
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'exact')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    # 執行屬性：標楷體 10pt
    r = etree.SubElement(p, qn('w:r'))
    rPr = etree.SubElement(r, qn('w:rPr'))
    rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), 'DFKai-SB')
    rFonts.set(qn('w:eastAsia'), 'DFKai-SB')
    rFonts.set(qn('w:hAnsi'), 'DFKai-SB')
    sz = etree.SubElement(rPr, qn('w:sz'))
    sz.set(qn('w:val'), '20')  # 10pt
    szCs = etree.SubElement(rPr, qn('w:szCs'))
    szCs.set(qn('w:val'), '20')
    t = etree.SubElement(r, qn('w:t'))
    t.text = text
    t.set(qn('xml:space'), 'preserve')


def set_vmerge(tc, restart=True):
    """設定 tc 垂直合併屬性（restart / continue）"""
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
    """移除表格中所有垂直合併屬性"""
    for row in table.rows:
        tr = row._tr
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                vm = tcPr.find(qn('w:vMerge'))
                if vm is not None:
                    tcPr.remove(vm)


def add_empty_rows(table, n_pairs, template_row_index=2):
    """為表格新增 n_pairs 組空白資料列（以 template_row_index 為樣板）"""
    if n_pairs <= 0:
        return
    template_tr = table.rows[template_row_index]._tr
    for _ in range(n_pairs * 2):
        table._tbl.append(deepcopy(template_tr))


def apply_header_merges(t0, merge_cols, r0_tc_map):
    """
    對標題列 R0/R1 套用垂直合併。

    R0 因 C7~C12 共用 tc[7]（gridSpan=6），需經 r0_tc_map 轉換。
    """
    r0_tcs = list(t0.rows[0]._tr.findall(qn('w:tc')))
    r1_tcs = list(t0.rows[1]._tr.findall(qn('w:tc')))

    # R0：合併欄位對應到唯一 tc 索引（去重）
    r0_done = set()
    for ci in merge_cols:
        ti = r0_tc_map[ci]
        if ti not in r0_done:
            set_vmerge(r0_tcs[ti], restart=True)
            r0_done.add(ti)

    # R1：合併欄位直接對應 tc 索引
    for ci in merge_cols:
        set_vmerge(r1_tcs[ci], restart=False)


def fill_data_pair(trs, ri, seq, item, name, qty_unit, merge_cols,
                   odd_h=None, even_h=None):
    """
    填入一組資料（奇偶列），對 merge_cols 套用垂直合併。

    合併欄位：奇數列 restart，偶數列 continue
    C1、C4、C5、C13 不併（獨立列）

    若 odd_h / even_h 有提供，設定各列 trHeight（atLeast）。
    """
    r1_tr = trs[ri]
    r2_tr = trs[ri + 1]
    r1_tcs = r1_tr.findall(qn('w:tc'))
    r2_tcs = r2_tr.findall(qn('w:tc'))

    # C1：奇=項次、偶=名稱（不併）
    set_tc_text(r1_tcs[1], item)
    set_tc_text(r2_tcs[1], name)

    # C4、C5、C13（不併，清空）
    for ci in [4, 5, 13]:
        set_tc_text(r1_tcs[ci], '')
        set_tc_text(r2_tcs[ci], '')

    # 合併欄位
    for ci in merge_cols:
        set_vmerge(r1_tcs[ci], restart=True)
        set_vmerge(r2_tcs[ci], restart=False)

    # C0（流水號）
    set_tc_text(r1_tcs[0], str(seq))
    set_tc_text(r2_tcs[0], str(seq))

    # C2（數量+單位）
    set_tc_text(r1_tcs[2], qty_unit)
    set_tc_text(r2_tcs[2], qty_unit)

    # 其他合併欄位清空
    for ci in [3, 6] + list(range(7, 13)) + [14]:
        set_tc_text(r1_tcs[ci], '')
        set_tc_text(r2_tcs[ci], '')

    if odd_h is not None:
        set_tr_height(r1_tr, odd_h)
    if even_h is not None:
        set_tr_height(r2_tr, even_h)


def add_page_break(body):
    """在 body 末尾插入分頁符號"""
    pb_p = etree.SubElement(body, qn('w:p'))
    pb_r = etree.SubElement(pb_p, qn('w:r'))
    pb_br = etree.SubElement(pb_r, qn('w:br'))
    pb_br.set(qn('w:type'), 'page')


def table_has_data(tbl_xml):
    """檢查表格 XML 是否有資料"""
    for tr in tbl_xml.findall(qn('w:tr')):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) > 0:
            txt = ''.join(t.text or '' for t in tcs[0].iter(qn('w:t'))).strip()
            if txt.replace(' ', '').isdigit():
                return True
    return False


def convert(
    price_path,
    template_path,
    output_path,
    exclude_units=None,
    max_pairs=20,
):
    """
    主轉換程式（動態分頁）。

    依 Pillow 測量各材料名稱字寬，搭配校準行高（174+4 twip/行），
    動態計算每頁可容納的資料列對數，確保不跨頁。
    """
    # 價目表資料
    items = load_price_sheet(price_path, exclude_units)
    print(f'價目表載入：{len(items)} 項')

    # 預先計算每組資料列所需高度及奇偶列高
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

    # ---- 標題合併 ----
    R0_MAP = {
        0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6,
        7: 7, 8: 7, 9: 7, 10: 7, 11: 7, 12: 7, 13: 8, 14: 9,
    }
    HDR_MERGE = {0, 2, 3, 6, 14}
    apply_header_merges(t0, HDR_MERGE, R0_MAP)

    # ---- 資料合併欄位 ----
    DAT_MERGE = {0, 2, 3, 6, 7, 8, 9, 10, 11, 12, 14}

    # 擴增空白資料列（大量，以頁面高度為唯一限制）
    cur_pairs = (len(t0.rows) - 2) // 2
    need = 50 - cur_pairs  # 50對 = 100列，應足夠填滿一頁
    if need > 0:
        add_empty_rows(t0, need)

    # 儲存主模板 XML
    master_xml = deepcopy(t0._tbl)

    # 清空文件，準備重建（保留 sectPr）
    body = doc.element.body
    old_sectPr = body.find(qn('w:sectPr'))
    for child in list(body):
        body.remove(child)

    # ---- 多頁動態分頁 ----
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
        description='材料設備送審管制總表轉換工具'
    )
    parser.add_argument(
        '-p', '--price',
        default='./02_成德-詳細價目表.xlsx',
        help='詳細價目表 Excel 路徑',
    )
    parser.add_argument(
        '-t', '--template',
        default='./表5.1.docx',
        help='表5 模板 docx 路徑',
    )
    parser.add_argument(
        '-o', '--output',
        default='../Vbaoffice/表5_完成10.docx',
        help='輸出 docx 路徑',
    )
    parser.add_argument(
        '--exclude-units',
        nargs='+',
        default=['式', '工'],
        help='排除單位（預設：式 工）',
    )
    parser.add_argument(
        '--max-pairs',
        type=int,
        default=20,
        help='每頁最多資料組數（預設：20）',
    )
    args = parser.parse_args()

    # 路徑基準：以本檔所在目錄為準
    base = os.path.dirname(os.path.abspath(__file__))
    price_path = os.path.normpath(os.path.join(base, args.price))
    template_path = os.path.normpath(os.path.join(base, args.template))
    output_path = os.path.normpath(os.path.join(base, args.output))

    for f, label in [(price_path, '價目表'), (template_path, '模板')]:
        if not os.path.isfile(f):
            print(f'錯誤：{label} {f} 不存在')
            sys.exit(1)

    convert(
        price_path,
        template_path,
        output_path,
        exclude_units=set(args.exclude_units),
    )


if __name__ == '__main__':
    main()
