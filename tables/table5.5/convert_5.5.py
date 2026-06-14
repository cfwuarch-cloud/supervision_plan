# -*- coding: utf-8 -*-
"""
不合格改善追蹤表轉換工具
======================
將 詳細價目表.xlsx 作為母本，填入 表5.5.docx 模板，
每項材料產出一份不合格改善追蹤表。

修正歷程：
  v1.0  2026/06/14  初始版本

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/14

相依套件：
  - openpyxl>=3.0.0,<4.0.0
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0
  - pandas>=1.5.0,<3.0.0

使用方法：
  python -X utf8 tables/table5.5/convert_5.5.py --exclude-units 式 工

功能說明：
  1. 讀取詳細價目表，每項材料產出一份不合格改善追蹤表
  2. 填入工程名稱、監造單位、材料名稱
  3. 其餘欄位保留空白（待現場填寫）

參數說明：
  -p, --price   詳細價目表 Excel 路徑（預設：../../data/02_成德-詳細價目表.xlsx）
  -t, --template  表5.5模板 docx 路徑（預設：./表5.5.docx）
  -o, --output   輸出 docx 路徑（預設：../../output/表5.5_完成.docx）
  --exclude-units  排除單位（預設：式 工）
  --test-num  測試流水號
  --max-items  最多產出項數（預設：0=不限）
"""
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from copy import deepcopy
import argparse
import os
import sys
import re


PROJECT_NAME = '臺南市政府社會局委託辦理北區成德公設民營托嬰中心室內裝修統包工程'
SUPERVISOR = '吳真福建築師事務所'


def clean_text(text):
    text = text.replace('\n', '').replace('\r', '')
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[\s]+([、，,。．])', r'\1', text)
    return text.strip()


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


def fill_tc_text(tr, col_idx, new_text, ns):
    tcs = tr.findall(ns + 'tc')
    if col_idx >= len(tcs):
        return False
    tc = tcs[col_idx]
    t_elements = [t for t in tc.iter(ns + 't')]
    for t in t_elements:
        t.text = ''
    if t_elements:
        t_elements[0].text = new_text
        return True
    p = tc.find(ns + 'p')
    if p is not None:
        from lxml import etree
        r = etree.SubElement(p, qn('w:r'))
        etree.SubElement(r, qn('w:t')).text = new_text
        return True
    return False


def main():
    parser = argparse.ArgumentParser(description='不合格改善追蹤表轉換工具')
    parser.add_argument('-p', '--price', default='../../data/02_成德-詳細價目表.xlsx')
    parser.add_argument('-t', '--template', default='./表5.5.docx')
    parser.add_argument('-o', '--output', default='../../output/表5.5_完成.docx')
    parser.add_argument('--exclude-units', nargs='+', default=['式', '工'])
    parser.add_argument('--test-num', type=int)
    parser.add_argument('--max-items', type=int, default=0)
    args = parser.parse_args()

    base = os.path.dirname(os.path.abspath(__file__))
    price_path = os.path.join(base, args.price) if not os.path.isabs(args.price) else args.price
    template_path = os.path.join(base, args.template) if not os.path.isabs(args.template) else args.template
    output_path = os.path.join(base, args.output) if not os.path.isabs(args.output) else args.output

    if args.test_num:
        d = os.path.dirname(output_path)
        bn = os.path.splitext(os.path.basename(output_path))[0]
        output_path = os.path.join(d, f'{bn}_test_{args.test_num}.docx')

    items = load_price_sheet(price_path, set(args.exclude_units))
    print(f'價目表載入：{len(items)} 項')
    if args.max_items > 0:
        items = items[:args.max_items]

    doc = Document(template_path)
    body = doc.element.body
    children = list(body)
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    sect = body.find(qn('w:sectPr'))

    tbl_el = None
    before = []
    after = []
    seen_tbl = False
    for child in children:
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == 'tbl':
            tbl_el = child
            seen_tbl = True
        elif tag == 'p' and not seen_tbl:
            before.append(child)
        elif tag == 'p' and seen_tbl:
            after.append(child)

    for child in list(body):
        body.remove(child)

    total = len(items)
    for idx, (item, name, qty_unit) in enumerate(items):
        if idx > 0:
            pb_p = etree.SubElement(body, qn('w:p'))
            pb_r = etree.SubElement(pb_p, qn('w:r'))
            pb_br = etree.SubElement(pb_r, qn('w:br'))
            pb_br.set(qn('w:type'), 'page')

        for p in before:
            body.append(deepcopy(p))

        new_tbl = deepcopy(tbl_el)
        rows = new_tbl.findall(ns + 'tr')

        # Row 0: project name (col 1)
        if len(rows) > 0:
            fill_tc_text(rows[0], 1, PROJECT_NAME, ns)

        # Row 1: supervisor (col 1)
        if len(rows) > 1:
            fill_tc_text(rows[1], 1, SUPERVISOR, ns)

        # Row 3: material name (col 0, merged cell contains "材料名稱：")
        if len(rows) > 3:
            fill_tc_text(rows[3], 0, f'材料名稱：{name}', ns)

        body.append(new_tbl)

        for p in after:
            body.append(deepcopy(p))

    if sect is not None:
        body.append(deepcopy(sect))

    doc.save(output_path)
    print(f'已完成：{total} 項 → {output_path}')
    return total


if __name__ == '__main__':
    main()
