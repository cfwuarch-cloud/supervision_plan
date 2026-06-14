# -*- coding: utf-8 -*-
"""
材料設備品質抽驗紀錄表轉換工具
=============================
將 詳細價目表.xlsx 作為母本，填入 表5.3.docx 模板，
每項材料產出一份抽驗紀錄表（一頁或多頁）。

修正歷程：
  v1.0  2026/06/14  初始版本：deepcopy 整份表格，填入材料資訊

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/14

相依套件：
  - openpyxl>=3.0.0,<4.0.0
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0
  - pandas>=1.5.0,<3.0.0

使用方法：
  python -X utf8 tables/table5.3/convert_5.3.py --exclude-units 式 工

功能說明：
  1. 讀取詳細價目表，依項次壹.三.1 之後、排除指定單位
  2. 每項材料產出一份空白抽驗紀錄表（所有欄位保留空白待現場填寫）

參數說明：
  -p, --price   詳細價目表 Excel 路徑（預設：../../data/02_成德-詳細價目表.xlsx）
  -t, --template  表5.3模板 docx 路徑（預設：./表5.3.docx）
  -o, --output   輸出 docx 路徑（預設：../../output/表5.3_完成.docx）
  --exclude-units  排除單位（預設：式 工）
  --test-num  測試流水號，輸出檔名自動插入 test_N
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


def clean_text(text):
    text = text.replace('\n', '').replace('\r', '')
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[\s]+([、，,。．])', r'\1', text)
    return text.strip()


def load_price_sheet(path, exclude_units=None):
    """回傳符合篩選條件的項數（不做資料填充用）"""
    if exclude_units is None:
        exclude_units = {'式', '工'}
    df = pd.read_excel(path, sheet_name='Table 1')
    c0 = df.columns[0]
    c1 = df.columns[1]
    cu = df.columns[6]
    cq = df.columns[7]

    count = 0
    started = False
    for _, row in df.iterrows():
        item = clean_text(str(row[c0]))
        name = clean_text(str(row[c1]))
        unit = clean_text(str(row[cu])) if pd.notna(row[cu]) else ''
        qty = row[cq]

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
        count += 1

    return count


def main():
    parser = argparse.ArgumentParser(description='材料設備品質抽驗紀錄表轉換工具')
    parser.add_argument('-p', '--price', default='../../data/02_成德-詳細價目表.xlsx')
    parser.add_argument('-t', '--template', default='./表5.3.docx')
    parser.add_argument('-o', '--output', default='../../output/表5.3_完成.docx')
    parser.add_argument('--exclude-units', nargs='+', default=['式', '工'])
    parser.add_argument('--test-num', type=int)
    parser.add_argument('--max-items', type=int, default=0)
    args = parser.parse_args()

    base = os.path.dirname(os.path.abspath(__file__))
    price_path = os.path.join(base, args.price) if not os.path.isabs(args.price) else args.price
    template_path = os.path.join(base, args.template) if not os.path.isabs(args.template) else args.template
    output_path = os.path.join(base, args.output) if not os.path.isabs(args.output) else args.output

    if args.test_num:
        output_dir = os.path.dirname(output_path)
        base_name = os.path.splitext(os.path.basename(output_path))[0]
        output_path = os.path.join(output_dir, f'{base_name}_test_{args.test_num}.docx')

    n_items = load_price_sheet(price_path, set(args.exclude_units))
    print(f'價目表載入：{n_items} 項')

    if args.max_items > 0:
        n_items = min(n_items, args.max_items)

    doc = Document(template_path)
    body = doc.element.body
    children = list(body)
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

    for idx in range(n_items):
        if idx > 0:
            pb_p = etree.SubElement(body, qn('w:p'))
            pb_r = etree.SubElement(pb_p, qn('w:r'))
            pb_br = etree.SubElement(pb_r, qn('w:br'))
            pb_br.set(qn('w:type'), 'page')

        for p in before:
            body.append(deepcopy(p))
        body.append(deepcopy(tbl_el))
        for p in after:
            body.append(deepcopy(p))

    if sect is not None:
        body.append(deepcopy(sect))

    doc.save(output_path)
    print(f'已完成：{n_items} 份空白表單 → {output_path}')
    return n_items


if __name__ == '__main__':
    main()
