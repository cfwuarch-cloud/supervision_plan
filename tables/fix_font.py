# -*- coding: utf-8 -*-
"""
Word 全域修改字型 / 字高工具（修正版）
============================
遍歷 docx 中所有的 run（包含表格與頁首頁尾），全域修改字型與字高。

使用方法：
  python fix_font.py -i 輸入.docx -o 輸出.docx -f 微軟正黑體
  python fix_font.py -i 輸入.docx -o 輸出.docx -s 12
  python fix_font.py -i 輸入.docx -o 輸出.docx -s 10.5
  python fix_font.py -i 輸入.docx -o 輸出.docx -f 標楷體 -s 11

參數：
  -i, --input     輸入 docx 路徑（必要）
  -o, --output    輸出 docx 路徑（必要）
  -f, --font-name 字型名稱（省略則不改字型）
  -s, --font-size 字型大小 pt，支援小數（省略則不改字高）
"""
import argparse
import os
import shutil
import sys
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def get_or_add_child(parent, tag_name):
    child = parent.find(qn(tag_name))
    if child is None:
        child = OxmlElement(tag_name)
        parent.append(child)
    return child


def main():
    parser = argparse.ArgumentParser(description='Word 全域修改字型 / 字高')
    parser.add_argument('-i', '--input', required=True, help='輸入 docx 路徑')
    parser.add_argument('-o', '--output', required=True, help='輸出 docx 路徑')
    parser.add_argument('-f', '--font-name', default=None, help='字型名稱（省略則不改字型）')
    parser.add_argument('-s', '--font-size', type=float, default=None, help='字型大小 pt，支援小數（省略則不改字高）')

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f'錯誤：輸入檔案不存在 {args.input}')
        sys.exit(1)

    if args.font_name is None and args.font_size is None:
        print('警告：未指定任何修改項目（-f 或 -s），直接複製檔案')
        shutil.copy2(args.input, args.output)
        print('完成')
        return

    doc = docx.Document(args.input)
    changed = 0
    sz_val = str(int(args.font_size * 2)) if args.font_size is not None else None

    # 走訪整份文件 XML（涵蓋 Body、Table、Header/Footer）
    for r in doc.element.iter(qn('w:r')):
        rPr = get_or_add_child(r, 'w:rPr')

        if args.font_name is not None:
            rFonts = get_or_add_child(rPr, 'w:rFonts')
            rFonts.set(qn('w:ascii'), args.font_name)
            rFonts.set(qn('w:hAnsi'), args.font_name)
            rFonts.set(qn('w:eastAsia'), args.font_name)
            rFonts.set(qn('w:cs'), args.font_name)

        if args.font_size is not None:
            sz = get_or_add_child(rPr, 'w:sz')
            sz.set(qn('w:val'), sz_val)
            szCs = get_or_add_child(rPr, 'w:szCs')
            szCs.set(qn('w:val'), sz_val)

        changed += 1

    doc.save(args.output)

    parts = []
    if args.font_name:
        parts.append(f'字型={args.font_name}')
    if args.font_size:
        parts.append(f'字高={args.font_size}pt')
    print(f'完成：修改 {changed} 個 run → {"，".join(parts)}')


if __name__ == '__main__':
    main()