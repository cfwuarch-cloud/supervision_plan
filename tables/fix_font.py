# -*- coding: utf-8 -*-
"""
Word 全域修改字型 + 字高工具
============================
遍歷 docx 中所有 run，將字型與字高統一為指定值。

使用方法：
  python fix_font.py -i 輸入.docx -o 輸出.docx
  python fix_font.py -i 輸入.docx -o 輸出.docx --font-name 標楷體 --font-size 11

參數：
  -i, --input     輸入 docx 路徑（必要）
  -o, --output    輸出 docx 路徑（必要）
  -f, --font-name 字型名稱（預設：微軟正黑體）
  -s, --font-size 字型大小（pt，預設：12）
"""
import argparse
import os
import sys
import docx
from docx.oxml.ns import qn


def main():
    parser = argparse.ArgumentParser(description='Word 全域修改字型 + 字高')
    parser.add_argument('-i', '--input', required=True, help='輸入 docx 路徑')
    parser.add_argument('-o', '--output', required=True, help='輸出 docx 路徑')
    parser.add_argument('-f', '--font-name', default='微軟正黑體', help='字型名稱（預設：微軟正黑體）')
    parser.add_argument('-s', '--font-size', type=int, default=12, help='字型大小 pt（預設：12）')

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f'錯誤：輸入檔案不存在 {args.input}')
        sys.exit(1)

    sz_val = str(args.font_size * 2)
    doc = docx.Document(args.input)
    changed = 0

    for p in doc.element.body.iter(qn('w:p')):
        for r in p.iter(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                continue
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                continue
            rFonts.set(qn('w:ascii'), args.font_name)
            rFonts.set(qn('w:hAnsi'), args.font_name)
            rFonts.set(qn('w:eastAsia'), args.font_name)

            sz = rPr.find(qn('w:sz'))
            if sz is None:
                continue
            sz.set(qn('w:val'), sz_val)
            szCs = rPr.find(qn('w:szCs'))
            if szCs is not None:
                szCs.set(qn('w:val'), sz_val)

            changed += 1

    doc.save(args.output)
    print(f'完成：修改 {changed} 個 run → 字型={args.font_name}，字高={args.font_size}pt')


if __name__ == '__main__':
    main()