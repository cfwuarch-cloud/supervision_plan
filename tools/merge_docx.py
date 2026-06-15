# -*- coding: utf-8 -*-
"""
監造計畫合併工具 v1.0
======================
將各章節獨立 docx 合併為一本完整監造計畫書。

修正歷程：
  v1.0  2026/06/15  初始版本

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/15

相依套件：
  - python-docx>=1.0.0,<2.0.0

使用方法：
  python -X utf8 tools/merge_docx.py

功能說明：
  依章節順序讀取各 docx 文件，合併為一本完整監造計畫書。
  自動處理分頁符號、頁面設定（以第一章為準）。

參數說明：
  -o, --output     輸出 docx 路徑（預設：../output/監造計畫書_完整版.docx）
  --chapters      自訂章節順序（預設：前言、Ch1~Ch4、Ch6、Ch8、Ch9）
"""
import argparse
import os
import sys
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml


# 預設章節路徑（相對於此腳本目錄的上層）
DEFAULT_CHAPTERS = [
    ('前言', '../output/前言.docx'),
    ('第一章  監造範圍', '../output/Ch1_監造範圍.docx'),
    ('第二章  監造組織及權責分工', '../output/Ch2_監造組織及權責分工.docx'),
    ('第三章  品質計畫審查作業程序', '../output/Ch3_品質計畫審查作業程序.docx'),
    ('第四章  施工計畫審查作業程序', '../output/Ch4_施工計畫審查作業程序.docx'),
    ('第六章  設備功能運轉測試抽驗程序及標準', '../output/Ch6_設備功能運轉測試抽驗程序及標準.docx'),
    ('第八章  品質稽核', '../output/Ch8_品質稽核.docx'),
    ('第九章  文件紀錄管理系統', '../output/Ch9_文件紀錄管理系統.docx'),
]


def merge_docs(chapter_list, output_path):
    """合併多個 docx 文件"""

    target_doc = None

    for chap_name, chap_path in chapter_list:
        if not os.path.isfile(chap_path):
            print(f'  警告：{chap_name} → {chap_path} 不存在，跳過')
            continue

        print(f'  合併：{chap_name}')
        source = Document(chap_path)

        if target_doc is None:
            # 以第一章為目標文件
            target_doc = source
            if target_doc is None:
                print('錯誤：無法建立目標文件')
                return False
            continue

        # 複製 source body 中的所有元素到 target
        source_body = source.element.body
        target_body = target_doc.element.body

        # 加入分頁符號
        add_page_break_element(target_body)

        # 複製所有子元素（跳過 sectPr）
        sect_pr_source = source_body.find(qn('w:sectPr'))
        for child in list(source_body):
            if child.tag == qn('w:sectPr'):
                continue
            target_body.append(deepcopy(child))

        # 保留最後文件的 sectPr（頁面設定）
        if sect_pr_source is not None:
            existing_sectpr = target_body.find(qn('w:sectPr'))
            if existing_sectpr is not None:
                target_body.remove(existing_sectpr)
            target_body.append(deepcopy(sect_pr_source))

    if target_doc is None:
        print('錯誤：沒有有效的章節文件可合併')
        return False

    target_doc.save(output_path)
    total = sum(1 for _, p in chapter_list if os.path.isfile(p))
    print(f'合併完成：{total} 章 → {output_path}')
    return True


def add_page_break_element(body):
    """在 body 末尾插入分頁符號元素"""
    from lxml import etree
    pb_p = etree.SubElement(body, qn('w:p'))
    pb_r = etree.SubElement(pb_p, qn('w:r'))
    pb_br = etree.SubElement(pb_r, qn('w:br'))
    pb_br.set(qn('w:type'), 'page')


def resolve_paths(base_dir, chapters):
    """將相對路徑轉為絕對路徑"""
    resolved = []
    for chap_name, chap_path in chapters:
        abs_path = os.path.normpath(os.path.join(base_dir, chap_path))
        resolved.append((chap_name, abs_path))
    return resolved


def main():
    parser = argparse.ArgumentParser(description='監造計畫合併工具 v1.0')
    parser.add_argument('-o', '--output', default='../output/監造計畫書_完整版.docx',
                        help='輸出 docx 路徑')
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.normpath(os.path.join(base_dir, args.output))

    chapters = resolve_paths(base_dir, DEFAULT_CHAPTERS)

    merge_docs(chapters, output_path)


if __name__ == '__main__':
    main()
