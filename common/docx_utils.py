# -*- coding: utf-8 -*-
"""
docx 共用工具函式

修正歷程：
  v1.0  2026/06/14  初始版本：擷取各轉換器重複之 docx 操作

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/14

相依套件：
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0

功能說明：
  提供各表格轉換器共用之 docx XML 操作（設定列高、合併、分頁等）

參數說明：
  無（純函式庫）
"""
from docx.oxml.ns import qn
from lxml import etree
from copy import deepcopy


def set_tr_height(tr, h_twip):
    """設定資料列的 trHeight（twip），hRule=exact"""
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = etree.SubElement(tr, qn('w:trPr'))
        tr.insert(0, trPr)
    trH = trPr.find(qn('w:trHeight'))
    if trH is None:
        trH = etree.SubElement(trPr, qn('w:trHeight'))
    trH.set(qn('w:val'), str(h_twip))
    trH.set(qn('w:hRule'), 'atLeast')


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


def add_page_break(body):
    """在 body 末尾插入分頁符號"""
    pb_p = etree.SubElement(body, qn('w:p'))
    pb_r = etree.SubElement(pb_p, qn('w:r'))
    pb_br = etree.SubElement(pb_r, qn('w:br'))
    pb_br.set(qn('w:type'), 'page')
