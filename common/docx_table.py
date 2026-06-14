# -*- coding: utf-8 -*-
"""
共用表格元件 — 建立 Word 表格儲存格格式（黑實線框、段落間距 0、最小行高 0pt）

修正歷程：
  v1.0  2026/06/14  初始版本

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/14

相依套件：
  - python-docx>=1.0.0,<2.0.0
  - lxml>=4.9.0,<6.0.0

功能說明：
  提供跨表格轉換器共用之儲存格建立函式：
  - 黑實線框（single, sz=8, color=000000）
  - 段落間距前後 0、行距最小行高 0pt
  - 內文置中或靠左，標楷體 11pt
"""
from docx.oxml.ns import qn
from lxml import etree


def set_cell_border(tcPr):
    """在 tcPr 中加入四邊黑實線框 (single, 1pt, 黑色)"""
    tcBorders = etree.SubElement(tcPr, qn('w:tcBorders'))
    for edge in ('top', 'left', 'bottom', 'right'):
        be = etree.SubElement(tcBorders, qn('w:' + edge))
        be.set(qn('w:val'), 'single')
        be.set(qn('w:sz'), '8')
        be.set(qn('w:space'), '0')
        be.set(qn('w:color'), '000000')


def set_cell_paragraph_spacing(pPr, line_twip=None):
    """在 pPr 中設定段落間距（前後 0），不指定行距讓 Word 自動計算"""
    sp = etree.SubElement(pPr, qn('w:spacing'))
    sp.set(qn('w:before'), '0')
    sp.set(qn('w:after'), '0')
    if line_twip is not None:
        sp.set(qn('w:line'), str(line_twip))
        sp.set(qn('w:lineRule'), 'exact')


def add_cell(tr, text, col_width, center=False, bold=False, font_size=22,
             merge_restart=None, line_twip=None):
    """在列中加入格式化儲存格（黑實線框、段落間距 0、最小行高 0pt）

    Parameters
    ----------
    tr : lxml Element (w:tr)
        所屬表格列
    text : str
        文字內容（空字串則只建立空白儲存格）
    col_width : int
        欄寬（twip）
    center : bool
        是否水平置中（預設靠左）
    bold : bool
        是否粗體
    font_size : int
        字級（半點，預設 22 = 11pt）
    merge_restart : bool or None
        None = 不垂直合併
        True  = vMerge restart（該欄第一列）
        False = vMerge continue（後續列）
    """
    tc = etree.SubElement(tr, qn('w:tc'))
    tcPr = etree.SubElement(tc, qn('w:tcPr'))
    tcW = etree.SubElement(tcPr, qn('w:tcW'))
    tcW.set(qn('w:w'), str(col_width))
    tcW.set(qn('w:type'), 'dxa')

    if merge_restart is not None:
        vm = etree.SubElement(tcPr, qn('w:vMerge'))
        if merge_restart:
            vm.set(qn('w:val'), 'restart')
    # 明確設定儲存格邊距為 0，避免 Word 預設值吃空間
    tcMar = etree.SubElement(tcPr, qn('w:tcMar'))
    for edge, val in [('top', '0'), ('bottom', '0'), ('left', '0'), ('right', '0')]:
        m = etree.SubElement(tcMar, qn('w:' + edge))
        m.set(qn('w:w'), val)
        m.set(qn('w:type'), 'dxa')
    vAlign = etree.SubElement(tcPr, qn('w:vAlign'))
    vAlign.set(qn('w:val'), 'center')

    set_cell_border(tcPr)

    p = etree.SubElement(tc, qn('w:p'))
    pPr = etree.SubElement(p, qn('w:pPr'))
    jc = etree.SubElement(pPr, qn('w:jc'))
    jc.set(qn('w:val'), 'center' if center else 'left')

    set_cell_paragraph_spacing(pPr, line_twip or 240)

    if text:
        r = etree.SubElement(p, qn('w:r'))
        rPr = etree.SubElement(r, qn('w:rPr'))
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:ascii'), '標楷體')
        rFonts.set(qn('w:eastAsia'), '標楷體')
        if bold:
            etree.SubElement(rPr, qn('w:b'))
        sz = etree.SubElement(rPr, qn('w:sz'))
        sz.set(qn('w:val'), str(font_size))
        szCs = etree.SubElement(rPr, qn('w:szCs'))
        szCs.set(qn('w:val'), str(font_size))
        t = etree.SubElement(r, qn('w:t'))
        t.text = text
        t.set(qn('xml:space'), 'preserve')
