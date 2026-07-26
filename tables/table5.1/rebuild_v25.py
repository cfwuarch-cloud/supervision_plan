# -*- coding: utf-8 -*-
"""表5-1 V2.5 — 以表5.1.docx 為模板，表格從零建立
"""
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy
import re

TEMPLATE = r'C:\OpenCodeLib\supervision_plan\tables\table5.1\表5.1.docx'
DATA_SRC = r'C:\OpenCodeLib\supervision_plan\output\(安南托育)表5-1_材料設備送審管制總表V2.docx'
OUT = r'C:\OpenCodeLib\supervision_plan\output\(安南托育)表5-1_材料設備送審管制總表V2.5.docx'

PAGE_PLAN = [12, 11, 10, 10, 12, 11, 10, 10, 12, 13, 7]
WITHIN = [0, 2, 6, 7, 8, 9, 10, 11, 12, 14]
RW = ['137','1353','357','241','566','262','210','234','191','220','195','208','161','496','169']

# ===== XML helpers =====
def T(w, g=1):
    tc = OxmlElement('w:tc'); p = OxmlElement('w:tcPr')
    e = OxmlElement('w:tcW'); e.set(qn('w:w'),w); e.set(qn('w:type'),'dxa'); p.append(e)
    if g>1:
        e = OxmlElement('w:gridSpan'); e.set(qn('w:val'),str(g)); p.append(e)
    e = OxmlElement('w:vAlign'); e.set(qn('w:val'),'center'); p.append(e)
    m = OxmlElement('w:tcMar')
    for s in ['top','left','bottom','right']:
        x = OxmlElement(f'w:{s}'); x.set(qn('w:w'),'0'); x.set(qn('w:type'),'dxa'); m.append(x)
    p.append(m); tc.append(p); return tc

def P(tc, txt=''):
    p = OxmlElement('w:p'); pp = OxmlElement('w:pPr')
    s = OxmlElement('w:spacing'); s.set(qn('w:before'),'0'); s.set(qn('w:after'),'0')
    s.set(qn('w:line'),'240'); s.set(qn('w:lineRule'),'exact'); pp.append(s)
    g = OxmlElement('w:snapToGrid'); g.set(qn('w:val'),'0'); pp.append(g)
    j = OxmlElement('w:jc'); j.set(qn('w:val'),'center'); pp.append(j); p.append(pp)
    r = OxmlElement('w:r'); rp = OxmlElement('w:rPr')
    f = OxmlElement('w:rFonts'); f.set(qn('w:ascii'),'標楷體'); f.set(qn('w:eastAsia'),'標楷體'); rp.append(f)
    for t in ['w:sz','w:szCs']:
        x = OxmlElement(t); x.set(qn('w:val'),'20'); rp.append(x)
    r.append(rp); t = OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text=txt; r.append(t)
    p.append(r); tc.append(p)

def TR():
    tr = OxmlElement('w:tr'); tp = OxmlElement('w:trPr')
    th = OxmlElement('w:trHeight'); th.set(qn('w:val'),'226'); tp.append(th); tr.append(tp); return tr

def VM(tc, mode='restart'):
    p = tc.find(qn('w:tcPr')); v = OxmlElement('w:vMerge')
    if mode=='restart': v.set(qn('w:val'),'restart'); p.append(v)

def txt(tr, ci):
    tcs = tr.findall(qn('w:tc'))
    if ci < len(tcs): return ''.join(t.text or '' for t in tcs[ci].iter(qn('w:t')))
    return ''

# ===== Phase 1: 讀模板 + 原始資料 =====
tmpl = Document(TEMPLATE)
tmpl_body = tmpl.element.body
tmpl_children = list(tmpl_body)

TITLE_PS = [deepcopy(tmpl_children[i]) for i in [0, 1, 2]]
SECT_PR = deepcopy(next(c for c in reversed(tmpl_children) if c.tag == qn('w:sectPr')))

# 從 _0715 取資料
src = Document(DATA_SRC)
src_body = src.element.body
src_tbls = [c for c in list(src_body) if c.tag == qn('w:tbl')]
raw_pairs = []
for ti, tbl_el in enumerate(src_tbls):
    trs = list(tbl_el.findall(qn('w:tr')))
    nh = 3 if ti == 0 else 2
    for i in range(nh, len(trs), 2):
        if i+1 >= len(trs): break
        raw_pairs.append((deepcopy(trs[i]), deepcopy(trs[i+1])))
assert len(raw_pairs) == sum(PAGE_PLAN), f'{len(raw_pairs)} != {sum(PAGE_PLAN)}'

# ===== Phase 2: 建立新文件 =====
doc_out = tmpl  # 以模板為基礎

# 清空 body（保留 sectPr）
for c in list(tmpl_body):
    if c.tag != qn('w:sectPr'):
        tmpl_body.remove(c)

scale = 9825.0 / sum(int(w) for w in RW)
pair_idx = 0

for pg, n_pairs in enumerate(PAGE_PLAN):
    pgn = pg + 1; ttl = len(PAGE_PLAN)

    if pg > 0:
        pb = OxmlElement('w:p')
        r = OxmlElement('w:r'); b = OxmlElement('w:br'); b.set(qn('w:type'),'page')
        r.append(b); pb.append(r); tmpl_body.append(pb)
        tmpl_body.append(OxmlElement('w:p'))

    for pi in range(3):
        p = deepcopy(TITLE_PS[pi])
        runs = [r for r in p.iter(qn('w:r'))]
        if pi == 0:  # 標題
            if len(runs) >= 4:
                t0 = runs[0].find(qn('w:t'))
                if t0 is not None: t0.text = f'表5-1 材料設備送審管制總表-{pgn}'
                for ri in [1,2,3]:
                    t = runs[ri].find(qn('w:t'))
                    if t is not None: t.text = ''
        elif pi == 2:  # 頁碼
            for ri, r in enumerate(runs):
                t = r.find(qn('w:t'))
                if t is None: continue
                tx = t.text or ''
                if tx == '  ' and ri == 5:  # 第 X 頁
                    t.text = str(pgn)
                elif tx == '  ' and ri == 7:  # 共 X 頁
                    t.text = str(ttl)
                elif ri == 15 and tx == ' ':  # E51-X
                    t.text = str(pgn)
        tmpl_body.append(p)

    # --- 建表格 ---
    tbl = OxmlElement('w:tbl')
    tp = OxmlElement('w:tblPr')
    x = OxmlElement('w:tblW'); x.set(qn('w:w'),'5289'); x.set(qn('w:type'),'pct'); tp.append(x)
    for e in ['top','left','bottom','right','insideH','insideV']:
        x = OxmlElement(f'w:{e}'); x.set(qn('w:val'),'single')
        x.set(qn('w:sz'),'4'); x.set(qn('w:space'),'0'); x.set(qn('w:color'),'000000'); tp.append(x)
    x = OxmlElement('w:tblLook'); x.set(qn('w:val'),'0000'); tp.append(x)
    tbl.append(tp)

    gd = OxmlElement('w:tblGrid')
    for w in RW:
        x = OxmlElement('w:gridCol'); x.set(qn('w:w'),str(round(int(w)*scale))); gd.append(x)
    tbl.append(gd)

    # H0
    h0 = TR()
    ts = ['項次','契約詳細價目表項次','契約數量','是否取樣試驗',
          '預定送審日期','是否驗廠','預定試驗單位',
          '送審資料(ˇ)','審查日期','備註(歸檔編號)']
    for ci in range(10):
        t = T(RW[ci] if ci<7 else str(sum(int(RW[i]) for i in range(7,13))) if ci==7 else RW[ci+5])
        if ci==7: t = T(str(sum(int(RW[i]) for i in range(7,13))), 6)
        P(t, ts[ci]); h0.append(t)
    tbl.append(h0)

    if pg == 0:
        h1 = TR()
        for ci in range(15):
            t = T(RW[ci]); P(t,'')
            VM(t, 'continue' if ci not in range(7,14) else 'restart')
            h1.append(t)
        tbl.append(h1)

    hl = TR()
    ls = ['','材料(設備)名稱','','','實際送審日期','驗廠日期','',
          '協力廠商資料','型錄','相關試驗報告','樣品','出廠證明','其他','審查結果','']
    for ci in range(15):
        t = T(RW[ci]); P(t,ls[ci])
        if pg==0: VM(t,'continue')
        hl.append(t)
    tbl.append(hl)

    offset = pair_idx
    for pi in range(n_pairs):
        o_s, e_s = raw_pairs[pair_idx]; seq = offset + pi + 1

        o = TR()
        for ci in range(15):
            t = T(RW[ci])
            tx = str(seq) if ci==0 else txt(o_s,ci)
            if ci==3: tx = '否' if pi==0 else ''
            P(t,tx)
            if ci in WITHIN: VM(t)
            if ci==3: VM(t,'restart' if pi==0 else 'continue')
            o.append(t)
        tbl.append(o)

        e = TR()
        for ci in range(15):
            t = T(RW[ci])
            tx = '' if ci==3 else txt(e_s,ci)
            P(t,tx)
            if ci in WITHIN: VM(t,'continue')
            if ci==3: VM(t,'continue')
            e.append(t)
        tbl.append(e)

        pair_idx += 1

    tmpl_body.append(tbl)

tmpl_body.append(SECT_PR)
doc_out.save(OUT)
print(f'Done! Pages={len(PAGE_PLAN)}, Items={pair_idx}')
print(f'File: {OUT}')
