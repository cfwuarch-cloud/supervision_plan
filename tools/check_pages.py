"""
檢查 docx 檔案頁數（估算 + Word COM 自動化）
用法: python check_pages.py 檔案路徑.docx

若 win32com 可用，直接開 Word 讀取實際頁數（最準確）。
否則以 trHeight 總和 / 頁面可用空間 估算頁數。
"""
import sys
import zipfile
from lxml import etree
import os


def estimate_pages(docx_path):
    """從 XML 資料估算檔案的 Word 頁數"""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    try:
        with zipfile.ZipFile(docx_path) as z:
            root = etree.fromstring(z.read('word/document.xml'))
    except FileNotFoundError:
        print(f'錯誤：找不到檔案 {docx_path}')
        return None

    body = root.find(f'{{{ns}}}body')

    # 讀取 sectPr（頁面設定）
    sectPrs = body.findall(f'{{{ns}}}sectPr')
    if not sectPrs:
        print('⚠ 文件無 sectPr 頁面設定，無法計算')
        return None

    sp = sectPrs[-1]
    pgSz = sp.find(f'{{{ns}}}pgSz')
    pgMar = sp.find(f'{{{ns}}}pgMar')
    if pgSz is None or pgMar is None:
        print('⚠ 頁面尺寸或邊距不完整')
        return None

    page_h = int(pgSz.get(f'{{{ns}}}h'))
    top_m = int(pgMar.get(f'{{{ns}}}top'))
    bot_m = int(pgMar.get(f'{{{ns}}}bottom'))
    avail_h = page_h - top_m - bot_m

    # 計算表格高度
    tbls = body.findall(f'{{{ns}}}tbl')
    if not tbls:
        print('⚠ 文件無表格')
        return None

    rows = tbls[0].findall(f'{{{ns}}}tr')
    total_h = 0
    for tr in rows:
        trPr = tr.find(f'{{{ns}}}trPr')
        if trPr is not None:
            trH = trPr.find(f'{{{ns}}}trHeight')
            if trH is not None:
                total_h += int(trH.get(f'{{{ns}}}val'))

    # 檢查是否有分頁符號
    page_breaks = 0
    for br in body.iter(f'{{{ns}}}br'):
        if br.get(f'{{{ns}}}type') == 'page':
            page_breaks += 1

    # 估算頁數
    if total_h <= avail_h:
        est_pages = 1
    else:
        est_pages = (total_h + avail_h - 1) // avail_h

    print(f'    頁面可用高度: {avail_h} twip')
    print(f'    表格 trHeight 總和: {total_h} twip')
    print(f'    分頁符號數量: {page_breaks}')
    print(f'    估算頁數: {est_pages}')

    if page_breaks > 0:
        print(f'    【注意】文件含 {page_breaks} 個分頁符號')
    if total_h > avail_h:
        print(f'    【注意】trHeight 總和超出頁面 {(total_h - avail_h)} twip')

    return est_pages


def try_word_com(docx_path):
    """嘗試用 Word COM 自動化讀取實際頁數"""
    try:
        import win32com.client
    except ImportError:
        return None

    word = None
    try:
        abs_path = os.path.abspath(docx_path)
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(abs_path)
        pages = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
        doc.Close()
        return pages
    except Exception as e:
        print(f'    Word COM 錯誤: {e}')
        return None
    finally:
        if word is not None:
            word.Quit()


def main():
    if len(sys.argv) < 2:
        print('用法: python check_pages.py 檔案路徑.docx')
        return

    path = sys.argv[1]
    if not os.path.exists(path):
        print(f'錯誤：找不到檔案 {path}')
        return

    print(f'檔案: {path}')
    print()

    # 先試 Word COM
    print('[1] 嘗試 Word COM 自動化...')
    actual = try_word_com(path)
    if actual is not None:
        print(f'    Word 實際頁數: {actual}')
        return

    # 再用估算
    print('[2] win32com 不可用，改以 XML 資料估算...')
    estimate_pages(path)


if __name__ == '__main__':
    main()
