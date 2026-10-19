# -*- coding: utf-8 -*-
"""文档格式转换核心模块

转换策略:
- Office (Word/Excel/PPT) -> PDF : 优先 COM 调用本机 MS Office / WPS, 保真度最高;
                                   失败时用 PyMuPDF Story 做"简易排版"兜底 (仅 Word/Excel)
- PDF -> Word / 图片             : pdf2docx / PyMuPDF 渲染
- 图片 -> PDF, PDF 工具          : PyMuPDF
- CSV / TXT / Excel / Word 互转  : openpyxl / python-docx
"""
import os
import csv
import html
import threading

# ---------------------------------------------------------------- COM 转换

_WORD_PROGIDS = ["Word.Application", "KWPS.Application"]
_EXCEL_PROGIDS = ["Excel.Application", "KET.Application"]
_PPT_PROGIDS = ["PowerPoint.Application", "KWPP.Application"]

_com_lock = threading.Lock()


def _com_to_pdf(src, dst, progids, page_setup=None):
    """通过 COM (MS Office 或 WPS) 将 office 文档导出为 PDF, 单线程串行调用
    page_setup: (landscape, fit_wide), 仅对 Excel 生效
      landscape: None=跟随文档 / True=横向 / False=纵向
      fit_wide: True=所有列缩放至一页宽 (宽表格推荐)"""
    last = "未知错误"
    with _com_lock:
        for progid in progids:
            app = None
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                app = win32com.client.DispatchEx(progid)
                app.Visible = False
                for attr, val in (("DisplayAlerts", 0), ("DisplayAlerts", False)):
                    try:
                        setattr(app, attr, val)
                        break
                    except Exception:
                        continue
                s, d = os.path.abspath(src), os.path.abspath(dst)
                if "Word" in progid or "KWPS" in progid:
                    doc = app.Documents.Open(s, ReadOnly=True)
                    doc.ExportAsFixedFormat(d, 17)          # wdExportFormatPDF
                    doc.Close(False)
                elif "Excel" in progid or "KET" in progid:
                    wb = app.Workbooks.Open(s, ReadOnly=True)
                    if page_setup and (page_setup[0] is not None or page_setup[1]):
                        landscape, fit_wide = page_setup
                        for ws in wb.Worksheets:
                            try:
                                ps = ws.PageSetup
                                if landscape is not None:
                                    ps.Orientation = 2 if landscape else 1  # 2=横向
                                if fit_wide:
                                    ps.Zoom = False
                                    ps.FitToPagesWide = 1   # 所有列缩放至一页宽
                                    ps.FitToPagesTall = False  # 高度自动分页
                            except Exception:
                                continue
                    wb.ExportAsFixedFormat(0, d)            # xlTypePDF
                    wb.Close(False)
                else:  # PowerPoint / KWPP
                    pres = app.Presentations.Open(s, ReadOnly=True, WithWindow=False)
                    pres.SaveAs(d, 32)                       # ppSaveAsPDF
                    pres.Close()
                if os.path.isfile(dst) and os.path.getsize(dst) > 0:
                    return True, f"使用 {progid} 导出成功"
                last = f"{progid} 未生成文件"
            except Exception as e:
                last = f"{progid}: {e}"
            finally:
                try:
                    if app is not None:
                        app.Quit()
                except Exception:
                    pass
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    return False, f"本机未找到可用的 Office/WPS 组件 ({last})"


# ---------------------------------------------------------------- PyMuPDF 兜底

def _story_html_to_pdf(html_text, dst, fontname="china-s"):
    """用 PyMuPDF Story 把简易 HTML 排版成 PDF (兜底方案, 样式有损)"""
    import pymupdf
    writer = pymupdf.DocumentWriter(dst)
    story = pymupdf.Story(html=html_text, archive=None)
    mediabox = pymupdf.paper_rect("a4")
    where = mediabox + (36, 36, -36, -36)
    more = 1
    while more:
        dev = writer.begin_page(mediabox)
        more, filled = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()


def _docx_text_fallback_pdf(src, dst):
    """无 Office 时的 Word->PDF 兜底: 提取文字按简易样式重排"""
    from docx import Document
    doc = Document(src)
    parts = ["<h2>Word 转 PDF（简易排版）</h2>"]
    for p in doc.paragraphs:
        t = html.escape(p.text.strip())
        if not t:
            parts.append("<br/>")
        elif p.style.name.lower().startswith(("heading", "标题")):
            parts.append(f"<h3>{t}</h3>")
        else:
            parts.append(f"<p>{t}</p>")
    for table in doc.tables:
        rows = ["<tr>" + "".join(f"<td>{html.escape(c.text)}</td>" for c in row.cells) + "</tr>"
                for row in table.rows]
        parts.append("<table border='1'>" + "".join(rows) + "</table>")
    _story_html_to_pdf("".join(parts), dst)
    return True, "未检测到 Office/WPS, 已用内置引擎简易排版(样式有损)"


def _xlsx_text_fallback_pdf(src, dst, landscape=None):
    """无 Office 时的 Excel->PDF 兜底: 表格转 HTML 再排版 (landscape=True 为横向)"""
    import openpyxl
    import pymupdf
    wb = openpyxl.load_workbook(src, data_only=True)
    parts = []
    for ws in wb.worksheets:
        parts.append(f"<h3>{html.escape(ws.title)}</h3><table border='1' cellspacing='0'>")
        for row in ws.iter_rows(max_row=500, max_col=30, values_only=True):
            cells = "".join(f"<td>{html.escape('' if v is None else str(v))}</td>" for v in row)
            parts.append(f"<tr>{cells}</tr>")
        parts.append("</table>")
    writer = pymupdf.DocumentWriter(dst)
    story = pymupdf.Story(html="".join(parts))
    if landscape:
        try:
            mediabox = pymupdf.paper_rect("a4-l")
        except Exception:
            mediabox = pymupdf.Rect(0, 0, 842, 595)
    else:
        mediabox = pymupdf.paper_rect("a4")
    where = mediabox + (24, 24, -24, -24)
    more = 1
    while more:
        dev = writer.begin_page(mediabox)
        more, filled = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()
    return True, "未检测到 Office/WPS, 已用内置引擎简易排版(样式有损)"


# ---------------------------------------------------------------- 各类转换

def word_to_pdf(src, dst, log=print):
    ok, msg = _com_to_pdf(src, dst, _WORD_PROGIDS)
    if ok:
        return True, msg
    log(msg + "，正在使用内置引擎兜底…")
    return _docx_text_fallback_pdf(src, dst)


def excel_to_pdf(src, dst, log=print, landscape=None, fit_wide=False):
    """Excel -> PDF
    landscape: None=跟随文档 / True=横向 / False=纵向
    fit_wide:  所有列缩放至一页宽 (宽表格显示更完整)"""
    ok, msg = _com_to_pdf(src, dst, _EXCEL_PROGIDS, page_setup=(landscape, fit_wide))
    if ok:
        return True, msg
    log(msg + "，正在使用内置引擎兜底…")
    return _xlsx_text_fallback_pdf(src, dst, landscape=landscape)


def ppt_to_pdf(src, dst, log=print):
    ok, msg = _com_to_pdf(src, dst, _PPT_PROGIDS)
    return (True, msg) if ok else (False, "PPT 转 PDF 需要本机安装 MS Office 或 WPS（" + msg + "）")


def pdf_to_word(src, dst, log=print):
    from pdf2docx import Converter
    cv = Converter(src)
    try:
        cv.convert(dst)
    finally:
        cv.close()
    return True, "布局还原转换完成"


def pdf_to_images(src, dst_dir_fmt, log=print, dpi=180, fmt="png"):
    """PDF 每页导出一张图片; dst_dir_fmt 形如 'C:/out/文件名_图片'"""
    import pymupdf
    os.makedirs(dst_dir_fmt, exist_ok=True)
    doc = pymupdf.open(src)
    n = len(doc)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        out = os.path.join(dst_dir_fmt, f"page_{i+1:03d}.{fmt}")
        pix.save(out)
        log(f"  第 {i+1}/{n} 页 -> {os.path.basename(out)}")
    doc.close()
    return True, f"共导出 {n} 张图片到 {dst_dir_fmt}"


def images_to_pdf(sources, dst, log=print):
    import pymupdf
    doc = pymupdf.open()
    for img in sources:
        try:
            img_doc = pymupdf.open(img)
            pdfbytes = img_doc.convert_to_pdf()
            img_doc.close()
            src_pdf = pymupdf.open("pdf", pdfbytes)
            doc.insert_pdf(src_pdf)
            src_pdf.close()
        except Exception as e:
            log(f"  跳过无法识别的图片 {os.path.basename(img)}: {e}")
    if len(doc) == 0:
        doc.close()
        return False, "没有可用的图片"
    doc.save(dst)
    doc.close()
    return True, f"已将 {len(sources)} 张图片合成为 PDF"


def word_to_txt(src, dst, log=print):
    from docx import Document
    doc = Document(src)
    lines = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            lines.append("\t".join(c.text.strip() for c in row.cells))
    with open(dst, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines))
    return True, "导出纯文本完成"


def txt_to_word(src, dst, log=print):
    from docx import Document
    doc = Document()
    with open(src, "r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            doc.add_paragraph(line.rstrip("\n"))
    doc.save(dst)
    return True, "生成 Word 文档完成"


def excel_to_csv(src, out_dir, log=print):
    """每个工作表导出一个 CSV, 输出到 out_dir"""
    import openpyxl
    os.makedirs(out_dir, exist_ok=True)
    wb = openpyxl.load_workbook(src, data_only=True)
    base = os.path.splitext(os.path.basename(src))[0]
    count = 0
    for ws in wb.worksheets:
        safe = "".join(c for c in ws.title if c not in '\\/:*?"<>|') or f"sheet{count+1}"
        out = os.path.join(out_dir, f"{base}_{safe}.csv")
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            n_rows = min(ws.max_row or 1, 100000)
            n_cols = min(ws.max_column or 1, 256)
            for row in ws.iter_rows(max_row=n_rows, max_col=n_cols, values_only=True):
                w.writerow(["" if v is None else v for v in row])
        count += 1
        log(f"  工作表[{ws.title}] -> {os.path.basename(out)}")
    return True, f"共导出 {count} 个 CSV 到 {out_dir}"


def csv_to_excel(src, dst, log=print):
    import openpyxl
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            with open(src, "r", encoding=enc, newline="") as f:
                rows = list(csv.reader(f))
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        return False, "无法识别 CSV 编码"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in rows:
        ws.append(row)
    wb.save(dst)
    return True, f"共写入 {len(rows)} 行"


# ---------------------------------------------------------------- PDF 工具

def pdf_merge(sources, dst, log=print):
    import pymupdf
    doc = pymupdf.open()
    for s in sources:
        try:
            src = pymupdf.open(s)
            doc.insert_pdf(src)
            src.close()
            log(f"  已并入 {os.path.basename(s)}")
        except Exception as e:
            log(f"  跳过 {os.path.basename(s)}: {e}")
    if len(doc) == 0:
        doc.close()
        return False, "没有可合并的 PDF"
    doc.save(dst)
    doc.close()
    return True, f"合并完成, 共 {len(sources)} 个文件"


def pdf_split(src, out_dir, ranges=None, log=print):
    """拆分 PDF。ranges 为 None 时每页一个文件; 否则形如 ['1-3','5'] 每段一个文件"""
    import pymupdf
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(src)
    base = os.path.splitext(os.path.basename(src))[0]
    total = len(doc)
    outputs = []
    if not ranges:
        for i in range(total):
            part = pymupdf.open()
            part.insert_pdf(doc, from_page=i, to_page=i)
            out = os.path.join(out_dir, f"{base}_p{i+1:03d}.pdf")
            part.save(out)
            part.close()
            outputs.append(out)
        msg = f"已按页拆分为 {total} 个文件"
    else:
        idx = 0
        for r in ranges:
            r = r.strip().replace("页", "")
            try:
                if "-" in r:
                    a, b = r.split("-", 1)
                    a, b = max(1, int(a)), min(total, int(b))
                else:
                    a = b = int(r)
                if a > b:
                    a, b = b, a
            except ValueError:
                log(f"  忽略无法解析的页码范围: {r}")
                continue
            part = pymupdf.open()
            part.insert_pdf(doc, from_page=a - 1, to_page=b - 1)
            out = os.path.join(out_dir, f"{base}_{a}-{b}.pdf")
            part.save(out)
            part.close()
            outputs.append(out)
            idx += 1
        msg = f"已拆分为 {idx} 个文件"
    doc.close()
    return True, msg


def pdf_encrypt(src, dst, user_pw, owner_pw=None, log=print):
    import pymupdf
    doc = pymupdf.open(src)
    doc.save(dst, encryption=pymupdf.PDF_ENCRYPT_AES_256,
             owner_pw=owner_pw or user_pw, user_pw=user_pw)
    doc.close()
    return True, "已加密保存 (AES-256)"


def pdf_decrypt(src, dst, password, log=print):
    import pymupdf
    doc = pymupdf.open(src)
    if doc.needs_pass:
        if not doc.authenticate(password):
            doc.close()
            return False, "密码错误, 解密失败"
    doc.save(dst, encryption=pymupdf.PDF_ENCRYPT_NONE)
    n = len(doc)
    doc.close()
    return True, f"解密成功, 共 {n} 页"
