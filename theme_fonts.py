#!/usr/bin/env python3
"""
theme_fonts.py - 修改 Word 文件的佈景主題字型 (Headings / Body)，等同 Customize Fonts。完全離線。

三種輸出方式
  覆寫檔案    直接修改選取的 .docx（先寫暫存檔，成功後才取代，中途失敗不會弄壞原檔）
  建立新檔    把修改後的副本另存成新檔，原檔不變
  建立新文件  不需要原檔，依設定的字型建立一份空白的新文件

用法
  GUI :  python3 theme_fonts.py                       (有 PySide6 為液態玻璃介面，否則為基本 Tk 介面)
  CLI :  python3 theme_fonts.py 論文.docx --latin "Times New Roman" --ea "標楷體"              # 覆寫
         python3 theme_fonts.py 論文.docx --latin Arial --ea "PingFang TC" -o 新檔.docx         # 建立新檔
         python3 theme_fonts.py --new-blank 空白.docx --major-latin Arial --minor-latin "Times New Roman"  # 建立新文件
  字型欄位留空 = 不修改該項。

原理: .docx 是 zip，主題字型存在 word/theme/theme1.xml，只改這個檔案，內文與其他部分原封不動。
"""
import argparse
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import quoteattr

THEME = "word/theme/theme1.xml"
FONT_KEYS = ("major_latin", "major_ea", "minor_latin", "minor_ea")


# ---------------------------------------------------------------- 主題字型修改
def _set_face(body, tag_pattern, face):
    """把符合 tag_pattern 的第一個標籤的 typeface 換成 face，並移除會與新字型衝突的 panose 等屬性。"""
    def repl(m):
        tag = m.group(0)
        tag = re.sub(r'typeface="[^"]*"', "typeface=" + quoteattr(face), tag)
        tag = re.sub(r'\s+(panose|pitchFamily|charset)="[^"]*"', "", tag)
        return tag
    return re.sub(tag_pattern, repl, body, count=1)


def _patch_block(xml, block, latin, ea):
    pat = re.compile(rf"(<a:{block}>)(.*?)(</a:{block}>)", re.S)
    m = pat.search(xml)
    if not m:
        raise ValueError(f"找不到 <a:{block}>，這份文件的主題格式不尋常。")
    body = m.group(2)
    if latin:
        body = _set_face(body, r"<a:latin\b[^>]*/>", latin)
    if ea:
        body = _set_face(body, r"<a:ea\b[^>]*/>", ea)
        # 繁體中文實際走的是 script="Hant" 這一項
        body = _set_face(body, r'<a:font\b[^>]*script="Hant"[^>]*/>', ea)
    return xml[: m.start()] + m.group(1) + body + m.group(3) + xml[m.end():]


def _patch_theme_xml(xml, major_latin, major_ea, minor_latin, minor_ea):
    xml = _patch_block(xml, "majorFont", major_latin, major_ea)  # Headings
    xml = _patch_block(xml, "minorFont", minor_latin, minor_ea)  # Body
    return xml


def _is_open_in_word(src):
    return any(f.name.endswith(src.name[2:]) for f in src.parent.glob("~$*"))


def set_theme_fonts(src, dst=None, major_latin=None, major_ea=None,
                    minor_latin=None, minor_ea=None):
    """修改 src 的主題字型。dst 省略（或與 src 相同）時直接覆寫 src，否則另存成 dst。"""
    src = Path(src)
    if src.suffix.lower() not in (".docx", ".dotx"):
        raise ValueError("只支援 .docx / .dotx 檔案。")
    if not src.exists():
        raise ValueError(f"找不到檔案：{src}")
    if not any((major_latin, major_ea, minor_latin, minor_ea)):
        raise ValueError("請至少設定一個字型（留空的欄位表示不修改）。")

    target = Path(dst) if dst else src
    overwriting = target.resolve() == src.resolve()
    if overwriting and _is_open_in_word(src):
        raise ValueError("這份文件似乎還在 Word 中開啟，請先關閉再執行。")
    tmp = target.with_name(target.name + ".tmp")

    try:
        with zipfile.ZipFile(src) as zin:
            if THEME not in zin.namelist():
                raise ValueError("此文件沒有主題檔 (theme1.xml)。")
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == THEME:
                        xml = _patch_theme_xml(data.decode("utf-8"), major_latin, major_ea,
                                               minor_latin, minor_ea)
                        data = xml.encode("utf-8")
                    zout.writestr(item, data)
        os.replace(tmp, target)  # 原子取代
    finally:
        if tmp.exists():
            tmp.unlink()
    return target


# ---------------------------------------------------------------- 建立空白新文件
_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HEAD_FONTS = ('<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" '
               'w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>')

_BLANK_PARTS = {
    "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>""",
    "_rels/.rels": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_NS_R}"><Relationship Id="rId1" Type="{_REL}/officeDocument" Target="word/document.xml"/></Relationships>""",
    "word/_rels/document.xml.rels": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_NS_R}">
<Relationship Id="rId1" Type="{_REL}/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="{_REL}/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="{_REL}/theme" Target="theme/theme1.xml"/>
</Relationships>""",
    "word/document.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_NS_W}"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""",
    "word/settings.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{_NS_W}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>""",
    # 內文走主題的 Body (minor) 字型，標題 / Heading 1-3 走 Headings (major) 字型
    "word/styles.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{_NS_W}">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr>{_HEAD_FONTS}<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr>{_HEAD_FONTS}<w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr>{_HEAD_FONTS}<w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr>{_HEAD_FONTS}<w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>""",
}


def _font_scheme(heading_latin, body_latin):
    def block(tag, latin):
        return (f'<a:{tag}><a:latin typeface="{latin}"/><a:ea typeface=""/><a:cs typeface=""/>'
                '<a:font script="Jpan" typeface="MS Mincho"/><a:font script="Hang" typeface="Malgun Gothic"/>'
                '<a:font script="Hans" typeface="SimSun"/><a:font script="Hant" typeface="PMingLiU"/>'
                f'</a:{tag}>')
    return ('<a:fontScheme name="Office">' + block("majorFont", heading_latin) +
            block("minorFont", body_latin) + '</a:fontScheme>')


def _base_theme_xml():
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    ln = lambda w: (f'<a:ln w="{w}" cap="flat" cmpd="sng" algn="ctr">{fill}'
                    '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>')
    clr = ('<a:clrScheme name="Office">'
           '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
           '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
           '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
           '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
           '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
           '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
           '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
           '</a:clrScheme>')
    fmt = ('<a:fmtScheme name="Office">'
           f'<a:fillStyleLst>{fill}{fill}{fill}</a:fillStyleLst>'
           f'<a:lnStyleLst>{ln(6350)}{ln(12700)}{ln(19050)}</a:lnStyleLst>'
           + '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + '</a:effectStyleLst>'
           f'<a:bgFillStyleLst>{fill}{fill}{fill}</a:bgFillStyleLst>'
           '</a:fmtScheme>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
            f'<a:themeElements>{clr}{_font_scheme("Calibri Light", "Calibri")}{fmt}</a:themeElements>'
            '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')


def create_blank_document(dst, major_latin=None, major_ea=None, minor_latin=None, minor_ea=None):
    """不需要原檔，依指定的主題字型建立一份空白的新 .docx（A4）。欄位留空則用 Office 預設。"""
    dst = Path(dst)
    if dst.suffix.lower() != ".docx":
        dst = dst.with_name(dst.name + ".docx")
    theme = _patch_theme_xml(_base_theme_xml(), major_latin, major_ea, minor_latin, minor_ea)
    tmp = dst.with_name(dst.name + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, content in _BLANK_PARTS.items():       # [Content_Types].xml 放第一個
                z.writestr(name, content.encode("utf-8"))
            z.writestr(THEME, theme.encode("utf-8"))
        os.replace(tmp, dst)
    finally:
        if tmp.exists():
            tmp.unlink()
    return dst


# ---------------------------------------------------------------- 共用：GUI 動作
def _default_copy_path(src):
    p = Path(src)
    return str(p.with_name(p.stem + "_fonts.docx"))


def _default_blank_path(src=None):
    base = Path(src).parent if src else Path.home()
    return str(base / "新文件.docx")


# ---------------------------------------------------------------- GUI (基本 Tk 版)
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter import font as tkfont

    root = tk.Tk()
    root.title("Word 主題字型修改器")
    root.resizable(False, False)

    families = sorted({f for f in tkfont.families() if not f.startswith("@")})
    pad = {"padx": 8, "pady": 5}

    file_var = tk.StringVar()
    vars_ = {k: tk.StringVar() for k in FONT_KEYS}

    def fonts():
        return {k: v.get().strip() or None for k, v in vars_.items()}

    def choose():
        p = filedialog.askopenfilename(filetypes=[("Word 文件", "*.docx *.dotx")])
        if p:
            file_var.set(p)

    def need_file():
        if not file_var.get().strip():
            messagebox.showwarning("提醒", "請先選擇 .docx 檔案。")
            return False
        return True

    def run(fn, *args, label):
        try:
            out = fn(*args, **fonts())
        except Exception as e:
            messagebox.showerror("失敗", str(e))
            return
        messagebox.showinfo("完成", f"{label}：\n{out}")

    def do_overwrite():
        if need_file():
            run(set_theme_fonts, file_var.get().strip(), label="已覆寫")

    def do_copy():
        if not need_file():
            return
        dst = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word 文件", "*.docx")],
                                           initialfile=Path(_default_copy_path(file_var.get().strip())).name,
                                           initialdir=str(Path(file_var.get().strip()).parent))
        if dst:
            run(set_theme_fonts, file_var.get().strip(), dst, label="已建立新檔")

    def do_blank():
        src = file_var.get().strip() or None
        dst = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word 文件", "*.docx")],
                                           initialfile="新文件.docx",
                                           initialdir=str(Path(_default_blank_path(src)).parent))
        if dst:
            run(create_blank_document, dst, label="已建立新文件")

    ttk.Label(root, text="文件").grid(row=0, column=0, sticky="e", **pad)
    ttk.Entry(root, textvariable=file_var, width=44).grid(row=0, column=1, columnspan=2, **pad)
    ttk.Button(root, text="選擇…", command=choose).grid(row=0, column=3, **pad)

    ttk.Label(root, text="英文字型 (Latin)").grid(row=1, column=1, **pad)
    ttk.Label(root, text="中文字型 (East Asian)").grid(row=1, column=2, **pad)

    rows = [("標題 Headings", "major_latin", "major_ea"), ("內文 Body", "minor_latin", "minor_ea")]
    for i, (label, lk, ek) in enumerate(rows, start=2):
        ttk.Label(root, text=label).grid(row=i, column=0, sticky="e", **pad)
        for col, key in ((1, lk), (2, ek)):
            cb = ttk.Combobox(root, textvariable=vars_[key], values=families, width=26)
            cb.grid(row=i, column=col, **pad)

            def filt(e, cb=cb):
                if e.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
                    return
                t = cb.get().lower()
                cb["values"] = [f for f in families if t in f.lower()] if t else families
            cb.bind("<KeyRelease>", filt)

    ttk.Label(root, text="字型欄位留空 = 不修改該項；可直接輸入文字搜尋字型。",
              foreground="gray").grid(row=4, column=1, columnspan=2, pady=(2, 0))

    bar = ttk.Frame(root)
    bar.grid(row=5, column=0, columnspan=4, pady=12)
    ttk.Button(bar, text="覆寫檔案", command=do_overwrite).pack(side="left", padx=6)
    ttk.Button(bar, text="建立新檔…", command=do_copy).pack(side="left", padx=6)
    ttk.Button(bar, text="建立新文件…", command=do_blank).pack(side="left", padx=6)

    ttk.Label(root, justify="left", foreground="gray", text=(
        "覆寫檔案：直接修改原檔（請先關閉 Word 內的該文件）\n"
        "建立新檔：把修改後的副本另存新檔，原檔不變\n"
        "建立新文件：不需原檔，依設定的字型建立空白文件")).grid(row=6, column=0, columnspan=4, pady=(0, 10))
    root.mainloop()


# ------------------------------------------------------- GUI (Liquid Glass, PySide6)
GLASS_QSS = """
QLabel { color: #1d1d1f; background: transparent; }
QLabel#title { font-size: 20px; font-weight: 600; }
QLabel#hint { color: rgba(0,0,0,0.50); font-size: 11px; }
QLineEdit, QComboBox {
    background: rgba(255,255,255,0.55);
    border: 1px solid rgba(255,255,255,0.85);
    border-radius: 12px; padding: 6px 10px; color: #1d1d1f; min-height: 22px;
    selection-background-color: rgba(10,132,255,0.35);
}
QLineEdit:focus, QComboBox:focus {
    border: 1.5px solid rgba(10,132,255,0.80); background: rgba(255,255,255,0.78);
}
QComboBox QAbstractItemView {
    background: rgba(255,255,255,0.97); border: 1px solid rgba(0,0,0,0.08);
    border-radius: 10px; outline: 0; padding: 4px; color: #1d1d1f;
    selection-background-color: rgba(10,132,255,0.25); selection-color: #1d1d1f;
}
QPushButton#glassBtn {
    background: rgba(255,255,255,0.50); border: 1px solid rgba(255,255,255,0.90);
    border-radius: 12px; padding: 6px 14px; color: #1d1d1f;
}
QPushButton#glassBtn:hover { background: rgba(255,255,255,0.78); }
QPushButton#act {
    background: rgba(255,255,255,0.58); border: 1px solid rgba(255,255,255,0.92);
    border-radius: 16px; padding: 10px 8px; color: #1d1d1f; font-weight: 600; min-height: 22px;
}
QPushButton#act:hover { background: rgba(255,255,255,0.85); }
QPushButton#act:pressed { background: rgba(255,255,255,1); }
QPushButton#primary {
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 rgba(72,160,255,0.97), stop:1 rgba(10,110,235,0.97));
    border: 1px solid rgba(255,255,255,0.65); border-radius: 16px;
    color: white; font-weight: 600; padding: 10px 8px; min-height: 22px;
}
QPushButton#primary:hover {
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 rgba(100,176,255,1), stop:1 rgba(30,125,245,1));
}
QPushButton#primary:pressed { background: rgba(10,100,215,1); }
"""


def _enable_macos_blur(win, radius):
    """實驗性：在視窗後方加入 macOS 原生毛玻璃 (NSVisualEffectView)。需要 pyobjc-framework-Cocoa。"""
    try:
        import ctypes
        import objc
        from AppKit import NSVisualEffectView
        view = objc.objc_object(c_void_p=ctypes.c_void_p(int(win.winId())))
        m = win.M
        w, h = win.width(), win.height()
        fx = NSVisualEffectView.alloc().initWithFrame_(((m, m), (w - 2 * m, h - 2 * m)))
        fx.setBlendingMode_(0)   # behind window
        fx.setMaterial_(13)      # HUD window
        fx.setState_(1)          # active
        fx.setAutoresizingMask_(18)
        fx.setWantsLayer_(True)
        fx.layer().setCornerRadius_(radius)
        fx.layer().setMasksToBounds_(True)
        view.superview().addSubview_positioned_relativeTo_(fx, -1, view)
        return True
    except Exception as e:  # 失敗就退回純半透明玻璃
        print(f"(毛玻璃效果未啟用：{e})")
        return False


def _build_qt(blur=False):
    from PySide6.QtCore import Qt, QRectF, QPointF
    from PySide6.QtGui import (QColor, QFontDatabase, QLinearGradient, QPainter,
                               QPainterPath, QPen, QBrush)
    from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                                   QGridLayout, QLabel, QLineEdit, QPushButton,
                                   QComboBox, QCompleter, QFileDialog)

    app = QApplication.instance() or QApplication(sys.argv)
    families = sorted({f for f in QFontDatabase.families() if not f.startswith(".")}, key=str.lower)

    class Glass(QWidget):
        M, R = 28, 28

        def __init__(self):
            super().__init__()
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setAcceptDrops(True)
            self.setWindowTitle("Word 主題字型修改器")
            self.base = 85 if blur else 170   # 無原生模糊時加強不透明度，確保文字可讀
            self._drag = None
            self.setStyleSheet(GLASS_QSS)
            self._ui()
            self.resize(660, 540)
            self.setFixedSize(self.size())

        # ---------- UI
        def _ui(self):
            m = self.M
            root = QVBoxLayout(self)
            root.setContentsMargins(m + 26, m + 16, m + 26, m + 22)
            root.setSpacing(12)

            bar = QHBoxLayout(); bar.setSpacing(8)
            for color, slot in (("#ff5f57", self.close), ("#febc2e", self.showMinimized)):
                b = QPushButton(); b.setFixedSize(13, 13); b.clicked.connect(slot)
                b.setStyleSheet(f"background:{color}; border:none; border-radius:6px;")
                bar.addWidget(b)
            bar.addStretch(); root.addLayout(bar)

            t = QLabel("主題字型修改器"); t.setObjectName("title"); root.addWidget(t)
            h = QLabel("只修改 Headings / Body 主題字型，內文不變。字型欄位留空 = 不修改該項。")
            h.setObjectName("hint"); h.setWordWrap(True); root.addWidget(h)

            row = QHBoxLayout()
            self.file = QLineEdit(); self.file.setPlaceholderText("選擇或拖曳 .docx 檔案到這裡")
            pick = QPushButton("選擇…"); pick.setObjectName("glassBtn"); pick.clicked.connect(self.choose)
            row.addWidget(self.file, 1); row.addWidget(pick); root.addLayout(row)

            grid = QGridLayout(); grid.setHorizontalSpacing(12); grid.setVerticalSpacing(10)
            grid.addWidget(QLabel("英文字型 (Latin)"), 0, 1)
            grid.addWidget(QLabel("中文字型 (East Asian)"), 0, 2)
            self.cb = {}
            for r, (label, lk, ek) in enumerate((("標題 Headings", "major_latin", "major_ea"),
                                                 ("內文 Body", "minor_latin", "minor_ea")), start=1):
                grid.addWidget(QLabel(label), r, 0)
                for c, key in ((1, lk), (2, ek)):
                    grid.addWidget(self._combo(key), r, c)
            grid.setColumnStretch(1, 1); grid.setColumnStretch(2, 1)
            root.addLayout(grid)

            tip = QLabel("在欄位內直接輸入即可搜尋本機字型（例如 ming、hei、times）。")
            tip.setObjectName("hint"); root.addWidget(tip)

            root.addStretch()
            self.status = QLabel(""); self.status.setWordWrap(True); self.status.setMinimumHeight(20)
            root.addWidget(self.status)

            btns = QHBoxLayout(); btns.setSpacing(10)
            for text, name, slot in (("覆寫檔案", "primary", self.do_overwrite),
                                     ("建立新檔…", "act", self.do_copy),
                                     ("建立新文件…", "act", self.do_blank)):
                b = QPushButton(text); b.setObjectName(name); b.clicked.connect(lambda _=False, s=slot: s())
                b.setCursor(Qt.PointingHandCursor)
                btns.addWidget(b, 1)
            root.addLayout(btns)

            note = QLabel("覆寫檔案：直接修改原檔（請先關閉 Word 中的該文件）\n"
                          "建立新檔：把修改後的副本另存新檔，原檔不變\n"
                          "建立新文件：不需原檔，依設定的字型建立空白文件")
            note.setObjectName("hint"); root.addWidget(note)

        def _combo(self, key):
            cb = QComboBox(); cb.setEditable(True); cb.setInsertPolicy(QComboBox.NoInsert)
            cb.addItems(families)
            cb.setCurrentIndex(-1)                       # 預設留空 = 不修改
            cb.lineEdit().setPlaceholderText("留空＝不修改")
            comp = cb.completer()
            comp.setFilterMode(Qt.MatchContains)
            comp.setCaseSensitivity(Qt.CaseInsensitive)
            comp.setCompletionMode(QCompleter.PopupCompletion)
            comp.popup().setStyleSheet(GLASS_QSS)
            self.cb[key] = cb
            return cb

        # ---------- 行為
        def choose(self):
            p, _ = QFileDialog.getOpenFileName(self, "選擇文件", "", "Word 文件 (*.docx *.dotx)")
            if p:
                self.file.setText(p)

        def say(self, msg, err=False):
            self.status.setStyleSheet(f"color:{'#c62828' if err else '#1b7f3b'}; background:transparent;")
            self.status.setText(msg)

        def fonts(self):
            return {k: c.currentText().strip() or None for k, c in self.cb.items()}

        def _src(self):
            p = self.file.text().strip()
            if not p:
                self.say("請先選擇 .docx 檔案。", True)
            return p

        def _ask_save(self, title, default):
            p, _ = QFileDialog.getSaveFileName(self, title, default, "Word 文件 (*.docx)")
            if p and not p.lower().endswith(".docx"):
                p += ".docx"
            return p

        def _run(self, fn, *args, done):
            try:
                out = fn(*args, **self.fonts())
            except Exception as e:
                return self.say(str(e), True)
            self.say(f"{done}：{Path(out).name}")

        def do_overwrite(self):
            src = self._src()
            if src:
                self._run(set_theme_fonts, src, done="完成，已覆寫")

        def do_copy(self, dst=None):
            src = self._src()
            if not src:
                return
            dst = dst or self._ask_save("建立新檔", _default_copy_path(src))
            if dst:
                self._run(set_theme_fonts, src, dst, done="完成，已建立新檔")

        def do_blank(self, dst=None):
            dst = dst or self._ask_save("建立新文件", _default_blank_path(self.file.text().strip() or None))
            if dst:
                self._run(create_blank_document, dst, done="完成，已建立新文件")

        # ---------- 拖放與移動
        def dragEnterEvent(self, e):
            if e.mimeData().hasUrls():
                e.acceptProposedAction()

        def dropEvent(self, e):
            u = e.mimeData().urls()
            if u:
                self.file.setText(u[0].toLocalFile())

        def mousePressEvent(self, e):
            if e.button() == Qt.LeftButton:
                self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

        def mouseMoveEvent(self, e):
            if self._drag is not None and e.buttons() & Qt.LeftButton:
                self.move(e.globalPosition().toPoint() - self._drag)

        def mouseReleaseEvent(self, e):
            self._drag = None

        # ---------- 玻璃繪製
        def paintEvent(self, _):
            p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
            r = QRectF(self.rect()).adjusted(self.M, self.M, -self.M, -self.M)
            R, a = self.R, self.base
            p.setPen(Qt.NoPen)
            for i in range(14):                       # 柔和陰影
                p.setBrush(QColor(0, 0, 0, 5))
                p.drawRoundedRect(r.adjusted(-i, -i + 7, i, i + 7), R + i, R + i)
            body = QLinearGradient(r.topLeft(), r.bottomRight())
            body.setColorAt(0.0, QColor(255, 255, 255, min(255, a + 45)))
            body.setColorAt(0.5, QColor(255, 255, 255, a))
            body.setColorAt(1.0, QColor(228, 238, 255, min(255, a + 25)))
            p.setBrush(QBrush(body)); p.drawRoundedRect(r, R, R)
            hl = QLinearGradient(r.topLeft(), QPointF(r.left(), r.top() + r.height() * 0.42))
            hl.setColorAt(0, QColor(255, 255, 255, 120)); hl.setColorAt(1, QColor(255, 255, 255, 0))
            path = QPainterPath(); path.addRoundedRect(r, R, R)
            p.fillPath(path, QBrush(hl))              # 頂部高光
            rim = QLinearGradient(r.topLeft(), r.bottomRight())
            rim.setColorAt(0, QColor(255, 255, 255, 235))
            rim.setColorAt(0.5, QColor(255, 255, 255, 70))
            rim.setColorAt(1, QColor(255, 255, 255, 170))
            p.setPen(QPen(QBrush(rim), 1.6)); p.setBrush(Qt.NoBrush)
            p.drawRoundedRect(r.adjusted(.8, .8, -.8, -.8), R, R)

    win = Glass()
    return app, win


def run_gui_qt(blur=False):
    app, win = _build_qt(blur)
    win.show()
    if blur:
        _enable_macos_blur(win, win.R)
    app.exec()


# ---------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description="修改 Word 主題字型 (Headings / Body)；字型參數留空 = 不修改該項")
    ap.add_argument("file", nargs="?", help=".docx 檔案；省略則開啟 GUI")
    ap.add_argument("-o", "--output", help="建立新檔：把修改後的副本另存成這個檔案（省略則覆寫原檔）")
    ap.add_argument("--new-blank", metavar="OUT.docx", help="建立新文件：依指定字型建立空白文件（不需要原檔）")
    ap.add_argument("--latin", help="標題與內文的英文字型")
    ap.add_argument("--ea", help="標題與內文的中文字型")
    ap.add_argument("--major-latin"); ap.add_argument("--major-ea")
    ap.add_argument("--minor-latin"); ap.add_argument("--minor-ea")
    ap.add_argument("--tk", action="store_true", help="使用基本 Tk 介面")
    ap.add_argument("--blur", action="store_true", help="(實驗性) 啟用 macOS 原生毛玻璃")
    a = ap.parse_args()

    fonts = dict(major_latin=a.major_latin or a.latin, major_ea=a.major_ea or a.ea,
                 minor_latin=a.minor_latin or a.latin, minor_ea=a.minor_ea or a.ea)

    if a.new_blank:
        try:
            out = create_blank_document(a.new_blank, **fonts)
        except Exception as e:
            sys.exit(f"失敗：{e}")
        print(f"已建立新文件：{out}")
        return

    if not a.file:
        if a.tk:
            run_gui()
            return
        try:
            import PySide6  # noqa: F401
        except ImportError:
            print("未安裝 PySide6，改用基本介面。想要液態玻璃介面請執行：pip3 install PySide6")
            run_gui()
        else:
            run_gui_qt(blur=a.blur)
        return

    try:
        out = set_theme_fonts(a.file, a.output, **fonts)
    except Exception as e:
        sys.exit(f"失敗：{e}")
    print(f"{'已建立新檔' if a.output else '已覆寫'}：{out}")


if __name__ == "__main__":
    main()
