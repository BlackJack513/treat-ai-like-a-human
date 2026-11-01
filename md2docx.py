#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx.py —— 把一整篇文章（若干 markdown 章节）合并导出成一份 Word 文档。

  * 有 index.md：按索引里链接的先后顺序拼接章节（索引本身只用来取标题、定顺序）
  * 没有 index.md：递归扫描目录下所有 .md，按文件名自然排序
  * 支持：标题、段落、粗体/斜体/删除线/行内代码、有序/无序列表（可嵌套）、
          表格（含表头）、围栏代码块、引用、分隔线、图片
  * 中文排版：A4、正文宋体 12pt / 1.5 倍行距、标题微软雅黑、页脚页码

用法
  python md2docx.py                          # 在当前目录找 index.md，导出 标题.docx
  python md2docx.py -o 我的文章.docx
  python md2docx.py -r 某本书 --title "书名"
  python md2docx.py --toc --no-page-break
  python md2docx.py 1.md 2.md 3.md -o 合集.docx

依赖
  pip install python-docx
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.stderr.write("缺少依赖 python-docx，请先执行：\n    pip install python-docx\n")
    raise SystemExit(1)


# ---------------------------------------------------------------- 排版参数 --

BODY_LATIN = "Times New Roman"
BODY_CJK = "宋体"
HEAD_LATIN = "微软雅黑"
HEAD_CJK = "微软雅黑"
CODE_LATIN = "Consolas"
CODE_CJK = "宋体"

BODY_SIZE = 12.0                       # 小四
HEAD_SIZES = {1: 20.0, 2: 16.0, 3: 14.0, 4: 12.5}
HEAD_SPACE = {1: (18, 10), 2: (14, 8), 3: (12, 6), 4: (10, 4)}

TEXT_WIDTH_CM = 14.0                   # 版心宽度，用于图片缩放
CODE_FILL = "F4F5F7"
RULE_COLOR = "D0D7DE"
QUOTE_COLOR = "9AA4B2"

THEME_ATTRS = ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme")
PPR_TAIL = ("w:rPr", "w:sectPr")
PPR_AFTER_BDR = ("w:shd", "w:tabs", "w:spacing", "w:ind", "w:jc") + PPR_TAIL
PPR_AFTER_SHD = ("w:tabs", "w:spacing", "w:ind", "w:jc") + PPR_TAIL


# ------------------------------------------------------------ OOXML 小工具 --

def _get_rfonts(rpr):
    """取出（必要时创建）w:rPr 下的 w:rFonts。"""
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    return rfonts


def _clear_theme_fonts(rfonts):
    """去掉 xxxTheme 属性，否则 Word 会优先用主题字体，覆盖显式字体。"""
    for attr in THEME_ATTRS:
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def set_run_east_asia(run, font_name):
    """单独指定 run 的中文字体：python-docx 的 font.name 只管西文。"""
    rpr = run._element.get_or_add_rPr()
    _get_rfonts(rpr).set(qn("w:eastAsia"), font_name)


def _style_rpr(style):
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        style.element.append(rpr)          # w:style 里 rPr 排在 pPr 之后，追加合法
    return rpr


def configure_style(style, latin, cjk, size=None, bold=None, color=None):
    """一次性设定样式的西文/中文字体、字号、加粗、颜色，并清掉主题字体。"""
    style.font.name = latin
    rpr = _style_rpr(style)
    rfonts = _get_rfonts(rpr)
    _clear_theme_fonts(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:cs"), latin)
    rfonts.set(qn("w:eastAsia"), cjk)

    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = RGBColor.from_string(color)
        color_el = rpr.find(qn("w:color"))
        if color_el is not None:
            for attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
                if color_el.get(qn(attr)) is not None:
                    del color_el.attrib[qn(attr)]


def ppr_add(p_element, child, followers=PPR_AFTER_SHD):
    """把 child 插到 w:pPr 里合法且靠前的位置（Word 对子元素顺序比较敏感）。"""
    ppr = p_element.get_or_add_pPr()
    for tag in followers:
        node = ppr.find(qn(tag))
        if node is not None:
            node.addprevious(child)
            return
    ppr.append(child)


def make_border(edge, size=6, color=RULE_COLOR, space=6):
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(size))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    return el


def make_shd(fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    return shd


# ------------------------------------------------------------ 行内 markdown --

INLINE_RE = re.compile(
    r"(?P<code>`[^`\n]+`)"
    r"|(?P<img>!\[[^\]\n]*\]\([^)\n]*\))"
    r"|(?P<link>\[[^\]\n]+\]\([^)\n]*\))"
    r"|(?P<bold>\*\*(?=\S).+?(?<=\S)\*\*)"
    r"|(?P<bold2>__(?=\S).+?(?<=\S)__)"
    r"|(?P<strike>~~(?=\S).+?(?<=\S)~~)"
    r"|(?P<italic>\*(?=\S)[^*\n]+?(?<=\S)\*)"
    r"|(?P<italic2>_(?=\S)[^_\n]+?(?<=\S)_)"
)


def _add_run(paragraph, text, bold=False, size=None, color=None):
    run = paragraph.add_run(text)
    if bold:
        run.font.bold = True
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def add_inline(paragraph, text, bold=False, size=None, color=None):
    """把一段 markdown 行内文本写进段落，处理粗体/斜体/代码/链接/图片占位。"""
    if not text:
        return
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], bold, size, color)
        kind = next(k for k, v in m.groupdict().items() if v)
        token = m.group(0)

        if kind == "code":
            run = _add_run(paragraph, token[1:-1], bold, size, color)
            run.font.name = CODE_LATIN
            set_run_east_asia(run, CODE_CJK)
            run.font.size = Pt((size or BODY_SIZE) - 1.5)
            run._element.get_or_add_rPr().append(make_shd("F0F1F3"))
        elif kind in ("bold", "bold2"):
            add_inline(paragraph, token[2:-2], True, size, color)
        elif kind == "strike":
            run = _add_run(paragraph, token[2:-2], bold, size, color)
            run.font.strike = True
        elif kind in ("italic", "italic2"):
            run = _add_run(paragraph, token[1:-1], bold, size, color)
            run.font.italic = True
        elif kind == "link":
            label = re.match(r"\[([^\]]*)\]", token).group(1)
            add_inline(paragraph, label or token, bold, size, color)
        elif kind == "img":
            alt = re.match(r"!\[([^\]]*)\]", token).group(1)
            _add_run(paragraph, alt or "[图片]", bold, size, color)

        pos = m.end()

    if pos < len(text):
        _add_run(paragraph, text[pos:], bold, size, color)


# --------------------------------------------------------------- 文档构建 --

class Builder:
    """把解析出来的块逐个写进 docx。"""

    def __init__(self, doc):
        self.doc = doc
        self.pending_break = False
        self.counts = {"heading": 0, "paragraph": 0, "table": 0,
                       "image": 0, "code": 0, "list": 0}

    # -- 内部 --------------------------------------------------------------
    def _break_pending(self, paragraph):
        if self.pending_break:
            paragraph.paragraph_format.page_break_before = True
            self.pending_break = False

    def _new_paragraph(self, style_name=None):
        p = (self.doc.add_paragraph(style=self._style(style_name))
             if style_name else self.doc.add_paragraph())
        self._break_pending(p)
        return p

    def _style(self, name, fallback="Normal"):
        try:
            return self.doc.styles[name]
        except KeyError:
            return self.doc.styles[fallback]

    # -- 块 -----------------------------------------------------------------
    def heading(self, level, text):
        level = max(1, min(level, 4))
        p = self.doc.add_paragraph(style=self._style(f"Heading {level}"))
        self._break_pending(p)
        add_inline(p, text, bold=True)
        self.counts["heading"] += 1

    def paragraph(self, text):
        p = self._new_paragraph()
        add_inline(p, text)
        self.counts["paragraph"] += 1

    def list_item(self, level, text, ordered):
        base = "List Number" if ordered else "List Bullet"
        name = base if level <= 0 else f"{base} {min(level + 1, 3)}"
        p = self._new_paragraph(name)
        p.paragraph_format.space_after = Pt(2)
        if level > 0:                           # 万一样式里没有深层列表，也能看出层级
            p.paragraph_format.left_indent = Cm(0.74 + 0.74 * (level - 1))
        add_inline(p, text)
        self.counts["list"] += 1

    def quote(self, text):
        p = self._new_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.8)
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        pbdr = OxmlElement("w:pBdr")
        pbdr.append(make_border("left", size=12, color=QUOTE_COLOR, space=8))
        ppr_add(p._p, pbdr, PPR_AFTER_BDR)
        add_inline(p, text, color="444444")

    def code(self, text):
        p = self._new_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.3)
        pf.right_indent = Cm(0.3)
        pf.space_before = Pt(6)
        pf.space_after = Pt(8)
        pf.line_spacing = 1.0

        pbdr = OxmlElement("w:pBdr")
        pbdr.append(make_border("left", size=18, color="BFC7D1", space=6))
        ppr_add(p._p, pbdr, PPR_AFTER_BDR)
        ppr_add(p._p, make_shd(CODE_FILL), PPR_AFTER_SHD)

        # 段落内换行由 python-docx 自动转成 <w:br/>
        run = p.add_run(text.rstrip("\n"))
        run.font.name = CODE_LATIN
        set_run_east_asia(run, CODE_CJK)
        run.font.size = Pt(9.5)
        self.counts["code"] += 1

    def rule(self):
        p = self._new_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(6)
        pf.space_after = Pt(10)
        pf.line_spacing = 1.0
        pbdr = OxmlElement("w:pBdr")
        pbdr.append(make_border("bottom", size=6, color=RULE_COLOR, space=1))
        ppr_add(p._p, pbdr, PPR_AFTER_BDR)

    def table(self, header, rows, aligns):
        cols = max(len(header), max((len(r) for r in rows), default=0))
        if cols == 0:
            return
        if self.pending_break:                  # 表格本身不能分页，先垫一个空段落
            self._new_paragraph()

        table = self.doc.add_table(rows=1, cols=cols)
        try:
            table.style = self.doc.styles["Table Grid"]
        except KeyError:
            pass
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True

        def fill(cell, text, bold, align):
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_before = Pt(2)
            pf.space_after = Pt(2)
            pf.line_spacing = 1.15
            if align == "center":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align == "right":
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            add_inline(p, text, bold=bold, size=BODY_SIZE - 1.5)

        for i in range(cols):
            cell = table.rows[0].cells[i]
            fill(cell, header[i] if i < len(header) else "", True,
                 aligns[i] if i < len(aligns) else None)
            cell._tc.get_or_add_tcPr().append(make_shd("F0F1F3"))

        tr = table.rows[0]._tr                  # 表头跨页重复
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr.insert(0, trPr)
        repeat = OxmlElement("w:tblHeader")
        repeat.set(qn("w:val"), "true")
        trPr.append(repeat)

        for row in rows:
            cells = table.add_row().cells
            for i in range(cols):
                fill(cells[i], row[i] if i < len(row) else "", False,
                     aligns[i] if i < len(aligns) else None)
        self.counts["table"] += 1

    def image(self, alt, src, base_dir):
        target = Path(src)
        if not target.is_absolute():
            target = base_dir / src
        if target.is_file() and target.suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".bmp"}:
            p = self._new_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p.add_run().add_picture(str(target), width=Cm(TEXT_WIDTH_CM))
                self.counts["image"] += 1
                return
            except Exception:                   # 图片损坏则退回文字说明
                p._p.getparent().remove(p._p)
        self.paragraph(f"[图片：{alt or src}]")


# ------------------------------------------------------------ markdown 解析 --

def split_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.replace("\\|", "|").strip() for c in re.split(r"(?<!\\)\|", s)]


def is_table_sep(line):
    s = line.strip()
    if "-" not in s:
        return False
    cells = split_row(s)
    if not cells:
        return False
    return all(re.fullmatch(r":?-{2,}:?", c.replace(" ", "")) for c in cells)


def sep_aligns(line):
    out = []
    for c in split_row(line):
        c = c.replace(" ", "")
        if c.startswith(":") and c.endswith(":"):
            out.append("center")
        elif c.endswith(":"):
            out.append("right")
        else:
            out.append(None)
    return out


def is_cjk(ch):
    code = ord(ch)
    return (0x3000 <= code <= 0x303F or 0x3400 <= code <= 0x4DBF
            or 0x4E00 <= code <= 0x9FFF or 0xFF00 <= code <= 0xFFEF)


def join_paragraph(parts, hard_breaks):
    """段落内的软换行：中文之间直接相连，西文之间补空格。"""
    out = ""
    for idx, part in enumerate(parts):
        if idx == 0:
            out = part
        elif hard_breaks[idx - 1]:
            out += "\n" + part
        elif is_cjk(out[-1]) or is_cjk(part[0]):
            out += part
        else:
            out += " " + part
    return out


def is_block_start(lines, i):
    line = lines[i]
    s = line.strip()
    if not s:
        return False
    if re.match(r"^#{1,6}\s", s):
        return True
    if re.match(r"^(`{3,}|~{3,})", s):
        return True
    if s.startswith(">"):
        return True
    if re.match(r"^([-*+]|\d{1,3}[.)])\s+", s):
        return True
    if re.fullmatch(r"!\[[^\]]*\]\([^)]*\)", s):
        return True
    if re.fullmatch(r"([-*_])(\s*\1){2,}", s):
        return True
    if "|" in line and i + 1 < len(lines) and is_table_sep(lines[i + 1]):
        return True
    return False


def strip_front_matter(lines):
    if lines and lines[0].strip() == "---":
        for j in range(1, min(len(lines), 40)):
            if lines[j].strip() in ("---", "..."):
                return lines[j + 1:]
    return lines


def render_markdown(builder, text, base_dir, skip_first_h1=False):
    lines = strip_front_matter(text.splitlines())
    n = len(lines)
    i = 0
    skipped_h1 = False

    while i < n:
        line = lines[i]
        s = line.strip()
        if not s:
            i += 1
            continue

        # 围栏代码块 ``` / ~~~
        m = re.match(r"^\s*(`{3,}|~{3,})(.*)$", line)
        if m:
            fence = m.group(1)
            closing = re.compile(r"^\s*" + re.escape(fence[0])
                                 + "{" + str(len(fence)) + r",}\s*$")
            i += 1
            body = []
            while i < n and not closing.match(lines[i]):
                body.append(lines[i])
                i += 1
            i += 1                                 # 跳过收尾围栏（没有则到底）
            builder.code("\n".join(body))
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.*?)\s*#*\s*$", line)
        if m:
            level, title = len(m.group(1)), m.group(2).strip()
            if skip_first_h1 and not skipped_h1 and level == 1:
                skipped_h1 = True
                i += 1
                continue
            builder.heading(level, title)
            i += 1
            continue

        # 分隔线
        if re.fullmatch(r"([-*_])(\s*\1){2,}", s):
            builder.rule()
            i += 1
            continue

        # 表格
        if "|" in line and i + 1 < n and is_table_sep(lines[i + 1]):
            header = split_row(line)
            aligns = sep_aligns(lines[i + 1])
            i += 2
            rows = []
            while i < n and lines[i].strip() and "|" in lines[i]:
                rows.append(split_row(lines[i]))
                i += 1
            builder.table(header, rows, aligns)
            continue

        # 引用
        if s.startswith(">"):
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(re.sub(r"^\s*>\s?", "", lines[i]).strip())
                i += 1
            buf = [b for b in buf if b]
            if buf:
                builder.quote(join_paragraph(buf, [False] * len(buf)))
            continue

        # 列表
        m = re.match(r"^(\s*)([-*+]|\d{1,3}[.)])\s+(.*)$", line)
        if m:
            indent = m.group(1).replace("\t", "  ")
            level = min(len(indent) // 2, 2)
            builder.list_item(level, m.group(3).strip(), m.group(2)[0].isdigit())
            i += 1
            continue

        # 独占一行的图片
        m = re.fullmatch(r"!\[([^\]]*)\]\(\s*<?([^)>\s]+)>?(?:\s+\"[^\"]*\")?\s*\)", s)
        if m:
            builder.image(m.group(1), m.group(2), base_dir)
            i += 1
            continue

        # 普通段落：连续多行合并，遇到空行或下一个块开始为止
        parts, hard = [], []
        while i < n:
            cur = lines[i]
            if not cur.strip():
                break
            if parts and is_block_start(lines, i):
                break
            raw = cur.rstrip()
            hard_break = raw.endswith("  ") or raw.endswith("\\")
            if raw.endswith("\\"):
                raw = raw[:-1]
            parts.append(raw.strip())
            hard.append(hard_break)
            i += 1
        if parts:
            builder.paragraph(join_paragraph(parts, hard))
        else:                                      # 兜底，避免死循环
            i += 1


# --------------------------------------------------------------- 文档骨架 --

def setup_document(doc, title):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(3.0)

    normal = doc.styles["Normal"]
    configure_style(normal, BODY_LATIN, BODY_CJK, BODY_SIZE)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    pf.widow_control = True

    for level in range(1, 5):
        style = doc.styles[f"Heading {level}"]
        configure_style(style, HEAD_LATIN, HEAD_CJK, HEAD_SIZES[level],
                        bold=True, color="1A1A1A")
        hpf = style.paragraph_format
        before, after = HEAD_SPACE[level]
        hpf.space_before = Pt(before)
        hpf.space_after = Pt(after)
        hpf.line_spacing = 1.2
        hpf.keep_with_next = True
        hpf.keep_together = True

    add_page_number(section)
    doc.core_properties.title = title


def add_page_number(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE ")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    fld.append(run)
    p._p.append(fld)


def add_cover(doc, title, subtitle, toc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(90)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(title)
    run.font.name = HEAD_LATIN
    set_run_east_asia(run, HEAD_CJK)
    run.font.size = Pt(26)
    run.font.bold = True

    if subtitle:
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.paragraph_format.space_after = Pt(24)
        srun = sp.add_run(subtitle)
        srun.font.name = HEAD_LATIN
        set_run_east_asia(srun, HEAD_CJK)
        srun.font.size = Pt(14)
        srun.font.color.rgb = RGBColor.from_string("595959")

    if toc:
        doc.add_paragraph().paragraph_format.page_break_before = True
        head = doc.add_paragraph(style="Heading 1")
        head.add_run("目录")
        p = doc.add_paragraph()
        run = p.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = ' TOC \\o "1-3" \\h \\z \\u '
        sep = OxmlElement("w:fldChar")
        sep.set(qn("w:fldCharType"), "separate")
        tip = OxmlElement("w:t")
        tip.text = "（在此处右键 → 更新域，即可生成目录）"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        for node in (begin, instr, sep, tip, end):
            run._r.append(node)


# --------------------------------------------------------------- 文件收集 --

LINK_RE = re.compile(r"!?\[[^\]]*\]\(\s*<?([^)>\s]+)>?(?:\s+\"[^\"]*\")?\s*\)")


def read_text(path):
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def natural_key(text):
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r"(\d+)", text)]


def links_from_index(index_path):
    files, seen = [], set()
    for m in LINK_RE.finditer(read_text(index_path)):
        target = m.group(1)
        if target.startswith(("http://", "https://", "mailto:", "tel:", "#")):
            continue
        target = target.split("#")[0].split("?")[0].strip()
        if not target.lower().endswith(".md"):
            continue
        candidate = index_path.parent / target
        if candidate.is_file():
            resolved = candidate.resolve()
            if resolved not in seen:
                seen.add(resolved)
                files.append(resolved)
    return files


def scan_folder(root, skip):
    files = []
    for path in root.rglob("*.md"):
        rel = path.relative_to(root)
        if any(part.startswith(".") for part in rel.parts):
            continue
        resolved = path.resolve()
        if resolved in skip:
            continue
        files.append(resolved)
    files.sort(key=lambda p: natural_key(str(p.relative_to(root))))
    return files


def first_h1(text):
    for line in text.splitlines():
        m = re.match(r"^#\s+(.*?)\s*#*\s*$", line)
        if m and m.group(1).strip():
            return m.group(1).strip()
    return None


def find_subtitle(text):
    """取索引里紧跟在 H1 后面的第一行普通文字，作为副标题。"""
    lines = strip_front_matter(text.splitlines())
    after_title = False
    for line in lines:
        s = line.strip()
        if not s:
            continue
        if not after_title:
            if re.match(r"^#\s", line):
                after_title = True
            continue
        if (s.startswith(("#", "|", ">", "```", "~~~"))
                or re.match(r"^([-*+]|\d{1,3}[.)])\s+", s)):
            return None
        return s
    return None


def safe_filename(name):
    cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', "_", name).strip(" .")
    return cleaned or "未命名文档"


def relative_label(path, root):
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


# ------------------------------------------------------------------- main --

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="把一整篇文章的 markdown 章节合并导出成一个 Word 文档。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python md2docx.py\n"
               "  python md2docx.py -r 某本书 --title 书名 -o 某本书.docx\n"
               "  python md2docx.py --toc --no-page-break\n")
    parser.add_argument("files", nargs="*", help="要合并的 markdown 文件，按给出顺序拼接")
    parser.add_argument("-r", "--root", default=".", help="文章根目录，默认当前目录")
    parser.add_argument("-i", "--index", default=None,
                        help="索引文件，默认自动找 index.md / README.md")
    parser.add_argument("-o", "--out", default=None, help="输出文件，默认 <标题>.docx")
    parser.add_argument("-t", "--title", default=None, help="文档标题")
    parser.add_argument("--no-page-break", action="store_true", help="章节之间不分页")
    parser.add_argument("--toc", action="store_true", help="插入目录域（打开后右键更新）")
    parser.add_argument("--keep-index", action="store_true", help="把索引文件的内容也写进正文")
    args = parser.parse_args(argv)

    root = Path(args.root).expanduser().resolve()
    if not root.is_dir():
        sys.stderr.write(f"目录不存在：{root}\n")
        return 2

    index_path = None
    if args.index:
        index_path = (root / args.index).resolve()
        if not index_path.is_file():
            sys.stderr.write(f"索引文件不存在：{index_path}\n")
            return 2
    else:
        for name in ("index.md", "Index.md", "INDEX.md", "README.md", "readme.md"):
            if (root / name).is_file():
                index_path = (root / name).resolve()
                break

    index_text = read_text(index_path) if index_path else ""

    if args.files:
        sources = [(root / f).resolve() for f in args.files]
        missing = [p for p in sources if not p.is_file()]
        if missing:
            sys.stderr.write("找不到文件：" + "、".join(str(m) for m in missing) + "\n")
            return 2
        sources = [p for p in sources if p != index_path]
    elif index_text:
        sources = links_from_index(index_path)
        if not sources:
            sys.stderr.write(f"{index_path.name} 里没有找到 .md 链接，改为扫描目录。\n")
            sources = scan_folder(root, {index_path})
    else:
        sources = scan_folder(root, set())

    if not sources:
        sys.stderr.write("没有找到任何 markdown 文件。\n")
        return 1

    title = (args.title
             or first_h1(index_text)
             or first_h1(read_text(sources[0]))
             or root.name)
    subtitle = find_subtitle(index_text) if index_text else None

    if args.out:
        out_path = Path(args.out).expanduser()
        if not out_path.is_absolute():
            out_path = (root / out_path).resolve()
    else:
        out_path = root / f"{safe_filename(title)}.docx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    setup_document(doc, title)
    add_cover(doc, title, subtitle, args.toc)

    builder = Builder(doc)
    body_started = False

    if index_text and args.keep_index:
        render_markdown(builder, index_text, index_path.parent, skip_first_h1=True)
        body_started = True

    for src in sources:
        if body_started:
            builder.pending_break = not args.no_page_break
        render_markdown(builder, read_text(src), src.parent)
        body_started = True

    doc.add_paragraph()                       # 末尾留一段，避免以表格/图结尾
    doc.save(str(out_path))

    print(f"标题：{title}")
    if subtitle:
        print(f"副标题：{subtitle}")
    print(f"章节：{len(sources)} 个文件")
    for src in sources:
        print(f"  - {relative_label(src, root)}")
    print(f"输出：{out_path}")
    print(f"大小：{out_path.stat().st_size / 1024:.1f} KB")

    try:                                      # 重新打开一次，确认文件没写坏
        check = Document(str(out_path))
        headings = sum(1 for p in check.paragraphs if p.style.name.startswith("Heading"))
        print(f"校验：段落 {len(check.paragraphs)} 个，标题 {headings} 个，"
              f"表格 {len(check.tables)} 个")
        print(f"明细：标题 {builder.counts['heading']}、段落 {builder.counts['paragraph']}、"
              f"代码块 {builder.counts['code']}、表格 {builder.counts['table']}、"
              f"列表项 {builder.counts['list']}、图片 {builder.counts['image']}")
    except Exception as exc:                  # noqa: BLE001
        sys.stderr.write(f"警告：生成的文件重新打开失败 —— {exc}\n")
        return 3

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
