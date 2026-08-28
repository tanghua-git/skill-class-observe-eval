#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx_report.py v2 — 课堂观察评课报告 Markdown → Word 转换脚本（专业排版版）

用法：
    python3 md2docx_report.py <md路径> <docx路径>

排版规格（与 SKILL.md 4.4 节一致）：
【主题色】深蓝 #1F4E79（标题/表头/引用线），浅蓝底 #EEF3FA / #F2F6FB（引用块/斑马纹）
【标题】H1 黑体20pt居中深蓝+底部粗线；H2 黑体14pt深蓝+底部细线；H3 黑体12pt深蓝；H4 黑体11pt黑
【正文】仿宋11pt，1.4倍行距，段后4pt，首行缩进2字符（列表/表格/引用除外）
【引用块】左侧深蓝竖线 + 浅蓝底色 + 灰黑字10.5pt
【表格】表头深蓝底白字黑体9.5pt（跨页自动重复）；数据行白/#F2F6FB斑马纹；灰边框0.5pt；
        单元格垂直居中；表格宽度100%撑满版心
【分隔线】markdown 的 --- 渲染为浅灰细线（不再静默忽略）
【页眉】居中"课堂观察评课报告"灰字 + 底部细线
【页脚】居中"第 X 页 共 Y 页"（PAGE/NUMPAGES 域，自动更新）
【其他】整行粗体（如"表 1-1-1 ……"表题）渲染为黑体10.5pt；复选框清单渲染 □/☑ 前缀
"""

import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HEI = "黑体"
FANGSONG = "仿宋"

# 主题配色
C_PRIMARY = "1F4E79"      # 深蓝主色
C_PRIMARY_RGB = RGBColor(0x1F, 0x4E, 0x79)
C_QUOTE_BG = "EEF3FA"     # 引用块底色
C_ZEBRA = "F2F6FB"        # 表格斑马纹
C_BORDER = "A6A6A6"       # 表格边框灰
C_RULE = "C9C9C9"         # 分隔线浅灰
C_GRAY_TEXT = RGBColor(0x59, 0x59, 0x59)


# ---------------------------------------------------------------- 基础工具

def set_run_font(run, name, size_pt, bold=False, color=None):
    """同时设置西文与东亚字体。"""
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:eastAsia"), name)


def add_runs_with_bold(paragraph, text, font, size_pt, base_bold=False, color=None):
    """解析行内 **粗体**，逐段写入 run。"""
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, font, size_pt, bold=True, color=color)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, font, size_pt, bold=base_bold, color=color)


def para_bottom_border(paragraph, color=C_PRIMARY, sz="12", space="4"):
    """段落底部边框（用于 H1/H2 标题下划线与分隔线）。"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), space)
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.insert(0, pBdr)


def style_quote_block(paragraph):
    """引用块：左侧深蓝竖线 + 浅蓝底色。"""
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), C_QUOTE_BG)
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), C_PRIMARY)
    pBdr.append(left)
    pPr.insert(0, shd)
    pPr.insert(0, pBdr)


# ---------------------------------------------------------------- 表格

def set_cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), C_BORDER)
        borders.append(el)
    tc_pr.append(borders)


def shade_cell(cell, hexcolor):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    tc_pr.append(shd)


def vcenter_cell(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tc_pr.append(vAlign)


def set_row_repeat_header(row):
    """表头行跨页自动重复。"""
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def set_table_fullwidth(table):
    """表格宽度 100% 撑满版心。"""
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"), "5000")


def write_cell(cell, text, is_header=False):
    cell.text = ""
    para = cell.paragraphs[0]
    para.paragraph_format.space_before = Pt(2)
    para.paragraph_format.space_after = Pt(2)
    para.paragraph_format.line_spacing = 1.15
    if is_header:
        add_runs_with_bold(para, text, HEI, 9.5, base_bold=True,
                           color=RGBColor(0xFF, 0xFF, 0xFF))
    else:
        add_runs_with_bold(para, text, FANGSONG, 9.5)
    set_cell_borders(cell)
    vcenter_cell(cell)
    if is_header:
        shade_cell(cell, C_PRIMARY)
    elif cell._tc.getparent().index(cell._tc) >= 0:
        pass  # 斑马纹按行处理


def add_table(doc, table_lines):
    """一次性写入完整表格块：第一行为表头，跳过分隔行，数据行隔行斑马纹。"""
    rows = []
    for line in table_lines:
        cells = parse_row(line)
        if is_separator_row(cells):
            continue
        rows.append(cells)
    if not rows:
        return

    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_fullwidth(table)

    for i, row_cells in enumerate(rows):
        for j in range(ncols):
            text = row_cells[j] if j < len(row_cells) else ""
            cell = table.cell(i, j)
            write_cell(cell, text, is_header=(i == 0))
            if i > 0 and (i - 1) % 2 == 1:  # 数据行斑马纹：第2、4…数据行
                shade_cell(cell, C_ZEBRA)
    set_row_repeat_header(table.rows[0])

    # 表格后小空距（不额外占行高）
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    for run in (p.add_run(""),):
        set_run_font(run, FANGSONG, 6)


# ---------------------------------------------------------------- 页眉页脚

def add_field_run(paragraph, instr, font, size, color):
    """插入 Word 域（PAGE / NUMPAGES），页码自动更新。"""
    r1 = paragraph.add_run()
    set_run_font(r1, font, size, color=color)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    r1._element.append(fld1)
    r2 = paragraph.add_run()
    set_run_font(r2, font, size, color=color)
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = instr
    r2._element.append(instrText)
    r3 = paragraph.add_run()
    set_run_font(r3, font, size, color=color)
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    r3._element.append(fld2)


def add_header_footer(doc):
    section = doc.sections[0]

    # 页眉：居中标题 + 底部细线
    header_para = section.header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header_para.add_run("课堂观察评课报告")
    set_run_font(run, HEI, 9, color=C_GRAY_TEXT)
    para_bottom_border(header_para, color=C_RULE, sz="6", space="2")

    # 页脚：居中"第 X 页 共 Y 页"（域自动更新）
    footer_para = section.footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_para.add_run("第 ")
    set_run_font(run, FANGSONG, 9, color=C_GRAY_TEXT)
    add_field_run(footer_para, "PAGE", FANGSONG, 9, C_GRAY_TEXT)
    run = footer_para.add_run(" 页 · 共 ")
    set_run_font(run, FANGSONG, 9, color=C_GRAY_TEXT)
    add_field_run(footer_para, "NUMPAGES", FANGSONG, 9, C_GRAY_TEXT)
    run = footer_para.add_run(" 页")
    set_run_font(run, FANGSONG, 9, color=C_GRAY_TEXT)


# ---------------------------------------------------------------- Markdown 解析

def is_table_line(line):
    s = line.strip()
    return s.startswith("|") and s.endswith("|") and s.count("|") >= 2


def is_separator_row(cells):
    return all(re.fullmatch(r":?-{2,}:?", c.strip()) for c in cells if c.strip())


def parse_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def add_heading(doc, text, level):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(6)
        pf.space_after = Pt(14)
        add_runs_with_bold(p, text, HEI, 20, base_bold=True, color=C_PRIMARY_RGB)
        para_bottom_border(p, color=C_PRIMARY, sz="18", space="6")
    elif level == 2:
        pf.space_before = Pt(16)
        pf.space_after = Pt(8)
        pf.keep_with_next = True
        add_runs_with_bold(p, text, HEI, 14, base_bold=True, color=C_PRIMARY_RGB)
        para_bottom_border(p, color="BDD0E9", sz="8", space="3")
    elif level == 3:
        pf.space_before = Pt(10)
        pf.space_after = Pt(4)
        pf.keep_with_next = True
        add_runs_with_bold(p, text, HEI, 12, base_bold=True, color=C_PRIMARY_RGB)
    else:
        pf.space_before = Pt(8)
        pf.space_after = Pt(3)
        add_runs_with_bold(p, text, HEI, 11, base_bold=True)
    return p


def convert(md_path, docx_path):
    with open(md_path, encoding="utf-8") as f:
        lines = f.read().splitlines()

    doc = Document()
    for section in doc.sections:
        section.page_width = Cm(21.0)      # A4 纸宽
        section.page_height = Cm(29.7)     # A4 纸高
        section.top_margin = Cm(2.4)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(2.6)
        section.right_margin = Cm(2.6)
    add_header_footer(doc)

    i = 0
    while i < len(lines):
        line = lines[i]

        # 表格块：连续收集后一次性写入
        if is_table_line(line):
            block = []
            while i < len(lines) and is_table_line(lines[i]):
                block.append(lines[i])
                i += 1
            add_table(doc, block)
            continue

        s = line.strip()

        if not s:
            i += 1
            continue

        # 标题（H1-H4）
        if s.startswith("#### "):
            add_heading(doc, s[5:], 4)
        elif s.startswith("### "):
            add_heading(doc, s[4:], 3)
        elif s.startswith("## "):
            add_heading(doc, s[3:], 2)
        elif s.startswith("# "):
            add_heading(doc, s[2:], 1)
        # 分隔线 ---：渲染为浅灰细线
        elif re.fullmatch(r"-{3,}", s):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            para_bottom_border(p, color=C_RULE, sz="6", space="1")
        # 引用块（> 开头）：左竖线 + 浅蓝底
        elif s.startswith(">"):
            content = s.lstrip(">").strip()
            if not content:
                i += 1
                continue
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(0.5)
            pf.right_indent = Cm(0.3)
            pf.space_before = Pt(3)
            pf.space_after = Pt(3)
            pf.line_spacing = 1.3
            add_runs_with_bold(p, content, FANGSONG, 10.5, color=RGBColor(0x33, 0x33, 0x33))
            style_quote_block(p)
        # 复选框清单（- [ ] / - [x]）——须在无序列表前判断
        elif re.match(r"^[-*]\s+\[[ xX]\]\s*", s):
            checked = re.match(r"^[-*]\s+\[([ xX])\]", s).group(1).lower() == "x"
            mark = "☑ " if checked else "□ "
            body = re.sub(r"^[-*]\s+\[[ xX]\]\s*", "", s)
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(0.75)
            pf.line_spacing = 1.35
            pf.space_after = Pt(2)
            add_runs_with_bold(p, mark + body, FANGSONG, 11)
        # 有序列表
        elif re.match(r"^\d+[.、]\s*", s):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(0.75)
            pf.line_spacing = 1.35
            pf.space_after = Pt(2)
            add_runs_with_bold(p, s, FANGSONG, 11)
        # 无序列表（- / * 开头）
        elif re.match(r"^[-*]\s+", s):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(0.75)
            pf.line_spacing = 1.35
            pf.space_after = Pt(2)
            add_runs_with_bold(p, "• " + re.sub(r"^[-*]\s+", "", s), FANGSONG, 11)
        # 整行粗体（表题/重点句，如 "**表 1-1-1 课堂思维观察表**"）
        elif re.fullmatch(r"\*\*[^*]+\*\*", s):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(6)
            pf.space_after = Pt(2)
            pf.keep_with_next = True
            add_runs_with_bold(p, s[2:-2], HEI, 10.5, base_bold=True)
        # 普通正文
        else:
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.first_line_indent = Pt(22)  # 首行缩进2字符
            pf.line_spacing = 1.4
            pf.space_after = Pt(4)
            add_runs_with_bold(p, s, FANGSONG, 11)

        i += 1

    doc.save(docx_path)


def main():
    if len(sys.argv) != 3:
        print("用法: python3 md2docx_report.py <md路径> <docx路径>")
        sys.exit(1)
    convert(sys.argv[1], sys.argv[2])
    print(f"OK: {sys.argv[2]}")


if __name__ == "__main__":
    main()
