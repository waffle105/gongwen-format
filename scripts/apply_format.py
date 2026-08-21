#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_format.py — 按《市委常委会会议材料（汇报材料）格式规范》批量格式化公文 docx。

用法：
    python apply_format.py <输入.docx> [输出.docx] [选项]

选项：
    --title-font NAME    主标题字体名，默认「方正小标宋简体」
                         部分机器实际为「方正小标宋_GBK」，可据此切换
    --no-cover           不自动处理封面标题区（仅格式化正文+页面）

说明：
    - 自动完成：页面边距/页脚、正文各级标题与正文的字体字号行距、
      数字与英文字母 Times New Roman、封面标题区（主标题/副标题/落款）。
    - 页面格式（边距、页脚）为确定值，正文层级通过文本模式识别。
    - 封面标题区通过「正文起点之前的所有段落」启发式识别，若文档结构特殊，
      请输出后人工核对封面部分。
"""

import re
import sys
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ===================== 常量：格式规范 =====================
# 字号（磅）
ERHAO = 22   # 二号
SANHAO = 16  # 三号
SIHAO = 14   # 四号

# 字体
FONT_TITLE = '方正小标宋简体'   # 大标题（可用 --title-font 覆盖）
FONT_TITLE_ALT = '方正小标宋_GBK'
FONT_KAI = '楷体_GB2312'       # 楷体
FONT_FANG = '仿宋_GB2312'      # 仿宋
FONT_HEI = '黑体'              # 黑体
FONT_ASCII = 'Times New Roman' # 数字/英文

# 行距（磅，固定值）
LINE_TITLE = 35   # 标题区行距
LINE_BODY = 28    # 正文行距

# 页面（厘米）
MARGIN_TOP = 3.0
MARGIN_BOTTOM = 3.0
MARGIN_LEFT = 2.7
MARGIN_RIGHT = 2.7
FOOTER_DIST = 2.4


# ===================== 层级识别 =====================
RE_LEVEL1 = re.compile(r'^[一二三四五六七八九十百]+、')      # 一、
RE_LEVEL2 = re.compile(r'^（[一二三四五六七八九十百]+）')     # （一）
RE_LEVEL3 = re.compile(r'^\d+[．.·]')                       # 1． / 1. / 1·
RE_LEVEL4 = re.compile(r'^（\d+）')                          # （1）


def classify_paragraph(text):
    """返回正文层级：1/2/3/4，或 0 表示正文，或 None 表示空行。"""
    t = text.strip()
    if not t:
        return None
    if RE_LEVEL4.match(t):
        return 4
    if RE_LEVEL3.match(t):
        return 3
    if RE_LEVEL2.match(t):
        return 2
    if RE_LEVEL1.match(t):
        return 1
    return 0


# ===================== 底层设置函数 =====================
def set_run_font(run, east_asia, ascii_font=FONT_ASCII, size=None, bold=None):
    """设置 run 的中文（eastAsia）与西文（ascii）字体、字号、加粗。"""
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), east_asia)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def set_paragraph_format(p, alignment=None, line_spacing=None, first_line_indent=None):
    """设置段落对齐、固定行距、首行缩进。"""
    pf = p.paragraph_format
    if alignment is not None:
        pf.alignment = alignment
    if line_spacing is not None:
        # 固定行距（磅）
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_spacing)
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent


def set_first_line_indent_chars(p, chars=2):
    """设置首行缩进为指定字符数（默认 2 字符）。"""
    pPr = p._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(chars * 100))
    ind.set(qn('w:firstLine'), str(int(chars * 320)))  # 兼容值：三号16磅×2字符=32磅=640twips


def format_body_paragraph(p, level, title_font, indent=True):
    """按正文层级格式化一个段落。level: 1/2/3/4 或 0（正文）。"""
    text = p.text.strip()
    if not text:
        return

    # 确定字体与是否加粗
    if level == 1:
        ea_font, bold = FONT_HEI, False
    elif level == 2:
        ea_font, bold = FONT_KAI, True
    elif level == 3:
        ea_font, bold = FONT_FANG, True
    elif level == 4:
        ea_font, bold = FONT_FANG, False
    else:  # 正文
        ea_font, bold = FONT_FANG, False

    set_paragraph_format(p, line_spacing=LINE_BODY)
    if indent and level == 0:
        # 正文段落首行缩进 2 字符（公文惯例）
        set_first_line_indent_chars(p, 2)
    for run in p.runs:
        set_run_font(run, ea_font, size=SANHAO, bold=bold)


def setup_page(doc):
    """设置页面边距与页脚距离（A4 默认，仅改边距）。"""
    for sec in doc.sections:
        sec.top_margin = Cm(MARGIN_TOP)
        sec.bottom_margin = Cm(MARGIN_BOTTOM)
        sec.left_margin = Cm(MARGIN_LEFT)
        sec.right_margin = Cm(MARGIN_RIGHT)
        sec.footer_distance = Cm(FOOTER_DIST)


def setup_page_number(doc):
    """在页脚添加「— 1 —」样式的页码（四号宋体，居中）。
    外侧（奇偶页）需在 Word 中手动启用「奇偶页不同」后微调。"""
    for sec in doc.sections:
        footer = sec.footer
        footer.is_linked_to_previous = False
        # 清空已有内容
        for p in list(footer.paragraphs):
            p.clear()
        if footer.paragraphs:
            p = footer.paragraphs[0]
        else:
            p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 前半部分「— 」
        run1 = p.add_run('— ')
        set_run_font(run1, FONT_FANG, size=SIHAO)
        # 页码字段
        run_field = p.add_run()
        set_run_font(run_field, FONT_FANG, size=SIHAO)
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run_field._element.append(fldChar1)
        run_field._element.append(instrText)
        run_field._element.append(fldChar2)
        # 后半部分「 —」
        run2 = p.add_run(' —')
        set_run_font(run2, FONT_FANG, size=SIHAO)


def _set_cover_para(p, font, size):
    """设置封面区段落：居中 + 35 磅行距 + 指定字体字号。"""
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=LINE_TITLE)
    for run in p.runs:
        set_run_font(run, font, size=size, bold=False)


def format_cover(doc, title_font):
    """启发式格式化封面标题区（正文起点之前的段落）。
    返回 (正文起点索引, 需按正文处理的封面区段落索引列表)。"""
    paras = doc.paragraphs

    # 找到正文起点：第一个匹配层级标题的段落
    body_start = len(paras)
    for i, p in enumerate(paras):
        if classify_paragraph(p.text) in (1, 2, 3, 4):
            body_start = i
            break

    cover_idx = [i for i in range(body_start) if paras[i].text.strip()]

    title_assigned = False
    date_seen = False
    extra_body = []  # 主标题之后的普通段落（如引言）→ 按正文处理

    for i in cover_idx:
        p = paras[i]
        t = p.text.strip()
        if '会议材料' in t or '汇报材料' in t:
            # 封面大标题（四号）
            _set_cover_para(p, title_font, SIHAO)
        elif date_seen:
            # 日期之后 → 落款单位
            _set_cover_para(p, FONT_KAI, SANHAO)
        elif '年' in t:
            # 落款日期（如「2026年8月」「××年××月」）
            _set_cover_para(p, FONT_KAI, SANHAO)
            date_seen = True
        elif t.startswith('（') or '通过' in t:
            # 副标题
            _set_cover_para(p, FONT_KAI, SANHAO)
        elif not title_assigned:
            # 第一个普通段落 → 主标题（二号）
            _set_cover_para(p, title_font, ERHAO)
            title_assigned = True
        else:
            # 主标题之后的普通段落 → 引言/说明，按正文处理
            extra_body.append(i)

    return body_start, extra_body


# ===================== 主流程 =====================
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    in_path = sys.argv[1]
    args = sys.argv[2:]

    # 解析输出路径与选项
    title_font = FONT_TITLE
    do_cover = True
    out_path = None
    i = 0
    while i < len(args):
        a = args[i]
        if a == '--title-font' and i + 1 < len(args):
            title_font = args[i + 1]; i += 2
        elif a == '--no-cover':
            do_cover = False; i += 1
        elif not a.startswith('--') and out_path is None:
            out_path = a; i += 1
        else:
            i += 1

    if out_path is None:
        base, ext = os.path.splitext(in_path)
        out_path = base + '_formatted' + ext

    doc = Document(in_path)

    # 1. 页面设置
    setup_page(doc)
    setup_page_number(doc)

    # 2. 封面标题区（返回正文起点与需按正文处理的封面段落）
    body_start = 0
    extra_body = []
    if do_cover:
        body_start, extra_body = format_cover(doc, title_font)

    # 3. 正文层级（处理正文起点之后的段落，以及封面区中按正文处理的段落）
    body_indices = list(extra_body) + list(range(body_start, len(doc.paragraphs)))
    for i in body_indices:
        p = doc.paragraphs[i]
        lv = classify_paragraph(p.text)
        if lv in (1, 2, 3, 4, 0):
            format_body_paragraph(p, lv, title_font)

    doc.save(out_path)
    print(f'[OK] 已格式化并保存：{out_path}')
    print(f'     主标题字体：{title_font}')
    print(f'     封面处理：{"是" if do_cover else "否"}')
    print(f'     提示：请人工核对封面标题区与页码「外侧」设置。')


if __name__ == '__main__':
    main()
