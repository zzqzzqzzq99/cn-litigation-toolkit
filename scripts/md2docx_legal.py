#!/usr/bin/env python3
"""
md2docx_legal.py — 法律文书 Markdown → Word 转换器

用法:
    python3 md2docx_legal.py [--spec format-spec.md] input.md output.docx

功能:
    1. 读取 format-spec.md 排版参数（字体/字号/行距/边距等）
    2. C1-C10 文本清洗（Markdown 标记清除、引号转换、空行压缩等）
    3. S1-S6 结构规范化（当事人分隔、落款右对齐、此致格式等）
    4. python-docx 精确生成 Word（宋体四号、固定行距25磅、首行缩进2字符）
    5. 内置 V1-V12 后验校验
    6. 4级降级链（md2docx → 外部MCP → pandoc → 纯Markdown兜底）

参数源: format-spec.md（套件根目录）
依赖: python-docx (pip install python-docx)
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Any

# 共享模块（消除 md2docx_legal.py / verify_docx.py 重复代码）
from legal_common import (
    DEFAULT_PARAMS as _COMMON_DEFAULT_PARAMS,
    SIGNATURE_KEYWORDS, PARTY_KEYWORDS, SECTION_TITLES, DATE_PATTERN,
    load_spec, is_title_paragraph, is_signature_paragraph,
    get_run_font_info, make_result as _make_result_common,
    force_utf8_output,
)

# ---------------------------------------------------------------------------
# 依赖检查
# ---------------------------------------------------------------------------
try:
    from docx import Document
    from docx.shared import Pt, Twips
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("错误: 需要安装 python-docx 库。请执行: pip install python-docx", file=sys.stderr)
    sys.exit(2)


# ===========================================================================
# 第一部分：排版参数（FormatSpec）
# ===========================================================================

# 默认参数 —— 使用共享模块中的定义
DEFAULT_PARAMS = _COMMON_DEFAULT_PARAMS


# ===========================================================================
# 第二部分：文本清洗（C1-C10）
# ===========================================================================



def _preprocess_text(text: str) -> str:
    """预处理：清除 BOM、零宽字符、统一换行符。"""
    # 清除 BOM
    if text.startswith("\ufeff"):
        text = text[1:]
    # 清除零宽字符（零宽空格、零宽连接符、零宽非连接符、BOM）
    for zw in ["\u200b", "\u200c", "\u200d", "\ufeff", "\u2060"]:
        text = text.replace(zw, "")
    # 统一换行符（Windows \r\n → \n, 旧 Mac \r → \n）
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def clean_text(text: str) -> str:
    """执行 C1-C10 文本清洗，返回清洗后的文本。"""
    # 预处理：BOM、零宽字符、换行符统一
    text = _preprocess_text(text)
    lines = text.split("\n")
    cleaned_lines = []

    for line in lines:
        # 预清洗：全角空格处理（保留缩进语义但规范化）
        if line.strip():
            # 非空行：全角空格转为两个普通空格（保持视觉缩进）
            line = line.replace("\u3000", "  ")
            # 连续空格压缩为单个（但保留行首缩进）
            stripped = line.lstrip()
            indent_len = len(line) - len(stripped)
            stripped = re.sub(r"[ \t]{2,}", " ", stripped)
            line = " " * indent_len + stripped
        # C1: Markdown 标题标记清除（行首 # ）
        line = re.sub(r"^#{1,6}\s+", "", line)
        # C2: 加粗标记清除
        line = re.sub(r"\*\*(.+?)\*\*", r"\1", line)
        line = re.sub(r"__(.+?)__", r"\1", line)
        # C3: 列表标记规范化（暂保留，后续结构化处理）
        # C4: HTML 上标清除
        line = re.sub(r"<sup>(.*?)</sup>", r"\1", line, flags=re.IGNORECASE)
        # C5: HTML 标签清除
        line = re.sub(r"<[^>]+>", "", line)
        # C6: Markdown 引用标记清除
        line = re.sub(r"^>\s+", "", line)
        # C7: 水平分割线清除
        if re.match(r"^\s*([-*_])\1{2,}\s*$", line):
            line = ""
        # C8: 直双引号转中文
        line = _convert_double_quotes(line)
        # C9: 直单引号转中文
        line = _convert_single_quotes(line)
        cleaned_lines.append(line)

    # C10: 连续空行压缩（3+ 空行 → 2 空行）
    result = _compress_blank_lines(cleaned_lines)

    return "\n".join(result)


def _convert_double_quotes(text: str) -> str:
    """直双引号 → 中文弯引号。"""
    result = []
    open_quote = True
    for ch in text:
        if ch == '"':
            result.append("\u201C" if open_quote else "\u201D")
            open_quote = not open_quote
        else:
            result.append(ch)
    return "".join(result)


def _convert_single_quotes(text: str) -> str:
    """直单引号 → 中文弯引号（仅在成对时转换）。"""
    # 简单策略：交替转换
    result = []
    open_quote = True
    for ch in text:
        if ch == "'":
            result.append("\u2018" if open_quote else "\u2019")
            open_quote = not open_quote
        else:
            result.append(ch)
    return "".join(result)


def _compress_blank_lines(lines: list[str]) -> list[str]:
    """连续 3+ 空行压缩为 2 空行。"""
    result = []
    blank_count = 0
    for line in lines:
        if line.strip() == "":
            blank_count += 1
            if blank_count <= 2:
                result.append(line)
        else:
            blank_count = 0
            result.append(line)
    # 去除末尾多余空行
    while result and result[-1].strip() == "":
        result.pop()
    return result


# ===========================================================================
# 第三部分：段落分类与结构规范化（S1-S6）
# ===========================================================================

class ParagraphType:
    TITLE = "title"                    # 文书标题
    PARTY_BLOCK = "party_block"        # 当事人板块
    SECTION_HEADING = "section_heading"  # 段落标题（如"事实与理由："）
    BODY = "body"                      # 正文段落
    CIZHI = "cizhi"                    # "此致"
    COURT_NAME = "court_name"          # 法院名称
    SIGNATURE = "signature"            # 签名行
    DATE = "date"                      # 日期行
    ATTACHMENT = "attachment"          # 附件
    BLANK = "blank"                    # 空行


def classify_paragraphs(lines: list[str]) -> list[tuple[str, str]]:
    """将每行分类为段落类型，返回 (type, text) 列表。

    落款识别采用 SKILL.md §4.1 的三条互补路径：
    1. 通用落款关键词白名单
    2. 落款前瞻判定（当事人关键词 + 下一行为日期 → 落款）
    3. 日期行留白兼容
    """
    result = []
    total = len(lines)
    in_signature_area = False
    in_attachment_area = False

    for idx, line in enumerate(lines):
        stripped = line.strip()

        # 空行
        if not stripped:
            result.append((ParagraphType.BLANK, ""))
            continue

        # S3: 附件区域检测
        if re.match(r"^[附附]\s*[件件]?[：:]", stripped) or stripped in ("附件", "附"):
            in_attachment_area = True
            result.append((ParagraphType.ATTACHMENT, stripped))
            continue

        if in_attachment_area:
            result.append((ParagraphType.ATTACHMENT, stripped))
            continue

        # 首个非空行视为标题
        if not any(r[0] not in (ParagraphType.BLANK,) for r in result):
            if len(stripped) <= 30 and not stripped.endswith("：") and not stripped.endswith(":"):
                result.append((ParagraphType.TITLE, stripped))
                continue

        # 落款区域检测（一旦进入签名区，后续行都按签名/日期处理）
        if in_signature_area:
            if DATE_PATTERN.match(stripped):
                result.append((ParagraphType.DATE, stripped))
                continue
            else:
                result.append((ParagraphType.SIGNATURE, stripped))
                continue

        # 落款关键词白名单（路径1）
        is_signature = False
        for kw in SIGNATURE_KEYWORDS:
            if stripped.startswith(kw) and (kw + "：" in stripped or kw + ":" in stripped):
                is_signature = True
                break

        # 落款前瞻判定（路径2）：当事人关键词 + 下一行为日期
        if not is_signature:
            for kw in PARTY_KEYWORDS:
                if stripped.startswith(kw) and (kw + "：" in stripped or kw + ":" in stripped):
                    # 向下前瞻
                    next_idx = idx + 1
                    while next_idx < total and not lines[next_idx].strip():
                        next_idx += 1
                    if next_idx < total and DATE_PATTERN.match(lines[next_idx].strip()):
                        is_signature = True
                    break

        if is_signature:
            in_signature_area = True
            result.append((ParagraphType.SIGNATURE, stripped))
            continue

        # 日期行（路径3：留白兼容）
        if DATE_PATTERN.match(stripped):
            in_signature_area = True
            result.append((ParagraphType.DATE, stripped))
            continue

        # 此致
        if stripped.startswith("此致"):
            result.append((ParagraphType.CIZHI, stripped))
            continue

        # 法院名称（此致下一行）
        if result and result[-1][0] == ParagraphType.CIZHI:
            result.append((ParagraphType.COURT_NAME, stripped))
            continue

        # 段落标题（中文编号 或 已知标题词）
        if re.match(r"^[一二三四五六七八九十]+、", stripped):
            result.append((ParagraphType.SECTION_HEADING, stripped))
            continue

        # 段落标题（已知标题词 + 冒号）
        for title_word in SECTION_TITLES:
            if stripped.startswith(title_word) and (stripped.endswith("：") or stripped.endswith(":")):
                result.append((ParagraphType.SECTION_HEADING, stripped))
                break
        else:
            # 当事人板块检测
            is_party = False
            for kw in PARTY_KEYWORDS:
                if stripped.startswith(kw) and (kw + "：" in stripped or kw + ":" in stripped):
                    is_party = True
                    break

            if is_party:
                result.append((ParagraphType.PARTY_BLOCK, stripped))
            else:
                result.append((ParagraphType.BODY, stripped))
                continue

    return result


def normalize_structure(lines: list[str]) -> list[tuple[str, str]]:
    """执行 S1-S6 结构规范化，返回分类后的段落列表。"""
    # 先清洗
    cleaned = clean_text("\n".join(lines)).split("\n")
    # 分类
    classified = classify_paragraphs(cleaned)

    # S1: 当事人板块分隔 —— 在不同当事人板块之间确保有空行
    normalized = []
    prev_type = None
    for ptype, text in classified:
        if ptype == ParagraphType.BLANK:
            normalized.append((ptype, text))
            prev_type = ptype
            continue

        # 当事人板块前插入空行（如果前一个不是空行且不是标题）
        if (ptype == ParagraphType.PARTY_BLOCK
            and prev_type not in (None, ParagraphType.BLANK, ParagraphType.TITLE)
            and prev_type != ParagraphType.PARTY_BLOCK):
            normalized.append((ParagraphType.BLANK, ""))

        normalized.append((ptype, text))
        prev_type = ptype

    return normalized


# ===========================================================================
# 第四部分：Word 文档生成（DocxBuilder）
# ===========================================================================

class DocxBuilder:
    """使用 python-docx 精确生成法律文书 Word 文档。"""

    def __init__(self, params: dict[str, Any]):
        self.params = params
        self.doc = Document()

    def build(self, paragraphs: list[tuple[str, str]]) -> Document:
        """根据分类段落列表生成 Word 文档。

        异常处理策略：单个段落处理失败不影响整体转换，
        降级为普通正文段落并记录错误。
        """
        self._setup_page()
        self._setup_default_style()
        self._build_errors = []

        for ptype, text in paragraphs:
            try:
                if ptype == ParagraphType.BLANK:
                    # 空段落（保留间距）
                    self.doc.add_paragraph()
                elif ptype == ParagraphType.TITLE:
                    self._add_title(text)
                elif ptype == ParagraphType.SECTION_HEADING:
                    self._add_section_heading(text)
                elif ptype == ParagraphType.PARTY_BLOCK:
                    self._add_party_block(text)
                elif ptype == ParagraphType.BODY:
                    self._add_body(text)
                elif ptype == ParagraphType.CIZHI:
                    self._add_cizhi(text)
                elif ptype == ParagraphType.COURT_NAME:
                    self._add_court_name(text)
                elif ptype == ParagraphType.SIGNATURE:
                    self._add_signature(text)
                elif ptype == ParagraphType.DATE:
                    self._add_date(text)
                elif ptype == ParagraphType.ATTACHMENT:
                    self._add_attachment(text)
            except Exception as e:
                # 单个段落处理失败不影响整体转换，降级为正文并记录
                self._build_errors.append(f"[{ptype}] {text[:30]}... → {e}")
                try:
                    body = self.params["body"]
                    p = self.doc.add_paragraph()
                    self._set_paragraph_spacing(p, body["line_val"], body["line_rule"])
                    self._set_first_line_indent(p)
                    run = p.add_run(text)
                    self._set_run_font(run, body["font"], body["sz"], False)
                except Exception:
                    pass  # 最终兜底：跳过无法处理的段落

        return self.doc

    def _setup_page(self) -> None:
        """设置页面大小和边距。"""
        page = self.params["page"]
        section = self.doc.sections[0]
        section.page_width = Twips(page["width"])
        section.page_height = Twips(page["height"])
        section.top_margin = Twips(page["margin_top"])
        section.bottom_margin = Twips(page["margin_bottom"])
        section.left_margin = Twips(page["margin_left"])
        section.right_margin = Twips(page["margin_right"])
        section.header_distance = Twips(page["header"])
        section.footer_distance = Twips(page["footer"])

    def _setup_default_style(self) -> None:
        """设置默认样式（正文宋体四号）。"""
        body = self.params["body"]
        style = self.doc.styles["Normal"]
        style.font.name = body["font"]
        style.font.size = Pt(body["size_pt"])
        # 设置东亚字体
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:ascii"), body["font"])
        rfonts.set(qn("w:eastAsia"), body["font"])
        rfonts.set(qn("w:hAnsi"), body["font"])

    def _set_run_font(self, run, font_name: str, sz: int, bold: bool = False) -> None:
        """设置 run 的字体属性（包括东亚字体）。"""
        run.font.name = font_name
        run.font.size = Pt(sz / 2)
        run.font.bold = bold
        # 设置东亚字体
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:ascii"), font_name)
        rfonts.set(qn("w:eastAsia"), font_name)
        rfonts.set(qn("w:hAnsi"), font_name)
        # 设置字号（half-point）
        sz_el = rpr.find(qn("w:sz"))
        if sz_el is None:
            sz_el = OxmlElement("w:sz")
            rpr.append(sz_el)
        sz_el.set(qn("w:val"), str(sz))
        sz_cs = rpr.find(qn("w:szCs"))
        if sz_cs is None:
            sz_cs = OxmlElement("w:szCs")
            rpr.append(sz_cs)
        sz_cs.set(qn("w:val"), str(sz))

    def _set_paragraph_spacing(self, paragraph, line_val: int = 500, line_rule: str = "exact") -> None:
        """设置段落行距（固定值）。"""
        ppr = paragraph._element.get_or_add_pPr()
        spacing = ppr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            ppr.append(spacing)
        spacing.set(qn("w:line"), str(line_val))
        spacing.set(qn("w:lineRule"), line_rule)
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")

    def _set_first_line_indent(self, paragraph, chars: int = 200, twips: int = 560) -> None:
        """设置首行缩进。"""
        ppr = paragraph._element.get_or_add_pPr()
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        ind.set(qn("w:firstLineChars"), str(chars))
        ind.set(qn("w:firstLine"), str(twips))

    def _set_no_indent(self, paragraph) -> None:
        """移除首行缩进（顶格）。"""
        ppr = paragraph._element.get_or_add_pPr()
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            ppr.remove(ind)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")
        ppr.append(ind)

    def _set_alignment(self, paragraph, align: str) -> None:
        """设置段落对齐。"""
        if align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == "justify":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def _add_title(self, text: str) -> None:
        """添加文书标题（宋体二号居中）。"""
        title = self.params["title"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, title["align"])
        self._set_paragraph_spacing(p)
        run = p.add_run(text)
        self._set_run_font(run, title["font"], title["sz"], title["bold"])

    def _add_section_heading(self, text: str) -> None:
        """添加段落标题（宋体四号左对齐+首行缩进）。"""
        heading = self.params["section_heading"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, heading["align"])
        self._set_paragraph_spacing(p)
        self._set_first_line_indent(p)
        run = p.add_run(text)
        self._set_run_font(run, heading["font"], heading["sz"], heading["bold"])

    def _add_party_block(self, text: str) -> None:
        """添加当事人板块（宋体四号左对齐+首行缩进）。"""
        body = self.params["body"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, "left")
        self._set_paragraph_spacing(p)
        self._set_first_line_indent(p)
        run = p.add_run(text)
        self._set_run_font(run, body["font"], body["sz"], False)

    def _add_body(self, text: str) -> None:
        """添加正文段落（宋体四号左对齐+首行缩进+固定行距25磅）。"""
        body = self.params["body"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, "left")
        self._set_paragraph_spacing(p, body["line_val"], body["line_rule"])
        self._set_first_line_indent(p, body["first_line_chars"], body["first_line_twips"])
        run = p.add_run(text)
        self._set_run_font(run, body["font"], body["sz"], False)

    def _add_cizhi(self, text: str) -> None:
        """添加"此致"（宋体四号左对齐+首行缩进2字符）。"""
        body = self.params["body"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, "left")
        self._set_paragraph_spacing(p)
        self._set_first_line_indent(p, body["first_line_chars"], body["first_line_twips"])
        run = p.add_run(text)
        self._set_run_font(run, body["font"], body["sz"], False)

    def _add_court_name(self, text: str) -> None:
        """添加法院名称（宋体四号左对齐顶格无缩进）。"""
        court = self.params["court_name"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, court["align"])
        self._set_paragraph_spacing(p)
        self._set_no_indent(p)
        run = p.add_run(text)
        self._set_run_font(run, court["font"], court["sz"], False)

    def _add_signature(self, text: str) -> None:
        """添加签名行（宋体四号右对齐）。"""
        sig = self.params["signature"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, sig["align"])
        self._set_paragraph_spacing(p)
        self._set_no_indent(p)
        run = p.add_run(text)
        self._set_run_font(run, sig["font"], sig["sz"], False)

    def _add_date(self, text: str) -> None:
        """添加日期行（宋体四号右对齐）。"""
        sig = self.params["signature"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, sig["align"])
        self._set_paragraph_spacing(p)
        self._set_no_indent(p)
        run = p.add_run(text)
        self._set_run_font(run, sig["font"], sig["sz"], False)

    def _add_attachment(self, text: str) -> None:
        """添加附件（宋体四号左对齐+首行缩进）。"""
        body = self.params["body"]
        p = self.doc.add_paragraph()
        self._set_alignment(p, "left")
        self._set_paragraph_spacing(p)
        self._set_first_line_indent(p, body["first_line_chars"], body["first_line_twips"])
        run = p.add_run(text)
        self._set_run_font(run, body["font"], body["sz"], False)


# ===========================================================================
# 第五部分：后验校验（V1-V12）
# ===========================================================================

def verify_docx(doc: Document, params: dict[str, Any]) -> list[dict]:
    """对生成的 Word 文档执行 V1-V12 后验校验。

    返回校验结果列表，每项含 check/name/passed/expected/actual。
    """
    results = []

    results.append(_verify_v1_page_size(doc, params))
    results.append(_verify_v2_margins(doc, params))
    results.append(_verify_v3_title_font(doc, params))
    results.append(_verify_v4_title_align(doc, params))
    results.append(_verify_v5_body_font(doc, params))
    results.append(_verify_v6_body_line_spacing(doc, params))
    results.append(_verify_v7_body_indent(doc, params))
    results.append(_verify_v8_party_separation(doc, params))
    results.append(_verify_v9_section_heading_bold(doc, params))
    results.append(_verify_v10_signature_align(doc, params))
    results.append(_verify_v11_no_markdown(doc, params))
    results.append(_verify_v12_no_straight_quotes(doc, params))

    return results




def _verify_v1_page_size(doc, params):
    page = params["page"]
    section = doc.sections[0]
    w = int(section.page_width.twips) if section.page_width else 0
    h = int(section.page_height.twips) if section.page_height else 0
    expected_w = page["width"]
    expected_h = page["height"]
    passed = (w == expected_w and h == expected_h)
    return _make_result_common("V1", "页面尺寸", passed,
                        f"{expected_w}x{expected_h} twips (A4)",
                        f"{w}x{h} twips")


def _verify_v2_margins(doc, params):
    page = params["page"]
    section = doc.sections[0]
    top = int(section.top_margin.twips) if section.top_margin else 0
    bottom = int(section.bottom_margin.twips) if section.bottom_margin else 0
    left = int(section.left_margin.twips) if section.left_margin else 0
    right = int(section.right_margin.twips) if section.right_margin else 0
    passed = (top == page["margin_top"] and bottom == page["margin_bottom"]
              and left == page["margin_left"] and right == page["margin_right"])
    return _make_result_common("V2", "页边距", passed,
                        f"上={page['margin_top']}, 下={page['margin_bottom']}, 左={page['margin_left']}, 右={page['margin_right']}",
                        f"上={top}, 下={bottom}, 左={left}, 右={right}")


def _verify_v3_title_font(doc, params):
    title = params["title"]
    if not doc.paragraphs:
        return _make_result_common("V3", "标题字体+字号", False, f"{title['font']}, sz={title['sz']}", "文档无段落")
    title_para = doc.paragraphs[0]
    if not title_para.runs:
        return _make_result_common("V3", "标题字体+字号", False, f"{title['font']}, sz={title['sz']}", "标题段落无 runs")
    run = title_para.runs[0]
    font_name, ea_font, sz_val = get_run_font_info(run)
    name_ok = (font_name == title["font"]) or (ea_font == title["font"])
    size_ok = (sz_val == title["sz"])
    actual = f"{ea_font or font_name or '未设置'}, sz={sz_val}"
    passed = name_ok and size_ok
    return _make_result_common("V3", "标题字体+字号", passed, f"{title['font']}, sz={title['sz']}", actual)


def _verify_v4_title_align(doc, params):
    title = params["title"]
    if not doc.paragraphs:
        return _make_result_common("V4", "标题对齐", False, title["align"], "文档无段落")
    title_para = doc.paragraphs[0]
    align = str(title_para.alignment) if title_para.alignment is not None else "None"
    # WD_ALIGN_PARAGRAPH.CENTER = 1
    expected_map = {"center": "CENTER (1)", "left": "LEFT (0)", "right": "RIGHT (2)"}
    expected_val = expected_map.get(title["align"], title["align"])
    passed = False
    if title["align"] == "center":
        passed = title_para.alignment == WD_ALIGN_PARAGRAPH.CENTER
    elif title["align"] == "left":
        passed = title_para.alignment in (WD_ALIGN_PARAGRAPH.LEFT, None)
    elif title["align"] == "right":
        passed = title_para.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    return _make_result_common("V4", "标题对齐", passed, expected_val, align)


def _verify_v5_body_font(doc, params):
    body = params["body"]
    issues = []
    checked = 0
    total = len(doc.paragraphs)
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if is_title_paragraph(para, idx):
            continue
        if is_signature_paragraph(para, idx, total):
            continue
        if not para.runs:
            continue
        run = para.runs[0]
        font_name, ea_font, sz_val = get_run_font_info(run)
        name_ok = (font_name == body["font"]) or (ea_font == body["font"])
        size_ok = (sz_val == body["sz"])
        if not name_ok or not size_ok:
            snippet = text[:20] + ("..." if len(text) > 20 else "")
            issues.append(f"段落[{idx}] \"{snippet}\": {ea_font or font_name}, sz={sz_val}")
        checked += 1
        if checked >= 50:
            break
    if checked == 0:
        return _make_result_common("V5", "正文字体+字号", True, f"{body['font']}, sz={body['sz']}", "未检测到正文段落", "warning")
    passed = len(issues) == 0
    actual = f"已检查 {checked} 段, 均符合" if passed else "; ".join(issues[:3])
    if not passed and len(issues) > 3:
        actual += f" ... 共 {len(issues)} 处不符"
    return _make_result_common("V5", "正文字体+字号", passed, f"{body['font']}, sz={body['sz']}", actual)


def _verify_v6_body_line_spacing(doc, params):
    body = params["body"]
    issues = []
    checked = 0
    total = len(doc.paragraphs)
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if is_title_paragraph(para, idx):
            continue
        if is_signature_paragraph(para, idx, total):
            continue
        ppr = para._element.find(qn("w:pPr"))
        if ppr is None:
            issues.append(f"段落[{idx}]: 无 pPr")
            checked += 1
            continue
        spacing_el = ppr.find(qn("w:spacing"))
        if spacing_el is None:
            issues.append(f"段落[{idx}]: 无 spacing")
            checked += 1
            continue
        line_val = spacing_el.get(qn("w:line"))
        rule_val = spacing_el.get(qn("w:lineRule"))
        line_ok = (line_val == str(body["line_val"]))
        rule_ok = (rule_val == body["line_rule"])
        if not line_ok or not rule_ok:
            issues.append(f"段落[{idx}]: line={line_val}, lineRule={rule_val}")
        checked += 1
        if checked >= 50:
            break
    if checked == 0:
        return _make_result_common("V6", "正文行距", True, f"line={body['line_val']}, lineRule={body['line_rule']}", "未检测到正文段落", "warning")
    passed = len(issues) == 0
    actual = f"已检查 {checked} 段, 均符合" if passed else "; ".join(issues[:3])
    if not passed and len(issues) > 3:
        actual += f" ... 共 {len(issues)} 处不符"
    return _make_result_common("V6", "正文行距", passed, f"line={body['line_val']}, lineRule={body['line_rule']}", actual)


def _verify_v7_body_indent(doc, params):
    body = params["body"]
    issues = []
    checked = 0
    total = len(doc.paragraphs)
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if is_title_paragraph(para, idx):
            continue
        if is_signature_paragraph(para, idx, total):
            continue
        ppr = para._element.find(qn("w:pPr"))
        if ppr is None:
            issues.append(f"段落[{idx}]: 无 pPr")
            checked += 1
            continue
        ind_el = ppr.find(qn("w:ind"))
        if ind_el is None:
            issues.append(f"段落[{idx}]: 无 ind")
            checked += 1
            continue
        first_line = ind_el.get(qn("w:firstLine"))
        first_chars = ind_el.get(qn("w:firstLineChars"))
        chars_ok = first_chars == str(body["first_line_chars"])
        twips_ok = first_line == str(body["first_line_twips"])
        if not chars_ok and not twips_ok:
            issues.append(f"段落[{idx}]: firstLine={first_line}, firstLineChars={first_chars}")
        checked += 1
        if checked >= 50:
            break
    if checked == 0:
        return _make_result_common("V7", "正文首行缩进", True, f"firstLineChars={body['first_line_chars']}, firstLine={body['first_line_twips']}", "未检测到正文段落", "warning")
    passed = len(issues) == 0
    actual = f"已检查 {checked} 段, 均符合" if passed else "; ".join(issues[:3])
    if not passed and len(issues) > 3:
        actual += f" ... 共 {len(issues)} 处不符"
    return _make_result_common("V7", "正文首行缩进", passed, f"firstLineChars={body['first_line_chars']}, firstLine={body['first_line_twips']}", actual)


def _verify_v8_party_separation(doc, params):
    """V8: 当事人板块前有空段落分隔。"""
    issues = []
    total = len(doc.paragraphs)
    party_indices = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        for kw in PARTY_KEYWORDS:
            if text.startswith(kw) and (kw + "：" in text or kw + ":" in text):
                # 排除落款区域
                if not is_signature_paragraph(para, idx, total):
                    party_indices.append(idx)
                break

    for idx in party_indices:
        if idx == 0:
            continue
        prev = doc.paragraphs[idx - 1]
        if prev.text.strip() != "":
            issues.append(f"当事人段落[{idx}]前无空行分隔")

    passed = len(issues) == 0
    actual = f"已检查 {len(party_indices)} 个当事人板块, 均符合" if passed else "; ".join(issues[:3])
    return _make_result_common("V8", "当事人间空段落", passed, "当事人板块前有空段落", actual)


def _verify_v9_section_heading_bold(doc, params):
    """V9: 事实与理由等节标题加粗检查（本规范为不加粗，校验是否误加粗）。"""
    issues = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if not is_title_paragraph(para, idx):
            continue
        if idx == 0:
            continue  # 文书标题不检查
        if not para.runs:
            continue
        run = para.runs[0]
        rpr = run._element.find(qn("w:rPr"))
        if rpr is not None:
            b_el = rpr.find(qn("w:b"))
            if b_el is not None:
                val = b_el.get(qn("w:val"))
                if val is None or val == "true" or val == "1":
                    # 节标题不应加粗（但"事实与理由"等大标题允许加粗）
                    if re.match(r"^[一二三四五六七八九十]+、", text):
                        pass  # 中文编号标题允许加粗
                    else:
                        issues.append(f"段落[{idx}]: \"{text[:20]}\" 不应加粗")

    passed = len(issues) == 0
    actual = "均符合" if passed else "; ".join(issues[:3])
    return _make_result_common("V9", "段落标题加粗", passed, "段落标题不加粗", actual)


def _verify_v10_signature_align(doc, params):
    """V10: 签名区右对齐。"""
    issues = []
    total = len(doc.paragraphs)
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if not is_signature_paragraph(para, idx, total):
            continue
        if text.startswith("此致"):
            continue  # 此致不是右对齐
        # 先看 w:jc，缺失即视为未设置（Word 按默认左对齐渲染），必须报不合格；
        # 不能因 para.alignment is None 而跳过该段落。
        ppr = para._element.find(qn("w:pPr"))
        jc_el = ppr.find(qn("w:jc")) if ppr is not None else None
        if jc_el is not None:
            jc_val = jc_el.get(qn("w:val"))
            if jc_val != "right":
                issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 对齐={jc_val}")
        elif para.alignment is None:
            issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 无对齐设置(默认left)")
        elif para.alignment != WD_ALIGN_PARAGRAPH.RIGHT:
            issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 对齐未设为right")

    passed = len(issues) == 0
    actual = "均符合" if passed else "; ".join(issues[:3])
    return _make_result_common("V10", "签名区右对齐", passed, "jc=right", actual)


def _verify_v11_no_markdown(doc, params):
    """V11: 全文无 Markdown 残留。"""
    md_patterns = [
        r"^#{1,6}\s",
        r"\*\*.*?\*\*",
        r"__.*?__",
        r"<sup>.*?</sup>",
        r"<sub>.*?</sub>",
        r"^>\s",
        r"^[-*_]{3,}$",
    ]
    issues = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        for pattern in md_patterns:
            if re.search(pattern, text, re.MULTILINE):
                issues.append(f"段落[{idx}]: 残留 \"{pattern}\"")
                break

    passed = len(issues) == 0
    actual = "无残留" if passed else "; ".join(issues[:3])
    return _make_result_common("V11", "无 Markdown 残留", passed, "全文无 Markdown 标记", actual)


def _verify_v12_no_straight_quotes(doc, params):
    """V12: 全文无直引号。"""
    issues = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        if '"' in text or "'" in text:
            count = text.count('"') + text.count("'")
            issues.append(f"段落[{idx}]: {count} 个直引号")

    passed = len(issues) == 0
    actual = "无直引号" if passed else "; ".join(issues[:3])
    return _make_result_common("V12", "无直引号", passed, "全文无直引号", actual)


# ===========================================================================
# 第六部分：转换主流程（4级降级链）
# ===========================================================================

def convert(md_path: str, docx_path: str, spec_path: str | None = None) -> dict:
    """执行 Markdown → Word 转换，返回结果报告。

    4级降级链:
    Tier 1: md2docx_legal.py (本脚本, python-docx) ← 首选
    Tier 2: 外部 DOCX MCP (通过环境变量/配置调用)
    Tier 3: pandoc + reference.docx
    Tier 4: 纯 Markdown 兜底交付
    """
    report = {
        "input": md_path,
        "output": docx_path,
        "tier": None,
        "verification": [],
        "errors": [],
    }

    # 读取 Markdown
    try:
        md_text = Path(md_path).read_text(encoding="utf-8")
    except Exception as e:
        report["errors"].append(f"读取 Markdown 失败: {e}")
        return report

    # 加载排版参数
    params = load_spec(spec_path)

    # Tier 1: python-docx 精确转换
    try:
        lines = md_text.split("\n")
        paragraphs = normalize_structure(lines)
        builder = DocxBuilder(params)
        doc = builder.build(paragraphs)

        # 保存
        doc.save(docx_path)
        report["tier"] = "Tier 1 (md2docx_legal.py)"

        # 后验校验
        report["verification"] = verify_docx(doc, params)

        # 检查校验结果
        failed = [v for v in report["verification"] if not v["passed"] and v["severity"] == "error"]
        if failed:
            thresholds = params["content_rules"]["verify_thresholds"]
            if len(failed) <= thresholds["auto_fix_max"]:
                # 自动修复
                report["auto_fixed"] = _auto_fix(doc, failed, params)
                doc.save(docx_path)
                # 重新校验
                report["verification"] = verify_docx(doc, params)
            # 否则保留结果，由调用方决定是否退回

        return report

    except Exception as e:
        report["errors"].append(f"Tier 1 失败: {e}")
        import traceback
        report["errors"].append(traceback.format_exc())

    # Tier 2: 外部 DOCX MCP（暂不实现，降级到 Tier 3）
    report["errors"].append("Tier 2 (外部 MCP) 不可用")

    # Tier 3: pandoc
    try:
        import subprocess
        # pandoc 模板路径：templates/ 与本脚本同级（均位于 scripts/ 下）
        _ref_doc = Path(__file__).resolve().parent / "templates" / "reference.docx"
        _cmd = ["pandoc", md_path, "-o", docx_path]
        # 模板缺失时不传 --reference-doc，否则 pandoc 直接报错退出
        if _ref_doc.is_file():
            _cmd.append(f"--reference-doc={_ref_doc}")
        result = subprocess.run(
            _cmd, capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0:
            report["tier"] = "Tier 3 (pandoc)"
            return report
        else:
            report["errors"].append(f"Tier 3 pandoc 失败: {result.stderr}")
    except FileNotFoundError:
        report["errors"].append("Tier 3: pandoc 未安装")
    except Exception as e:
        report["errors"].append(f"Tier 3 失败: {e}")

    # Tier 4: 纯 Markdown 兜底
    cleaned_md = clean_text(md_text)
    md_out = docx_path.rsplit(".", 1)[0] + ".md"
    Path(md_out).write_text(cleaned_md, encoding="utf-8")
    report["tier"] = "Tier 4 (纯 Markdown 兜底)"
    report["output"] = md_out
    report["errors"].append("Word 转换全部失败，已交付清洗后的 Markdown")

    return report


def _auto_fix(doc: Document, failed_items: list[dict], params: dict) -> list[str]:
    """自动修复可精确定向修正的偏差。"""
    fixed = []
    body = params["body"]
    title = params["title"]
    builder = DocxBuilder(params)  # 创建一次，复用
    # DocxBuilder.__init__ 会新建一个空白 Document，而 _setup_page() 作用于
    # builder.doc。此处指向待修复文档，否则页面修复会写进那个空白文档。
    builder.doc = doc

    for item in failed_items:
        check = item["check"]

        if check == "V3":
            if doc.paragraphs and doc.paragraphs[0].runs:
                run = doc.paragraphs[0].runs[0]
                builder._set_run_font(run, title["font"], title["sz"], title["bold"])
                fixed.append("V3: 标题字体已修复")

        elif check in ("V5",):
            total = len(doc.paragraphs)
            for idx, para in enumerate(doc.paragraphs):
                if not para.text.strip():
                    continue
                if is_title_paragraph(para, idx):
                    continue
                if is_signature_paragraph(para, idx, total):
                    continue
                for run in para.runs:
                    builder._set_run_font(run, body["font"], body["sz"], False)
            fixed.append("V5: 正文字体已修复")

        elif check == "V6":
            total = len(doc.paragraphs)
            for idx, para in enumerate(doc.paragraphs):
                if not para.text.strip():
                    continue
                if is_title_paragraph(para, idx):
                    continue
                if is_signature_paragraph(para, idx, total):
                    continue
                builder._set_paragraph_spacing(para, body["line_val"], body["line_rule"])
            fixed.append("V6: 行距已修复")

        elif check == "V7":
            total = len(doc.paragraphs)
            for idx, para in enumerate(doc.paragraphs):
                if not para.text.strip():
                    continue
                if is_title_paragraph(para, idx):
                    continue
                if is_signature_paragraph(para, idx, total):
                    continue
                builder._set_first_line_indent(para, body["first_line_chars"], body["first_line_twips"])
            fixed.append("V7: 首行缩进已修复")

        elif check in ("V1", "V2"):
            builder._setup_page()
            fixed.append(f"{check}: 页面设置已修复")

    return fixed


# ===========================================================================
# 第七部分：命令行入口
# ===========================================================================

def main() -> int:
    force_utf8_output()
    parser = argparse.ArgumentParser(
        description="法律文书 Markdown → Word 转换器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python3 md2docx_legal.py --spec format-spec.md input.md output.docx
  python3 md2docx_legal.py input.md output.docx
        """
    )
    parser.add_argument("--spec", default=None, help="排版规范文件路径 (format-spec.md)")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("output", help="输出 Word 文件路径")
    parser.add_argument("--verify-only", action="store_true", help="仅校验已有 docx，不转换")
    args = parser.parse_args()

    if args.verify_only:
        # 仅校验模式
        params = load_spec(args.spec)
        try:
            doc = Document(args.input)
        except Exception as e:
            print(f"错误: 无法打开 {args.input}: {e}", file=sys.stderr)
            return 2
        results = verify_docx(doc, params)
        _print_verification_report(results)
        failed = [r for r in results if not r["passed"] and r["severity"] == "error"]
        if not failed:
            return 0
        elif len(failed) <= params["content_rules"]["verify_thresholds"]["auto_fix_max"]:
            return 1  # 警告
        else:
            return 2  # 错误

    # 转换模式
    report = convert(args.input, args.output, args.spec)

    # 输出报告
    print(f"输入: {report['input']}")
    print(f"输出: {report['output']}")
    print(f"转换器: {report['tier']}")

    if report.get("auto_fixed"):
        print("\n自动修复:")
        for fix in report["auto_fixed"]:
            print(f"  ✅ {fix}")

    if report["verification"]:
        _print_verification_report(report["verification"])

    if report["errors"]:
        print("\n错误信息:", file=sys.stderr)
        for err in report["errors"]:
            print(f"  {err}", file=sys.stderr)

    # 判定退出码
    if report["verification"]:
        failed = [r for r in report["verification"] if not r["passed"] and r["severity"] == "error"]
        if not failed:
            print("\n✅ 全部校验通过")
            return 0
        elif len(failed) <= 2:
            print(f"\n⚠️ {len(failed)} 项警告")
            return 1
        else:
            print(f"\n❌ {len(failed)} 项错误")
            return 2

    return 0 if not report["errors"] else 2


def _print_verification_report(results: list[dict]) -> None:
    """打印后验校验报告。"""
    print("\n━━━ Word 后验校验 ━━━")
    passed_count = 0
    failed_count = 0
    for r in results:
        status = "✅" if r["passed"] else ("⚠️" if r["severity"] == "warning" else "❌")
        line = f"{r['check']:4s} {r['name']:16s} {status} 期望: {r['expected']}"
        if not r["passed"]:
            line += f" | 实际: {r['actual']}"
        print(line)
        if r["passed"]:
            passed_count += 1
        elif r["severity"] == "error":
            failed_count += 1
    print("━━━━━━━━━━━━━━━━━━━━")
    total = len(results)
    print(f"结论: {'✅ 全部通过' if failed_count == 0 else '❌ 有失败项'} ({passed_count}/{total} 通过)")


if __name__ == "__main__":
    sys.exit(main())
