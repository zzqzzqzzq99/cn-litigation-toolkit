#!/usr/bin/env python3
"""
legal_common.py — 民商事诉讼工具箱共享模块

本模块从 md2docx_legal.py 和 verify_docx.py 中提取公共代码，消除重复和不一致。
包含：
  - 排版默认参数 (DEFAULT_PARAMS)
  - 段落分类关键词 (SIGNATURE_KEYWORDS, PARTY_KEYWORDS, SECTION_TITLES, DATE_PATTERN)
  - format-spec.md 参数加载 (load_spec 及其辅助函数)
  - 段落判定工具 (is_title_paragraph, is_signature_paragraph, is_empty_paragraph)
  - run 字体信息提取 (get_run_font_info)
  - 校验结果构造 (make_result)

本模块自包含，不依赖 md2docx_legal.py 或 verify_docx.py。
依赖: python-docx (pip install python-docx)
"""
from __future__ import annotations

import copy
import re
from pathlib import Path
from typing import Any

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    raise ImportError("需要安装 python-docx 库。请执行: pip install python-docx")


# ===========================================================================
# 排版参数（FormatSpec）
# ===========================================================================

# 默认参数 —— 与 format-spec.md 保持一致
DEFAULT_PARAMS: dict[str, Any] = {
    "page": {
        "width": 11906,       # A4 宽 (twips)
        "height": 16838,      # A4 高 (twips)
        "margin_top": 1440,   # 2.54cm
        "margin_bottom": 1440,
        "margin_left": 1800,  # 3.17cm
        "margin_right": 1800,
        "header": 851,
        "footer": 992,
        "line_pitch": 312,
    },
    "title": {
        "font": "宋体",
        "size_pt": 22,        # 二号
        "sz": 44,             # half-point
        "bold": False,
        "align": "center",
    },
    "body": {
        "font": "宋体",
        "size_pt": 14,        # 四号
        "sz": 28,
        "line_pt": 25,        # 固定行距
        "line_val": 500,      # OOXML line value
        "line_rule": "exact",
        "first_line_chars": 200,   # 2 字符
        "first_line_twips": 560,
        "space_before": 0,
        "space_after": 0,
    },
    "section_heading": {
        "font": "宋体",
        "size_pt": 14,
        "sz": 28,
        "bold": False,
        "align": "left",
    },
    "signature": {
        "font": "宋体",
        "size_pt": 14,
        "sz": 28,
        "align": "right",
    },
    "court_name": {
        "font": "宋体",
        "size_pt": 14,
        "sz": 28,
        "align": "left",       # 顶格无缩进
    },
    "content_rules": {
        "blocked_authors": [],
        "blocked_books": [],
        "citation_tags": [
            "[参考：litigation-knowledge-base]",
            "[参考：联网检索]",
            "[参考：legal-verification]",
        ],
        "verify_thresholds": {
            "auto_fix_max": 2,
            "reject_min": 3,
        },
    },
}


# ===========================================================================
# 段落分类关键词
# ===========================================================================

# 落款关键词白名单（扩展版，覆盖各类诉讼文书落款）
SIGNATURE_KEYWORDS = [
    "具状人", "答辩人", "申请人", "上诉人", "提交人", "落款人",
    "代理人", "辩护人", "质证人", "委托代理人", "委托诉讼代理人",
    "原审原告", "原审被告", "再审申请人", "申诉人",
]

# 当事人关键词
PARTY_KEYWORDS = [
    "原告", "被告", "第三人", "答辩人", "被答辩人",
    "申请人", "被申请人", "上诉人", "被上诉人", "反诉原告", "反诉被告",
]

# 已知段落标题词
SECTION_TITLES = [
    "诉讼请求", "事实与理由", "事实和理由", "事实与依据",
    "案由", "附", "附件", "此致", "请求事项",
]

# 日期行正则（支持完整日期和留白日期，兼容有无空格）
DATE_PATTERN = re.compile(
    r"^\d{4}\s*年\s*\d{0,2}\s*月\s*\d{0,2}\s*日"
)


# ===========================================================================
# format-spec.md 参数加载
# ===========================================================================

def load_spec(spec_path: str | None) -> dict[str, Any]:
    """加载 format-spec.md，解析参数并覆写默认值。

    解析策略：用正则从 Markdown 表格中提取关键参数值。
    解析失败时使用默认值，不报错（鲁棒性优先）。
    """
    params = copy.deepcopy(DEFAULT_PARAMS)

    if spec_path is None:
        # 自动发现：向上查找 format-spec.md
        cwd = Path.cwd()
        for p in [cwd] + list(cwd.parents):
            candidate = p / "format-spec.md"
            if candidate.exists():
                spec_path = str(candidate)
                break
    if spec_path is None or not Path(spec_path).exists():
        return params

    try:
        text = Path(spec_path).read_text(encoding="utf-8")
    except Exception:
        return params

    # 解析页面设置
    _parse_page_params(text, params)
    # 解析标题参数
    _parse_title_params(text, params)
    # 解析正文参数
    _parse_body_params(text, params)
    # 解析内容规则
    _parse_content_rules(text, params)

    return params


def _extract_table_value(text: str, section_start: str, label: str) -> str | None:
    """从 Markdown 表格中提取某行某列的值。

    格式: | label | value |
    """
    idx = text.find(section_start)
    if idx < 0:
        return None
    # 向后搜索 3000 字符
    chunk = text[idx:idx + 3000]
    # 匹配 | label | value | 或 | label | value | ooxml |
    pattern = rf"\|\s*{re.escape(label)}\s*\|\s*([^|]+?)\s*\|"
    m = re.search(pattern, chunk)
    if m:
        return m.group(1).strip()
    return None


def _extract_ooxml_value(text: str, section: str, label: str, prefix: str) -> int | None:
    """从 OOXML 列提取数值。"""
    idx = text.find(section)
    if idx < 0:
        return None
    chunk = text[idx:idx + 3000]
    pattern = rf"\|\s*{re.escape(label)}\s*\|[^|]*\|[^|]*{re.escape(prefix)}\"?(\d+)\"?"
    m = re.search(pattern, chunk)
    if m:
        return int(m.group(1))
    # 备用：直接在值列搜索
    pattern2 = rf"\|\s*{re.escape(label)}\s*\|\s*[^|]*\|\s*{re.escape(prefix)}=?\"?(\d+)\"?"
    m = re.search(pattern2, chunk)
    if m:
        return int(m.group(1))
    return None


def _parse_page_params(text: str, params: dict) -> None:
    """解析页面设置参数。"""
    section = "一、页面设置"

    # 边距：从 OOXML 列提取（更精确）
    val = _extract_ooxml_value(text, section, "上边距", "top=")
    if val:
        params["page"]["margin_top"] = val
    val = _extract_ooxml_value(text, section, "下边距", "bottom=")
    if val:
        params["page"]["margin_bottom"] = val
    val = _extract_ooxml_value(text, section, "左边距", "left=")
    if val:
        params["page"]["margin_left"] = val
    val = _extract_ooxml_value(text, section, "右边距", "right=")
    if val:
        params["page"]["margin_right"] = val

    # 页面尺寸
    val = _extract_table_value(text, section, "纸张")
    if val and "11906" in val:
        params["page"]["width"] = 11906
        params["page"]["height"] = 16838


def _parse_title_params(text: str, params: dict) -> None:
    """解析标题参数。"""
    section = "二、文书标题"
    # 字号
    val = _extract_table_value(text, section, "字号")
    if val:
        sz = _parse_sz(val)
        if sz:
            params["title"]["sz"] = sz
            params["title"]["size_pt"] = sz / 2
    # 加粗
    val = _extract_table_value(text, section, "加粗")
    if val and ("是" in val or "true" in val.lower()):
        params["title"]["bold"] = True
    elif val and ("否" in val or "false" in val.lower()):
        params["title"]["bold"] = False
    # 对齐
    val = _extract_table_value(text, section, "对齐")
    if val and "居中" in val:
        params["title"]["align"] = "center"


def _parse_body_params(text: str, params: dict) -> None:
    """解析正文参数。"""
    section = "三、正文"
    # 字号
    val = _extract_table_value(text, section, "字号")
    if val:
        sz = _parse_sz(val)
        if sz:
            params["body"]["sz"] = sz
            params["body"]["size_pt"] = sz / 2
    # 行距
    val = _extract_table_value(text, section, "行距")
    if val:
        line_val = _parse_line(val)
        if line_val:
            params["body"]["line_val"] = line_val
            params["body"]["line_pt"] = line_val / 20
    # 首行缩进
    val = _extract_table_value(text, section, "首行缩进")
    if val:
        chars = re.search(r"(\d+)\s*字符", val)
        if chars:
            params["body"]["first_line_chars"] = int(chars.group(1)) * 100
        twips = re.search(r"(\d+)\s*twips", val, re.IGNORECASE)
        if twips:
            params["body"]["first_line_twips"] = int(twips.group(1))


def _parse_content_rules(text: str, params: dict) -> None:
    """解析内容规则参数（YAML 代码块）。"""
    # blocked_authors
    m = re.search(r"blocked_authors:\s*\[(.*?)\]", text, re.DOTALL)
    if m and m.group(1).strip():
        items = [a.strip().strip('"').strip("'") for a in m.group(1).split(",") if a.strip()]
        params["content_rules"]["blocked_authors"] = items

    # blocked_books
    m = re.search(r"blocked_books:\s*\[(.*?)\]", text, re.DOTALL)
    if m and m.group(1).strip():
        items = [a.strip().strip('"').strip("'") for a in m.group(1).split(",") if a.strip()]
        params["content_rules"]["blocked_books"] = items

    # verify_thresholds
    m = re.search(r"auto_fix_max:\s*(\d+)", text)
    if m:
        params["content_rules"]["verify_thresholds"]["auto_fix_max"] = int(m.group(1))
    m = re.search(r"reject_min:\s*(\d+)", text)
    if m:
        params["content_rules"]["verify_thresholds"]["reject_min"] = int(m.group(1))


def _parse_sz(text: str) -> int | None:
    """从文本中解析 sz 值（half-point）。"""
    m = re.search(r"sz\s*=\s*(\d+)", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)pt", text)
    if m:
        return int(float(m.group(1)) * 2)
    return None


def _parse_line(text: str) -> int | None:
    """从文本中解析 line 值。"""
    m = re.search(r"line\s*=\s*(\d+)", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        return int(float(m.group(1)) * 20)
    m = re.search(r"(\d+(?:\.\d+)?)\s*pt", text)
    if m:
        return int(float(m.group(1)) * 20)
    return None


# ===========================================================================
# 段落判定工具
# ===========================================================================

def is_title_paragraph(para, idx) -> bool:
    """判断段落是否为标题（首段、带编号的节标题或文书章节标题）。"""
    if idx == 0:
        return True
    text = para.text.strip()
    # 匹配 "一、""二、" 等中文编号标题
    if re.match(r"^[一二三四五六七八九十]+、", text):
        return True
    # 文书章节标题
    section_titles = ["诉讼请求", "事实与理由", "事实与依据", "事实和理由", "案由", "附"]
    if text in section_titles or re.match(r"^(诉讼请求|事实与理由|事实与依据|事实和理由|案由|附)[：:]?\s*$", text):
        return True
    return False


def is_signature_paragraph(para, idx=None, total=None) -> bool:
    """判断段落是否为签名/日期段落。

    两阶段检查：
    1. 如果提供 idx 和 total，只在文档最后 15 个段落中识别签名段落，
       避免将正文中包含日期的段落误识别为签名行。
    2. 排除当事人板块中的"答辩人："等（PARTY_KEYWORDS 中不在 SIGNATURE_KEYWORDS 的词）。
    """
    if idx is not None and total is not None:
        if idx < total - 15:
            return False
    text = para.text.strip()
    if not text:
        return False
    # 签名行通常较短（< 40 字符），正文段落通常较长
    if len(text) > 40:
        return False
    # 签名区特定关键词
    sig_pattern = r"(具状人|签名|盖章|质证人|代理人|委托代理|辩护人|答辩人|申请人|上诉人|提交人|落款人)"
    if re.search(sig_pattern, text):
        # 排除当事人板块中的"答辩人："等（落款区域的"答辩人："后跟名称且下一行是日期）
        for kw in SIGNATURE_KEYWORDS:
            if text.startswith(kw) and (kw + "：" in text or kw + ":" in text):
                return True
        if re.search(sig_pattern, text) and not any(
            text.startswith(p) and (p + "：" in text or p + ":" in text)
            for p in PARTY_KEYWORDS if p not in SIGNATURE_KEYWORDS
        ):
            return True
    # 以日期格式开头的段落
    if DATE_PATTERN.match(text):
        return True
    # 此致
    if text.startswith("此致"):
        return True
    return False


def is_empty_paragraph(para) -> bool:
    """判断段落是否为空。"""
    return para.text.strip() == ""


# ===========================================================================
# run 字体信息提取
# ===========================================================================

def get_run_font_info(run) -> tuple[str | None, str | None, int | None]:
    """获取 run 的字体名、东亚字体和字号。返回 (font_name, ea_font, sz_val)。

    font_name 优先使用 run.font.name，若为空则回退到 ea_font（东亚字体）。
    """
    rpr = run._element.find(qn("w:rPr"))
    ea_font = None
    sz_val = None
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            ea_font = rfonts.get(qn("w:eastAsia"))
        sz_el = rpr.find(qn("w:sz"))
        if sz_el is not None:
            sz_val = int(sz_el.get(qn("w:val")))
    font_name = run.font.name or ea_font
    return (font_name, ea_font, sz_val)


# ===========================================================================
# 校验结果构造
# ===========================================================================

def make_result(check, name, passed, expected, actual, severity="error") -> dict[str, Any]:
    """构造单条校验结果。"""
    return {
        "check": check,
        "name": name,
        "passed": passed,
        "expected": expected,
        "actual": actual,
        "severity": severity,
    }
