#!/usr/bin/env python3
"""
extract_format.py — 从 format-spec.md 提取排版参数为 JSON

用法：
    python3 extract_format.py [--spec /path/to/format-spec.md] [--out params.json]

输出：包含页面、字体、字号、行距、缩进等关键参数的 JSON，供 md2docx_legal.py 或
其它渲染脚本消费。若某参数在 format-spec.md 中缺失，则使用「中国大陆法院通用」默认值。

设计原则：
    - 只读，不修改 format-spec.md
    - 用正则从表格与元信息块中提取
    - 缺项自动落默认值，不报错，附 warnings 列表
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

DEFAULTS: dict[str, Any] = {
    "spec_name": "中国大陆法院诉讼文书通用格式",
    "spec_version": "1.0.0",
    "page": {
        "paper": "A4",
        "margin_top_cm": 2.54,
        "margin_bottom_cm": 2.54,
        "margin_left_cm": 3.17,
        "margin_right_cm": 3.17,
    },
    "title": {"font": "宋体", "size_pt": 22, "bold": False, "align": "center"},
    "body": {
        "font": "宋体",
        "size_pt": 14,
        "line_pt": 25,
        "first_line_chars": 2,
        "space_before": 0,
        "space_after": 0,
    },
    "section_heading": {"font": "宋体", "size_pt": 14, "bold": False, "align": "left+indent2"},
    "signature": {
        "date_format": "YYYY年M月D日",
        "signature_align": "right",
        "court_name_align": "left",
    },
}


def parse_pt(text: str) -> float | None:
    """从文本中解析 pt 值。"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*pt", text)
    if m:
        return float(m.group(1))
    m = re.search(r"(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        return float(m.group(1))
    return None


def parse_cm(text: str) -> float | None:
    """从文本中解析 cm 值。"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*cm", text)
    return float(m.group(1)) if m else None


def parse_chars(text: str) -> int | None:
    """从文本中解析字符数。"""
    m = re.search(r"(\d+)\s*字符", text)
    return int(m.group(1)) if m else None


def slice_between(text: str, start: str, end: str) -> str:
    """截取两个标记之间的文本。"""
    i = text.find(start)
    if i < 0:
        return ""
    j = text.find(end, i + 1)
    if j < 0:
        return text[i:]
    return text[i:j]


def find_pt_in_row(block: str, label: str) -> float | None:
    """从表格行中提取 pt 值。"""
    m = re.search(rf"\|\s*{re.escape(label)}\s*\|\s*([^|]+)\|", block)
    if not m:
        return None
    return parse_pt(m.group(1))


def find_chars_in_row(block: str, label: str) -> int | None:
    """从表格行中提取字符数。"""
    m = re.search(rf"\|\s*{re.escape(label)}\s*\|\s*([^|]+)\|", block)
    if not m:
        return None
    return parse_chars(m.group(1))


def find_sz_in_row(block: str, label: str) -> int | None:
    """从表格行中提取 OOXML sz 值。"""
    m = re.search(rf"\|\s*{re.escape(label)}\s*\|\s*([^|]+)\|", block)
    if not m:
        return None
    val = m.group(1)
    m2 = re.search(r"sz\s*=\s*(\d+)", val)
    if m2:
        return int(m2.group(1))
    return None


def extract(spec_path: Path) -> dict[str, Any]:
    """从 format-spec.md 提取排版参数。"""
    warnings: list[str] = []

    if not spec_path.exists():
        warnings.append(f"format-spec.md not found at {spec_path}; using defaults")
        result = json.loads(json.dumps(DEFAULTS))
        result["_warnings"] = warnings
        result["source"] = str(spec_path)
        return result

    text = spec_path.read_text(encoding="utf-8")
    result: dict[str, Any] = json.loads(json.dumps(DEFAULTS))
    result["source"] = str(spec_path)

    # 元信息
    m = re.search(r"规范名[^\n|]*[|：]\s*([^\n|]+)", text)
    if m:
        result["spec_name"] = m.group(1).strip()
    m = re.search(r"规范版本[^\n|]*[|：]\s*([^\n|]+)", text)
    if m:
        result["spec_version"] = m.group(1).strip()

    # 页面设置
    for key, label in (
        ("margin_top_cm", "上边距"),
        ("margin_bottom_cm", "下边距"),
        ("margin_left_cm", "左边距"),
        ("margin_right_cm", "右边距"),
    ):
        m = re.search(rf"\|\s*{label}\s*\|\s*([^|]+)\|", text)
        if m:
            cm = parse_cm(m.group(1))
            if cm is not None:
                result["page"][key] = cm
            else:
                warnings.append(f"{label} 值无法解析：{m.group(1).strip()}")

    # 标题字号
    title_block = slice_between(text, "## 二、文书标题", "## 三、正文")
    if title_block:
        sz = find_sz_in_row(title_block, "字号")
        if sz:
            result["title"]["size_pt"] = sz / 2
        # 加粗
        m = re.search(r"\|\s*加粗\s*\|\s*([^|]+)\|", title_block)
        if m:
            val = m.group(1).strip()
            result["title"]["bold"] = ("是" in val or "true" in val.lower())

    # 正文
    body_block = slice_between(text, "## 三、正文", "## 四、")
    if body_block:
        sz = find_sz_in_row(body_block, "字号")
        if sz:
            result["body"]["size_pt"] = sz / 2
        line_pt = find_pt_in_row(body_block, "行距")
        if line_pt:
            result["body"]["line_pt"] = line_pt
        chars = find_chars_in_row(body_block, "首行缩进")
        if chars is not None:
            result["body"]["first_line_chars"] = chars

    # 段落标题
    heading_block = slice_between(text, "## 四、段落标题", "## 五、")
    if heading_block:
        sz = find_sz_in_row(heading_block, "字号")
        if sz:
            result["section_heading"]["size_pt"] = sz / 2

    # 生成 OOXML 兼容值（供 md2docx_legal.py / legal_common.py 消费）
    page = result["page"]
    result["ooxml"] = {
        "page": {
            "width_twips": 11906,   # A4 default
            "height_twips": 16838,
            "margin_top_twips": int(round(page.get("margin_top_cm", 2.54) * 567)),
            "margin_bottom_twips": int(round(page.get("margin_bottom_cm", 2.54) * 567)),
            "margin_left_twips": int(round(page.get("margin_left_cm", 3.17) * 567)),
            "margin_right_twips": int(round(page.get("margin_right_cm", 3.17) * 567)),
        },
        "title": {
            "sz": int(result["title"]["size_pt"] * 2),
        },
        "body": {
            "sz": int(result["body"]["size_pt"] * 2),
            "line_val": int(result["body"]["line_pt"] * 20),
            "first_line_twips": int(result["body"]["first_line_chars"] * 280),
            "first_line_chars": result["body"]["first_line_chars"] * 100,
        },
        "section_heading": {
            "sz": int(result["section_heading"]["size_pt"] * 2),
        },
    }

    result["_warnings"] = warnings
    return result


def main() -> int:
    p = argparse.ArgumentParser(description="从 format-spec.md 提取排版参数为 JSON")
    p.add_argument("--spec", default="format-spec.md", help="Path to format-spec.md")
    p.add_argument("--out", default=None, help="Write JSON to file instead of stdout")
    args = p.parse_args()

    data = extract(Path(args.spec))
    payload = json.dumps(data, ensure_ascii=False, indent=2)

    if args.out:
        Path(args.out).write_text(payload, encoding="utf-8")
        print(f"Wrote {args.out}", file=sys.stderr)
    else:
        print(payload)

    warnings = data.get("_warnings") or []
    if warnings:
        print(f"[extract_format] {len(warnings)} warning(s):", file=sys.stderr)
        for w in warnings:
            print(f"  - {w}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
