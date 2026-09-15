#!/usr/bin/env python3
"""
verify_docx.py — 法律文书 Word 排版校验脚本（独立 12 项校验器）

用法:
    python3 verify_docx.py [--spec format-spec.md] output.docx

输出: JSON 格式的校验报告
退出码: 0=全部通过, 1=有警告, 2=有错误

校验项:
    V1  页面尺寸         A4 (11906×16838 twips)
    V2  页边距           上/下 1440, 左/右 1800 twips
    V3  标题字体+字号     宋体, sz=44 (二号)
    V4  标题对齐         居中 (center)
    V5  正文字体+字号     宋体, sz=28 (四号)
    V6  正文行距         固定值 25 磅 (line=500, lineRule=exact)
    V7  正文首行缩进     560 twips (2 字符)
    V8  当事人间空段落    每个当事人板块前有空段落
    V9  段落标题加粗检查  节标题不应误加粗
    V10 签名区右对齐     jc=right
    V11 无 Markdown 残留 全文无 ###/**/<sup> 等
    V12 无直引号         全文无 " / '
"""
import sys
import json
import re
import argparse
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("错误: 需要安装 python-docx 库。请执行: pip install python-docx", file=sys.stderr)
    sys.exit(2)


# 共享模块
from legal_common import (
    PARTY_KEYWORDS, DATE_PATTERN,
    load_spec, is_title_paragraph, is_signature_paragraph,
    is_empty_paragraph, get_run_font_info, make_result,
    force_utf8_output,
)

# ===========================================================================
# V1-V12 校验函数
# ===========================================================================

# 对齐值 → 中文说明（用于校验报告的"期望"列）
ALIGN_LABELS = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}

# 对齐值 → python-docx 枚举
ALIGN_ENUMS = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


# ---------------------------------------------------------------------------
# V1: 页面尺寸 — 11906×16838 twips (A4)
# ---------------------------------------------------------------------------
def verify_v1(doc, params):
    page = params["page"]
    section = doc.sections[0]
    w = int(section.page_width.twips) if section.page_width else 0
    h = int(section.page_height.twips) if section.page_height else 0
    actual = f"{w}x{h} twips"
    passed = (w == page["width"] and h == page["height"])
    return make_result("V1", "页面尺寸", passed,
                       f"{page['width']}x{page['height']} twips (A4)", actual)


# ---------------------------------------------------------------------------
# V2: 页边距 — 上/下 1440, 左/右 1800 twips
# ---------------------------------------------------------------------------
def verify_v2(doc, params):
    page = params["page"]
    section = doc.sections[0]
    top = int(section.top_margin.twips) if section.top_margin else 0
    bottom = int(section.bottom_margin.twips) if section.bottom_margin else 0
    left = int(section.left_margin.twips) if section.left_margin else 0
    right = int(section.right_margin.twips) if section.right_margin else 0
    expected_parts = [f"上={page['margin_top']}", f"下={page['margin_bottom']}",
                      f"左={page['margin_left']}", f"右={page['margin_right']}"]
    actual_parts = [f"上={top}", f"下={bottom}", f"左={left}", f"右={right}"]
    passed = (top == page["margin_top"] and bottom == page["margin_bottom"]
              and left == page["margin_left"] and right == page["margin_right"])
    return make_result("V2", "页边距", passed,
                       ", ".join(expected_parts) + " twips",
                       ", ".join(actual_parts) + " twips")


# ---------------------------------------------------------------------------
# V3: 标题字体+字号 — 宋体 + sz=44 (二号)
# ---------------------------------------------------------------------------
def verify_v3(doc, params):
    title = params["title"]
    if not doc.paragraphs:
        return make_result("V3", "标题字体+字号", False, f"{title['font']}, sz={title['sz']}", "文档无段落")
    title_para = doc.paragraphs[0]
    if not title_para.runs:
        return make_result("V3", "标题字体+字号", False, f"{title['font']}, sz={title['sz']}", "标题段落无 runs")
    run = title_para.runs[0]
    font_name, ea_font, sz_val = get_run_font_info(run)
    name_ok = (font_name == title["font"]) or (ea_font == title["font"])
    size_ok = (sz_val == title["sz"])
    actual_name = ea_font or font_name or "未设置"
    actual_size = f"sz={sz_val}" if sz_val else "未设置"
    actual = f"{actual_name}, {actual_size}"
    passed = name_ok and size_ok
    return make_result("V3", "标题字体+字号", passed, f"{title['font']}, sz={title['sz']}", actual)


# ---------------------------------------------------------------------------
# V4: 标题对齐 — 居中 (center)
# ---------------------------------------------------------------------------
def verify_v4(doc, params):
    title = params["title"]
    # 期望对齐取自 format-spec，而非硬编码 center
    expected = title.get("align", "center")
    expected_desc = f"{expected} ({ALIGN_LABELS.get(expected, expected)})"
    if not doc.paragraphs:
        return make_result("V4", "标题对齐", False, expected_desc, "文档无段落")
    title_para = doc.paragraphs[0]
    ppr = title_para._element.find(qn("w:pPr"))
    jc_val = None
    if ppr is not None:
        jc_el = ppr.find(qn("w:jc"))
        if jc_el is not None:
            jc_val = jc_el.get(qn("w:val"))
    # 也检查 alignment 属性
    if jc_val is None and title_para.alignment is not None:
        align_map = {WD_ALIGN_PARAGRAPH.CENTER: "center", WD_ALIGN_PARAGRAPH.LEFT: "left",
                     WD_ALIGN_PARAGRAPH.RIGHT: "right", WD_ALIGN_PARAGRAPH.JUSTIFY: "both"}
        jc_val = align_map.get(title_para.alignment, str(title_para.alignment))
    # 未设置 w:jc 时 Word 按左对齐渲染，故期望 left 时视为合格
    passed = (jc_val == expected) if jc_val is not None else (expected == "left")
    actual = jc_val if jc_val else "未设置(默认left)"
    return make_result("V4", "标题对齐", passed, expected_desc, actual)


# ---------------------------------------------------------------------------
# V5: 正文字体+字号 — 宋体 + sz=28 (四号)
# ---------------------------------------------------------------------------
def verify_v5(doc, params):
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
            actual_name = ea_font or font_name or "未设置"
            issues.append(f"段落[{idx}] \"{snippet}\": {actual_name}, sz={sz_val}")
        checked += 1
        if checked >= 50:
            break
    if checked == 0:
        return make_result("V5", "正文字体+字号", True, f"{body['font']}, sz={body['sz']}", "未检测到正文段落", "warning")
    passed = len(issues) == 0
    if passed:
        actual = f"已检查 {checked} 个段落, 均符合"
    else:
        shown = issues[:3]
        actual = "; ".join(shown)
        if len(issues) > 3:
            actual += f" ... 共 {len(issues)} 处不符"
    return make_result("V5", "正文字体+字号", passed, f"{body['font']}, sz={body['sz']}", actual)


# ---------------------------------------------------------------------------
# V6: 正文行距 — 固定值 25 磅 (line=500, lineRule=exact)
# ---------------------------------------------------------------------------
def verify_v6(doc, params):
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
            issues.append(f"段落[{idx}]: 无 spacing 设置")
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
        return make_result("V6", "正文行距", True,
                           f"line={body['line_val']}, lineRule={body['line_rule']}",
                           "未检测到正文段落", "warning")
    passed = len(issues) == 0
    if passed:
        actual = f"已检查 {checked} 个段落, 均符合"
    else:
        shown = issues[:3]
        actual = "; ".join(shown)
        if len(issues) > 3:
            actual += f" ... 共 {len(issues)} 处不符"
    return make_result("V6", "正文行距", passed,
                       f"line={body['line_val']}, lineRule={body['line_rule']}", actual)


# ---------------------------------------------------------------------------
# V7: 正文首行缩进 — 560 twips (2 字符)
# ---------------------------------------------------------------------------
def verify_v7(doc, params):
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
            issues.append(f"段落[{idx}]: 无缩进设置")
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
        return make_result("V7", "正文首行缩进", True,
                           f"firstLineChars={body['first_line_chars']}, firstLine={body['first_line_twips']}",
                           "未检测到正文段落", "warning")
    passed = len(issues) == 0
    if passed:
        actual = f"已检查 {checked} 个段落, 均符合"
    else:
        shown = issues[:3]
        actual = "; ".join(shown)
        if len(issues) > 3:
            actual += f" ... 共 {len(issues)} 处不符"
    return make_result("V7", "正文首行缩进", passed,
                       f"firstLineChars={body['first_line_chars']}, firstLine={body['first_line_twips']}", actual)


# ---------------------------------------------------------------------------
# V8: 当事人间空段落 — 每个当事人板块前有空段落
# ---------------------------------------------------------------------------
def verify_v8(doc, params):
    issues = []
    total = len(doc.paragraphs)
    party_indices = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        for kw in PARTY_KEYWORDS:
            if text.startswith(kw) and (kw + "：" in text or kw + ":" in text):
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
    return make_result("V8", "当事人间空段落", passed, "当事人板块前有空段落", actual)


# ---------------------------------------------------------------------------
# V9: 段落标题加粗检查 — 节标题不应误加粗（中文编号标题除外）
# ---------------------------------------------------------------------------
def verify_v9(doc, params):
    issues = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if not is_title_paragraph(para, idx):
            continue
        if idx == 0:
            continue  # 文书标题单独由 V3 检查
        if not para.runs:
            continue
        run = para.runs[0]
        rpr = run._element.find(qn("w:rPr"))
        if rpr is not None:
            b_el = rpr.find(qn("w:b"))
            if b_el is not None:
                val = b_el.get(qn("w:val"))
                if val is None or val == "true" or val == "1":
                    # 中文编号标题（一、二、三...）允许加粗
                    if re.match(r"^[一二三四五六七八九十]+、", text):
                        pass
                    else:
                        issues.append(f"段落[{idx}]: \"{text[:20]}\" 不应加粗")

    passed = len(issues) == 0
    actual = "均符合" if passed else "; ".join(issues[:3])
    return make_result("V9", "段落标题加粗", passed, "段落标题不加粗", actual)


# ---------------------------------------------------------------------------
# V10: 签名区右对齐 — jc=right
# ---------------------------------------------------------------------------
def verify_v10(doc, params):
    issues = []
    # 期望对齐取自 format-spec，而非硬编码 right
    expected = params.get("signature", {}).get("align", "right")
    expected_enum = ALIGN_ENUMS.get(expected)
    total = len(doc.paragraphs)
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if not is_signature_paragraph(para, idx, total):
            continue
        if text.startswith("此致"):
            continue  # 此致是首行缩进，不是右对齐
        ppr = para._element.find(qn("w:pPr"))
        if ppr is None:
            issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 无 pPr")
            continue
        jc_el = ppr.find(qn("w:jc"))
        if jc_el is None:
            # 检查 alignment 属性
            if para.alignment is not None and para.alignment != expected_enum:
                issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 对齐未设为{expected}")
            elif para.alignment is None and expected != "left":
                issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 无对齐设置(默认left)")
        else:
            jc_val = jc_el.get(qn("w:val"))
            if jc_val != expected:
                issues.append(f"签名段落[{idx}]: \"{text[:20]}\" 对齐={jc_val}")

    passed = len(issues) == 0
    actual = "均符合" if passed else "; ".join(issues[:3])
    return make_result("V10", "签名区右对齐", passed, f"jc={expected}", actual)


# ---------------------------------------------------------------------------
# V11: 无 Markdown 残留 — 全文无 ###/**/<sup> 等
# ---------------------------------------------------------------------------
def verify_v11(doc, params):
    md_patterns = [
        (r"^#{1,6}\s", "标题标记"),
        (r"\*\*.*?\*\*", "加粗标记"),
        (r"__.*?__", "加粗标记"),
        (r"<sup>.*?</sup>", "上标标签"),
        (r"<sub>.*?</sub>", "下标标签"),
        (r"<br\s*/?>", "br标签"),
        (r"^>\s", "引用标记"),
        (r"^[-*_]{3,}$", "分割线"),
    ]
    issues = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        for pattern, label in md_patterns:
            if re.search(pattern, text, re.MULTILINE | re.IGNORECASE):
                snippet = text[:30].replace("\n", " ")
                issues.append(f"段落[{idx}]: 残留{label} \"{snippet}\"")
                break

    passed = len(issues) == 0
    actual = "无残留" if passed else "; ".join(issues[:3])
    return make_result("V11", "无 Markdown 残留", passed, "全文无 Markdown 标记", actual)


# ---------------------------------------------------------------------------
# V12: 无直引号 — 全文无 " / '
# ---------------------------------------------------------------------------
def verify_v12(doc, params):
    issues = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        straight_double = text.count('"')
        straight_single = text.count("'")
        if straight_double > 0 or straight_single > 0:
            total = straight_double + straight_single
            issues.append(f"段落[{idx}]: {total} 个直引号 (\"={straight_double}, '={straight_single})")

    passed = len(issues) == 0
    actual = "无直引号" if passed else "; ".join(issues[:3])
    return make_result("V12", "无直引号", passed, "全文无直引号", actual)


# ===========================================================================
# 主校验函数
# ===========================================================================

def verify_docx(doc, params):
    """执行全部 12 项校验，返回结果列表。"""
    results = [
        verify_v1(doc, params),
        verify_v2(doc, params),
        verify_v3(doc, params),
        verify_v4(doc, params),
        verify_v5(doc, params),
        verify_v6(doc, params),
        verify_v7(doc, params),
        verify_v8(doc, params),
        verify_v9(doc, params),
        verify_v10(doc, params),
        verify_v11(doc, params),
        verify_v12(doc, params),
    ]
    return results


# ===========================================================================
# 命令行入口
# ===========================================================================

def main():
    force_utf8_output()
    parser = argparse.ArgumentParser(
        description="法律文书 Word 排版校验脚本（独立 12 项校验器）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python3 verify_docx.py output.docx
  python3 verify_docx.py --spec format-spec.md output.docx
  python3 verify_docx.py --spec format-spec.shanghai-2nd.md --json output.docx

退出码:
  0 = 全部通过
  1 = 有警告（非阻塞性问题）
  2 = 有错误（阻塞性问题）
        """
    )
    parser.add_argument("--spec", default=None, help="排版规范文件路径 (format-spec.md)")
    parser.add_argument("--json", action="store_true", help="输出 JSON 格式（默认人类可读格式）")
    parser.add_argument("docx", help="待校验的 Word 文件路径")
    args = parser.parse_args()

    # 加载参数
    params = load_spec(args.spec)

    # 打开文档
    try:
        doc = Document(args.docx)
    except Exception as e:
        error_msg = f"错误: 无法打开 {args.docx}: {e}"
        if args.json:
            print(json.dumps({"error": error_msg, "results": []}, ensure_ascii=False, indent=2))
        else:
            print(error_msg, file=sys.stderr)
        return 2

    # 执行校验
    results = verify_docx(doc, params)

    # 统计
    passed_count = sum(1 for r in results if r["passed"])
    warning_count = sum(1 for r in results if not r["passed"] and r["severity"] == "warning")
    error_count = sum(1 for r in results if not r["passed"] and r["severity"] == "error")

    # 输出
    if args.json:
        output = {
            "file": args.docx,
            "spec": args.spec or "auto-detected",
            "summary": {
                "total": len(results),
                "passed": passed_count,
                "warnings": warning_count,
                "errors": error_count,
            },
            "results": results,
        }
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        print(f"\n━━━ Word 排版校验报告 ━━━")
        print(f"文件: {args.docx}")
        print(f"规范: {args.spec or '自动发现 format-spec.md'}")
        print(f"{'─' * 50}")
        for r in results:
            if r["passed"]:
                status = "✅"
            elif r["severity"] == "warning":
                status = "⚠️"
            else:
                status = "❌"
            line = f"  {r['check']:4s} {r['name']:16s} {status}"
            if not r["passed"]:
                line += f"\n       期望: {r['expected']}"
                line += f"\n       实际: {r['actual']}"
            print(line)
        print(f"{'─' * 50}")
        print(f"  结论: {'✅ 全部通过' if error_count == 0 else '❌ 有失败项'} "
              f"({passed_count}/{len(results)} 通过, {warning_count} 警告, {error_count} 错误)")
        print(f"{'━' * 50}\n")

    # 退出码
    if error_count == 0 and warning_count == 0:
        return 0
    elif error_count == 0:
        return 1  # 仅有警告
    else:
        thresholds = params["content_rules"]["verify_thresholds"]
        if error_count <= thresholds["auto_fix_max"]:
            return 1  # 可自动修复的偏差
        else:
            return 2  # 需退回


if __name__ == "__main__":
    sys.exit(main())
