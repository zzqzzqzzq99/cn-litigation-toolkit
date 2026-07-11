#!/usr/bin/env python3
"""
detect_capabilities.py — 探测当前环境已连接的能力槽（LAW/BIZ/PARSE/DOCX/SEARCH/KB/COLLAB）

用法：
    python3 detect_capabilities.py [--profile profile.md] [--out capabilities.json]

工作原理（三步）：
    1. 读 profile.md 中「外部能力后端」章节，抽取用户显式声明的 provider 映射；
    2. 扫描环境（PATH 中的 CLI、常见 MCP socket、环境变量），检测哪些 provider 实际可用；
    3. 按能力槽汇总为 JSON：{ slot: { declared: [...], available: [...], missing: bool, degradation: '...' } }

本脚本不依赖任何第三方库，只用标准库。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Any

SLOTS = ["LAW", "BIZ", "PARSE", "DOCX", "SEARCH", "KB", "COLLAB"]

# 已知 provider → 探测特征（CLI 名 / 环境变量 / 文件路径）
KNOWN_PROVIDERS: dict[str, dict[str, Any]] = {
    # LAW.*
    "yuandian": {"slot": "LAW", "check": {"env": "YUANDIAN_TOKEN"}},
    "pkulaw": {"slot": "LAW", "check": {"env": "PKULAW_TOKEN"}},
    # BIZ.*
    "qcc": {"slot": "BIZ", "check": {"env": "QCC_TOKEN"}},
    "tianyancha": {"slot": "BIZ", "check": {"env": "TIANYANCHA_TOKEN"}},
    "qixinbao": {"slot": "BIZ", "check": {"env": "QIXINBAO_TOKEN"}},
    # PARSE.*
    "mineru": {"slot": "PARSE", "check": {"cli": "mineru"}},
    "bailian-parser": {"slot": "PARSE", "check": {"env": "BAILIAN_API_KEY"}},
    "markdownify": {"slot": "PARSE", "check": {"cli": "markdownify"}},
    "pypdf": {"slot": "PARSE", "check": {"python": "pypdf"}, "bundled": True},
    # DOCX.*
    "md2docx_legal": {"slot": "DOCX", "check": {"file": "scripts/md2docx_legal.py"}, "bundled": True},
    "pandoc": {"slot": "DOCX", "check": {"cli": "pandoc"}},
    # SEARCH.*
    "tavily": {"slot": "SEARCH", "check": {"env": "TAVILY_API_KEY"}},
    "google-web-search": {"slot": "SEARCH", "check": {"env": "GOOGLE_CSE_KEY"}},
    "iflow-web-search": {"slot": "SEARCH", "check": {"env": "IFLOW_TOKEN"}},
    # KB.*
    "qmind": {"slot": "KB", "check": {"cli": "qmind-darwin-arm64"}},
    "obsidian-vault": {"slot": "KB", "check": {"env": "OBSIDIAN_VAULT"}},
    "ima": {"slot": "KB", "check": {"env": "IMA_MCP_TOKEN"}},
    "feishu-kb": {"slot": "KB", "check": {"cli": "lark"}},
    "dingtalk-kb": {"slot": "KB", "check": {"cli": "dws"}},
    "yuque": {"slot": "KB", "check": {"env": "YUQUE_TOKEN"}},
    # COLLAB.*
    "dws": {"slot": "COLLAB", "check": {"cli": "dws"}},
    "lark-cli": {"slot": "COLLAB", "check": {"cli": "lark"}},
}

# 中文别名 → 英文 provider 名（profile.md 中可能使用中文名称）
CHINESE_ALIASES: dict[str, str] = {
    "北大法宝": "pkulaw",
    "元典": "yuandian",
    "企查查": "qcc",
    "天眼查": "tianyancha",
    "启信宝": "qixinbao",
    "ima": "ima",
    "飞书": "feishu-kb",
    "飞书知识库": "feishu-kb",
    "钉钉": "dingtalk-kb",
    "钉钉知识库": "dingtalk-kb",
    "语雀": "yuque",
}

DEGRADATION_HINTS = {
    "LAW": "无法自动核验法条 → 标注 [L4-法条待验证]，人工复核",
    "BIZ": "SEARCH.web_search 降级到公开信息检索",
    "PARSE": "内置 pypdf → fitz 视觉 → Read；全部失败停下报告",
    "DOCX": "四阶降级链：md2docx_legal → 外部 DOCX MCP → pandoc+ref → 仅 Markdown",
    "SEARCH": "WebFetch 直接抓取已知 URL",
    "KB": "仅依赖对话内上下文，套件仍可运行",
    "COLLAB": "本地 JSON + Markdown 零依赖运行",
}


def check_provider(spec: dict[str, Any]) -> bool:
    """检查单个 provider 是否可用。"""
    ch = spec.get("check", {})
    if "cli" in ch:
        return shutil.which(ch["cli"]) is not None
    if "env" in ch:
        return bool(os.environ.get(ch["env"]))
    if "python" in ch:
        try:
            __import__(ch["python"])
            return True
        except ImportError:
            return False
    if "file" in ch:
        # 相对于当前工作目录或套件根目录查找
        cwd = Path.cwd()
        candidates = [
            cwd / ch["file"],
            cwd.parent / ch["file"],
            cwd / "中国民商事诉讼工具箱" / ch["file"],
        ]
        return any(c.exists() for c in candidates)
    return False


def parse_profile(profile_path: Path) -> dict[str, list[str]]:
    """从 profile.md 抽取用户在「外部能力后端」章节声明的 provider（每槽多家）。"""
    declared: dict[str, list[str]] = {s: [] for s in SLOTS}
    if not profile_path.exists():
        return declared

    text = profile_path.read_text(encoding="utf-8")

    # 逐行扫描：找形如 "LAW.*: xxx" 的行，提取该行中提到的 provider
    for line in text.split("\n"):
        stripped = line.strip()
        for slot in SLOTS:
            # 匹配 "SLOT.*:" 开头的行
            if re.match(rf"{slot}\.\*", stripped):
                # 在该行中搜索已知 provider 名称
                for prov in KNOWN_PROVIDERS:
                    if re.search(re.escape(prov), stripped, flags=re.IGNORECASE):
                        if prov not in declared[slot]:
                            declared[slot].append(prov)
                # 也检查中文别名
                for cn_name, prov_name in CHINESE_ALIASES.items():
                    if cn_name in stripped:
                        prov_spec = KNOWN_PROVIDERS.get(prov_name)
                        if prov_spec and prov_spec["slot"] == slot and prov_name not in declared[slot]:
                            declared[slot].append(prov_name)
                break  # 一行只属于一个 slot

    return declared


def main() -> int:
    ap = argparse.ArgumentParser(description="探测当前环境已连接的能力槽")
    ap.add_argument("--profile", default="profile.md", help="Path to profile.md")
    ap.add_argument("--out", default=None, help="Write JSON to file instead of stdout")
    args = ap.parse_args()

    declared = parse_profile(Path(args.profile))
    result: dict[str, Any] = {}

    for slot in SLOTS:
        avail: list[str] = []
        for prov, spec in KNOWN_PROVIDERS.items():
            if spec["slot"] != slot:
                continue
            if check_provider(spec):
                avail.append(prov)

        result[slot] = {
            "declared": declared.get(slot, []),
            "available": avail,
            "missing": len(avail) == 0,
            "degradation": DEGRADATION_HINTS.get(slot, ""),
        }

    payload = json.dumps(result, ensure_ascii=False, indent=2)

    if args.out:
        Path(args.out).write_text(payload, encoding="utf-8")
        print(f"Wrote {args.out}", file=sys.stderr)
    else:
        print(payload)

    missing = [s for s, v in result.items() if v["missing"]]
    if missing:
        print(
            f"[detect_capabilities] {len(missing)} slot(s) with no available provider: {', '.join(missing)}",
            file=sys.stderr,
        )
        print("  套件仍可运行（降级模式），但相关能力标注 [Ln-待验证]。", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
