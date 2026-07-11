# 中国民商事诉讼工具箱

> cn-litigation-toolkit — 面向执业律师与公司法务的全流程诉讼工具箱

## 概览

本插件为中国民商事诉讼提供从案件分析到文书生成的完整工作流支持，覆盖主诉与被诉双视角。基于请求权基础方法和法庭报告技术，将方法论转化为实务落地工具。

**目标用户**：执业律师、公司法务

**技能总数**：23 项

**适用平台**：任何支持 MCP 协议的 Agent 运行时（Cursor、Claude Code、Gemini CLI、OpenCode 等）

---

## 架构

```
┌─────────────────────────────────────────────────────────────────┐
│                    配置与路由层（3项）                              │
│  诉讼套件中枢 ─── 冷启动访谈 ─── 自定义                            │
└───────────────────────────┬─────────────────────────────────────┘
                            │
        ┌───────────────────┼───────────────────┐
        ▼                   ▼                   ▼
┌───────────────┐  ┌────────────────┐  ┌──────────────────┐
│  基础设施（4项） │  │  文书生成（8项）  │  │  辅助工作（6项）   │
│               │  │                │  │                  │
│ · 文档解析           │  │ · 主诉诉状             │  │ · 诉讼可视化                │
│ · 法律核验           │  │ · 被诉答辩状         │  │ · 案件事实梳理         │
│ · 文书排版           │  │ · 证据目录           │  │ · 要件攻防分析           │
│ · 诉讼知识库         │  │ · 证据装册           │  │ · 财产线索调查           │
│               │  │ · 质证意见           │  │ · 法律研究报告           │
│               │  │ · 庭审提纲           │  │ · 新法解读                 │
│               │  │ · 代理词             │  │                  │
│               │  │ · 程序性文书系列 │  │                  │
└───────────────┘  └───────┬────────┘  └──────────────────┘
                           │ 输出前必过
                           ▼
                  ┌──────────────────────────────────┐
                  │       输出侧闸门 · 质控围栏          │
                  │       · 法律核验（法律检索后端）             │
                  └──────────────────────────────────┘
                                       │
                               ┌───────┴────────┐
                               │  案件管理（2项）  │
                               │ · 案件管家             │
                               │ · 案件复盘             │
                               └────────────────┘
```

---

## 首次使用

安装后首次使用，运行 `/onboarding-interview` 完成实践画像配置。套件会根据你的角色（法务/律师）、立场偏好（主诉/被诉）、风险校准等信息，为后续所有技能提供个性化调整。

如果跳过冷启动，所有技能仍可正常使用，但不会有个性化推荐。

---

## 知识库后端

套件通过 `KB.*` 抽象能力槽调用知识库，支持六种后端，可按三区（A部门法知识库/B办案笔记/C工作笔记）分开绑定：

| 后端 | 类型 | 特点 |
|------|------|------|
| qmind | 云端语义检索 | 亚秒级响应，RAG 问答 |
| 飞书知识库 | 团队协同 | 在线协作 |
| 钉钉知识库 | 团队协同 | 在线协作 |
| Obsidian Vault | 本地 Markdown | 零依赖，Grep 检索 |
| ima 知识库 | 云端语义检索 | 腾讯 ima，RAG 问答 |
| 语雀 | 文档知识库 | 在线协作 |

---

## MCP 依赖（可选增强 · 平台无关 · 多供应商灾备）

本插件的所有 MCP 依赖均为**可选**——不连接任何 MCP 也能产出完整草稿。详见 `connectors.md`。

| 能力后端 | 能力槽 | 缺位降级 |
|---------|---------|---------|
| 法律检索 | `LAW.*` | 标注 `[L4-法条待验证]` |
| 工商信息 | `BIZ.*` | WebSearch 降级 |
| 全网搜索 | `SEARCH.*` | WebFetch 降级 |
| 文档解析 | `PARSE.*` | 内置 pypdf/Read |
| 协作平台 | `COLLAB.*` | 本地 JSON+Markdown |
| 知识库 | `KB.*` | 对话内上下文 |

---

## 文件结构

```
中国民商事诉讼工具箱/
├── plugin.json                    ← 插件元数据
├── Expert.md                      ← 实践画像（用户/AI 协同维护）
├── profile.md                     ← 实践画像模板
├── format-spec.md                 ← 排版参数源（可编辑/可派生）
├── connectors.md                  ← 连接器/MCP 依赖说明
├── ONBOARDING.md                  ← 首次使用指南
├── LICENSE                        ← CC BY-NC-ND 4.0
├── README.md                      ← 本文件
├── matters/                       ← 案件级配置
│   └── matter-template.md         ← 案件配置模板
├── scripts/                       ← 排版与校验脚本
│   ├── md2docx_legal.py           ← Markdown → Word 转换（含 verify）
│   ├── verify_docx.py             ← 独立 12 项校验器
│   ├── extract_format.py          ← 从 format-spec.md 提取参数为 JSON
│   ├── detect_capabilities.py     ← 能力槽探测器
│   └── templates/
│       └── reference.docx         ← pandoc Tier 3 降级模板
└── skills/                        ← 23 个技能
    ├── litigation-hub/SKILL.md
    ├── onboarding-interview/SKILL.md
    ├── customize/SKILL.md
    ├── document-parse/SKILL.md
    ├── legal-verification/SKILL.md
    ├── document-formatting/SKILL.md
    ├── litigation-knowledge-base/SKILL.md
    ├── litigation-visualization/SKILL.md
    ├── case-facts/SKILL.md
    ├── element-analysis/SKILL.md
    ├── asset-discovery/SKILL.md
    ├── legal-research/SKILL.md
    ├── new-law-explained/SKILL.md
    ├── plaintiff-complaint/SKILL.md
    ├── defense-statement/SKILL.md
    ├── evidence-index/SKILL.md
    ├── evidence-binder/SKILL.md
    ├── cross-examination/SKILL.md
    ├── trial-outline/SKILL.md
    ├── agency-opinion/SKILL.md
    ├── procedural-documents/SKILL.md
    ├── case-manager/SKILL.md
    └── case-review/SKILL.md
```

---

## 作者与许可

本项目由游初（You Chu）原创开发，基于 CC BY-NC-ND 4.0 许议发布。详见 [LICENSE](./LICENSE)。

- 微信: Cyruswcx
- 邮箱: 994559732@qq.com
