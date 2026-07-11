# 首次使用指南（ONBOARDING）

> 5 分钟走完首次配置，10 分钟跑完第一个案件流程。

---

## 步骤一 · 安装

克隆或下载本仓库到本地。

根据你使用的 Agent 平台，将本目录（或其中的 `skills/` 子目录）挂载到对应位置：

- **QoderWork**：通过"连接器 > 专家套件"导入本目录
- **Claude Code**：将 `skills/` 内子目录逐一软链或复制到 `~/.claude/skills/`
- **Cursor / Gemini CLI / OpenCode / 其他 Agent 平台**：参照各平台的技能安装方式，将 `skills/` 子目录挂载到对应路径

确认平台已识别到套件（在 Agent 中输入 `/` 应能看到 `/litigation-hub`（诉讼套件中枢）、`/onboarding-interview`（冷启动访谈）等技能）。

### 通用要求

- Python 3.9+（用于 `scripts/md2docx_legal.py` 与 `scripts/verify_docx.py`）
- `pip install python-docx`（Word 转换与校验）
- 可选：pandoc（Tier 3 降级）

---

## 步骤二 · 冷启动访谈

在 Agent 会话中运行：

```
/onboarding-interview
```

访谈将在 5 分钟内完成以下配置并写入项目根 `profile.md`：

| 阶段 | 收集内容 |
|------|----------|
| Phase 1 · 身份与角色 | 律师 / 法务，主诉 / 被诉 / 混合 |
| Phase 2 · 风格与偏好 | 风险偏好、文书语调 |
| Phase 3 · 案件管理 | 台账载体（钉钉 / 飞书 / 本地文件）+ 期限提醒配置 |
| Phase 3.5 · 排版规范 | 默认模板 / 自定义 format-spec.md / 暂不配置 |
| Phase 4 · 知识库 | qmind / 协作平台文档 / 两者并用 / 暂不配置 |
| Phase 5 · 后端探测 | 自动探测已连接的 MCP，汇总能力槽映射 |

访谈可随时"跳过"某一项，后续任何时候通过 `/customize`（自定义）单项修改。

---

## 步骤三 · 检查能力槽

访谈完成后，运行：

```
/litigation-hub
```

查看当前能力槽映射与可用性诊断：

| 能力槽 | 用途 | 推荐后端 |
|--------|------|----------|
| `LAW.*` | 法律检索 | 元典 / 北大法宝 |
| `BIZ.*` | 工商信息 | 企查查 / 天眼查 / 启信宝 |
| `PARSE.*` | 文档解析 | MinerU / 云厂商 / 内置 pypdf 降级 |
| `DOCX.*` | Word 生成 | md2docx_legal.py → pandoc → 纯 md 降级 |
| `SEARCH.*` | 全网搜索 | GoogleWebSearch / Tavily / WebFetch 降级 |
| `KB.*` | 知识库 | qmind / 飞书 / 钉钉 / 语雀 / Obsidian+Grep |
| `COLLAB.*` | 协作平台 | 钉钉 / 飞书 / 本地文件降级 |

无 MCP 连接时：套件仍可运行，法条/主体信息将标注 `[L4-待验证]`。

---

## 步骤四 · 第一个案件

假设你现在收到一份新案件的应诉通知。

### 4.1 建档

```
/case-manager 收案
```

按提示提供：案件简称、当事人、案由、法院、案号、开庭日期等基本信息。案件管家会：
- 写入个人台账（若已配置 `COLLAB.sheet_*`）
- 计算首批期限（应诉举证 15 日 / 举证期 30 日等）
- 创建关键节点日程（若已配置 `COLLAB.calendar_*`）

### 4.2 事实梳理

```
/case-facts
```

上传起诉状、合同、往来函件、聊天记录等。技能会产出：
- `事实梳理主文档.md`
- `法律关系分析图.html`（内嵌 SVG）
- `法律事实时间轴图.html`（内嵌 SVG）

### 4.3 要件攻防

```
/element-analysis
```

依赖 4.2 的事实梳理产出，按请求权基础方法逐要件审查。产出 `要件攻防报告.md`，含核心争点、证据缺口、可行策略。

### 4.4 文书起草

按需选择：

```
/defense-statement       # 一审 / 上诉 / 再审 / 仲裁答辩
/plaintiff-complaint         # 起诉 / 上诉 / 再审 / 仲裁申请
/evidence-index         # 逐项证据说明
/cross-examination         # 逐项三性质证
/trial-outline         # 应对预案 + 询问清单
/agency-opinion           # 庭前 / 庭后代理意见
/procedural-documents   # 保全 / 管辖异议 / 变更诉请等 40 余种
```

所有文书交付前必过 `/legal-verification`（输入侧闸门，逐条核验法条与案例）与 `/document-formatting`（输出侧闸门，格式合规校验），放行后交付 Markdown，用户确认后再转 Word。

### 4.5 台账自动回写

上述每一步产出案件级交付物时，案件管家会自动追加进展到个人台账，无需手动。

### 4.6 结案与复盘

```
/case-review        # 三级覆盖：庭审 / 裁判文书 / 结案复盘
/case-manager 结案
```

复盘报告推荐沉淀到 `/litigation-knowledge-base`（诉讼知识库）B 区（办案笔记），后续同类案件可通过语义检索复用经验。

---

## 步骤五 · 定期维护

- 修改单项配置：`/customize`，不需重跑冷启动
- 修改排版参数：直接编辑项目根 `format-spec.md`
- 新增 MCP 后端：在 Agent 平台连接器中安装后，重跑 `/litigation-hub` 探测

---

## 常见问题

**Q1：某些 MCP 我不想付费/没有账号怎么办？**
A：所有 MCP 都是可选。无 MCP 时套件走内置降级（WebFetch / pypdf / 本地文件 / 纯 Markdown），产出仍完整，仅法条案例标注 `[L4-待验证]` 需人工复核。

**Q2：我的团队用飞书不用钉钉。**
A：`COLLAB.*` 平台无关。在 `profile.md` 中把 `COLLAB.*` 指向 `lark-cli` 等飞书 MCP 即可，套件调用不变。

**Q3：我不想用 qmind，能用本地 Markdown 库吗？**
A：可以。`KB.*` 支持 Obsidian Vault + Grep 降级。冷启动访谈选"暂不配置"，然后手动在 `profile.md` 添加 `kb_backend: obsidian_grep`。

**Q4：可以商用吗？**
A：默认许可（CC BY-NC-ND 4.0）禁止商业用途。商用请联系版权人（微信: Cyruswcx）协商专业版授权。

---

## 下一步

- 阅读 `README.md` 了解总体架构
- 阅读 `connectors.md` 了解每个能力槽的推荐 MCP 供应商
- 从 `matters/matter-template.md` 了解案件级配置
- 编辑 `format-spec.md` 定制排版参数
