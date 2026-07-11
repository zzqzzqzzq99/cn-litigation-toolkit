# profile.md · 实践画像模板

> 本文件由冷启动访谈（onboarding-interview）初始写入，用户可随时手工编辑同步。
> 所有 skill 在运行时按能力槽（LAW/BIZ/PARSE/DOCX/SEARCH/KB/COLLAB）从本文件读取用户配置与降级规则。

---

## 实践画像

### 1. 身份与角色

```yaml
role:          # 执业律师 / 公司法务 / 兼任
stance:        # 主诉为主 / 被诉为主 / 混合
team:          # 团队/组织名称
industry:      # 主要业务领域
case_types:    # 案件类型
```

### 2. 风格与偏好

```yaml
risk_calibration:   # 风险偏好（一句话描述）
document_tone:      # 文书语调（正式 / 半正式 / 因场景而异）
reporting_style:    # 汇报风格
sample_documents:   # 参考文书样本路径
```

### 3. 案件管理

#### 3.1 台账绑定

```yaml
ledger_backend:     # COLLAB 平台在线表格 / 本地 JSON+Markdown / 手工维护
ledger_id:          # baseId 或本地路径
ledger_fields: |
  # 台账字段定义
```

#### 3.2 沟通渠道分组

```yaml
case_channels: |
  # 案件沟通群组映射
patrol_blacklist:   # 巡检黑名单
```

#### 3.3 期限与提醒

```yaml
reminder_channel: IM推送      # IM推送 / 日历日程 / 仅查看
reminder_time: 09:00          # 定时提醒时间
holiday_calendar:             # 节假日日历来源
```

#### 3.4 排版规范绑定

```yaml
format_spec_path:   # format-spec.md 或 format-spec.<court>.md 路径
```

### 4. 知识库

```yaml
kb_backend:         # qmind / ima / obsidian / 飞书知识库 / 钉钉知识库 / 语雀 / 本地 md / 暂不配置
kb_registry: |
  # 知识库注册表
  # 板块 | 位置 | 状态
```

### 5. 外部能力后端（能力槽映射）

```yaml
LAW.*:              # 法律检索后端（元典 / 北大法宝）
BIZ.*:              # 工商信息后端（企查查 / 天眼查 / 启信宝）
PARSE.to_markdown:  # 文档解析后端（MinerU / 阿里云百炼 / markdownify）
DOCX.*: md2docx_legal.py  # Word 转换后端
SEARCH.*:           # 全网搜索后端
KB.*:               # 知识库后端
COLLAB.*:           # 协作平台后端（钉钉 / 飞书）
```

### 6. 共享护栏

> 核验闸门（法律核验闸门 / 文书排版闸门）的完整定义与操作流程见 `Expert.md`「共享护栏」节。profile.md 不重复。

#### 6.2 四级溯源体系

| 级别 | 含义 |
|------|------|
| L1 | 经法律检索/工商后端核验 |
| L2 | 用户上传原件/法院文书 |
| L3 | 用户陈述 |
| L4 | 模型推断-待验证（后端不可用时的降级标记，须人工复核） |

#### 6.3 MCP 预检协议

执行任何需要外部能力的技能前，先探测当前会话可用的 MCP 服务，按能力槽映射使用。详见 `Expert.md`。

#### 6.4 降级规则

| 后端类别 | 降级标注 | 降级行为 |
|---------|---------|---------|
| 法律检索 (LAW.*) | `[L4-法条待验证]` | 保留法条引用但标注 |
| 工商信息 (BIZ.*) | `[L4-主体信息待验证]` | 保留用户提供的主体信息但标注 |
| 协作平台 (COLLAB.*) | `[本地模式-协作平台未连接]` | 回退本地 JSON + Markdown |
| 全网搜索 (SEARCH.*) | `[L4-信息待验证]` | 依赖模型内置知识并标注 |
| document-parse (PARSE.*) | `[视觉降级-待人工复核]` | 按降级链回退 |
| Word 转换 (DOCX.*) | 无标注 | 仅交付 Markdown |
| 知识库 (KB.*) | `[知识库不可用]` | 仅依赖对话内上下文 |

### 7. 联动扩展

```yaml
downstream_writeback: 启用    # 产出案件成果后自动回写案件管家台账
auto_visualization: 询问一次   # 事实梳理后是否自动生成可视化图表
```

---

## 案件级配置（Matter Profile）

全局画像定义"我是谁"，案件画像定义"这个案件用什么策略"。

- **路径**：`matters/{slug}/matter.md`（由案件管家建档时自动生成）
- **激活**：说"切换到 X 案"
- **覆盖规则**：`matter.md` 有值字段覆盖本文件同名字段；优先级 `matter.md` > `profile.md` > 默认值
