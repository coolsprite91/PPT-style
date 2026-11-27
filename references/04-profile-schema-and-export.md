# 04｜三层活资产档案结构、绑定混搭、多轮时间线与导出分享（V0.2.0）

> 本架构实现「**L1 写稿场景 × L2 多套单页框架与表达 × L3 多套真实母版皮肤**」的积木式活资产沉淀与自由排列组合，同时支持多轮人机交替修订时间线与一键导出分享。

---

## 一、全局档案目录结构（`~/.ppt-style/`）

所有跨项目持久化的风格档案、修订时间线快照与备份均存放在全局隐藏目录 `~/.ppt-style/` 中（二级沙箱降级时回落至当前工作区隐藏目录），**绝不污染用户的业务工作文件夹**：

```text
~/.ppt-style/
├── registry.json                     # 最近活跃工坊指针、默认档案 ID、默认皮肤 ID
├── snapshots/
│   ├── index.json                    # 各文档 doc_id、latest_rev_id、reviewed_rev_id、revisions[]
│   └── <doc_id>/                     # 每轮快照文件 snap-R1-...pptx / snap-R2-...pptx（>50MB 剥离音视频）
├── backups/                          # 旧版技能备份区（绝不放在全局 skills/ 同级目录）
└── profiles/
    └── <profile_id>/                 # 默认为 default
        ├── core_l1_planning.json     # L1：scenarios[]（写稿场景库）+ rules[]（页面规划习惯）
        ├── core_l2_content.json      # L2：frameworks[]（按场景挂接的多套单页框架与置信度）+ rules[] + settings
        ├── masks/
        │   ├── templates/            # L3：从代表作剥离的真实空白母版文件 base_<mask_id>.pptx
        │   ├── default.json          # L3：theme_tokens + template_pptx + binding_mode + bound_scenarios + rules[]
        │   └── <mask_id>.json        # 可选拆出或重采的多套视觉母版皮肤
        └── history/                  # 每次更新规则前的自动版本快照，支持一键回滚
```

---

## 二、核心管理与排列组合命令（全有大白话口令）

### 1. L1 / L2 / L3 资产管理与绑定混搭
- **确认或补充写稿场景**（`profile confirm-scenario`）：把旧稿没覆盖到但平时常写的场景（如“年度工作总结”）加进 L1，支持自定义常见节奏骨架。
- **确认一般置信度框架与优化建议**（`profile confirm-framework --decision accept|adjust|drop`）：对置信度一般的单页框架，按用户选择原样收录、按建议优化版收录或丢弃。
- **配置母版与场景的绑定/独立关系**（`profile bind-mask`）：
  - 默认绑定模式（`binding_mode: "bound"`）：一发起某 L1 场景自动套用对应 L3 母版；
  - 独立解耦模式（`binding_mode: "independent"`）：不绑死场景，随时按需切换。
- **写作前解析组合**（`profile resolve-combo --scenario <场景> [--mask <指定皮肤>]`）：自动取出“L1 场景 ＋ 该场景配套的 L2 框架 ＋ 绑定（或指定混搭）的 L3 真实母版与精确色板”。

### 2. 多面具拆分、跳槽换壳与反悔重置
- **拆分场景皮肤**（`profile split-mask`）：保留共用 L1/L2 内核，按场景拆出独立 L3 皮肤；
- **跳槽换新公司模板**（`profile swap-visual-shell`）：先存历史快照，**锁定保留全部 L1 场景、L2 框架与写作习惯**，仅清空指定皮肤的 L3 视觉壳，引导丢入新公司模板重采；
- **查看/删除单条规则**（`profile delete-rule`）：按大白话编号展示与删除（自动留快照）；
- **回滚与出厂重置**（`profile rollback / reset`）：一键退回上一次好版本或安全重置。

### 3. 多轮修订时间线与检查点管理（`snapshot_diff.py`）
- **查看多轮记录**：`snapshot_diff.py history --pptx <文件>`；
- **单轮增量比对＋推进检查点**：`snapshot_diff.py diff --pptx <文件> --mode latest --mark-reviewed`；
- **跨多轮纯人工汇总**：`snapshot_diff.py diff --pptx <最终定稿> --mode cumulative_human`（只汇总 `human_edit` 区间，剔除中间 `ai_revise`）。

### 4. 按需打包导出独立分享包（`profile export-share`）
- 将目标档案对应的三层规则（均经过 `sanitize-rule` 脱敏确认无敏感业务词与数字）＋ `PPT-style` 运行框架打包成独立自包含文件夹；同事双击安装即可直接按该风格使用。

### 5. 调用指南查询通道（`usage-guide`）
- 任何时候用户说“怎么发起写 PPT”或“怎么调用我的风格”，调用 `workspace_manager.py usage-guide` 生成固定 5 段范式大白话指南（自动过 `output_guard.py` 脱敏校验）。
