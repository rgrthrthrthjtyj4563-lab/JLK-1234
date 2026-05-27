---
name: "JLK-Pt-skill"
description: "Use when generating patient questionnaire analysis reports for pharmaceutical clients from uploaded survey spreadsheets or questionnaire tables, especially when the output must inherit the original Word template, including TOC, preface, header, title hierarchy, tables, and chart styles."
---

# JLK Patient Report Skill

## When To Use
- 输入是患者问卷数据或问卷统计表。
- 输出必须是 `Word 报告`。
- 报告对外标题统一使用通用名称 `问卷调研分析报告`，不直接暴露内部模板分类名。
- 文档必须继承原模板的页面组件：
  - `问卷调研服务结算`
  - `目录`
  - `前言`
  - 报告大标题
  - 页眉
- 正文结构必须贴合真实模板：
  - `项目背景`
  - `项目开展情况`
  - `问卷说明`
  - `问卷结果分析`
  - `调研结果`
  - `附件1`
  - `附件2`
  - `免责申明`
- `问卷结果分析` 必须使用模板固定 `4.x + 小标题 + 表格 + 分析` 结构，每道题独立为一个分析单元：`小标题 → 统计表 → 分析正文`。多题不得合并在同一小标题下。

## Workflow
1. 用 `scripts/parse_questionnaire.py` 解析附件为 `questionnaire.json`。
2. 用 `scripts/cluster_dimensions.py` 识别问卷维度结构。当 `report_content.md` 的 front matter 中包含 `dimensions_json` 时，优先使用 AI 声明的维度定义（维度名、引言、小标题、图表配置）；否则回退到内置硬编码模板匹配。
3. AI 按模板章节撰写 `report_content.md`，但只允许提供：
   - `前言` 正文
   - `项目背景` 正文
   - 各 `4.x` 小节的分析正文
   - `5.2 调研结果分析` 正文
   - `5.3 建议` 正文
   不能决定模板分类、`4.x` 标题、`4.x` 引言、`5.1` 重点题目或图表点位。`4.x` 小标题由程序语义映射给出归纳性兜底值（如"漏服应对行为分析"），AI 草稿可在此基础上优化覆盖，但校验拒绝口语化截取类型的标题（如"您忘记服药后通常分析"）。
   - 可选：在 front matter 中通过 `dimensions_json` 字段声明维度结构（JSON 格式），AI 可以据此为任意新品种定义维度名称、引言、小标题和图表配置，实现新品种的动态适配。
   - `dimensions_json` 格式示例：
     ```json
     {"dimensions": [
       {"name": "药物认知与信息获取", "intro": "本维度用于观察...",
        "subtopics": [{"patterns": ["作用", "认知"], "subtitle": "药物认知情况分析"}],
        "charts": [{"patterns": ["认知"], "chart_type": "pie", "chart_style_profile": "efficacy_pie"}]}
     ]}
     ```
   - `dimensions_json` 会执行硬校验：`dimensions` 不能为空，维度 `name/intro/subtopics` 必填，`patterns` 必须是可编译正则，图表类型与样式必须在程序白名单内；校验失败时 pipeline 直接报错，不回退硬编码模板。
   - `前言` 必须强制为 `2` 段，总字数强制收敛在约 `400` 字区间内，写作顺序固定为：`疾病/品类背景 -> 产品价值 -> 区域与对象 -> 调研目的 -> 应用价值`。
   - `项目背景` 必须为 `2` 段，第一段写 `宏观背景/区域特征`，第二段写 `现有数据缺口/问题` 与 `本次调研必要性`。
   - `前言` 与 `项目背景` 必须采用正式、客观的报告文风，不得写成营销文案、空泛套话或与模板无关的通用降压药兜底表述。
   - `前言` 与 `项目背景` 不得大面积重复；缺少区域、对象、目的或价值落点时，视为不合格草稿。
   - `问卷结果分析` 下每个小标题对应 `1` 道题，分析正文必须为 `1` 段，严格控制在 `250-300` 字之间。`cluster_dimensions.py` 会自动将多题 subtopic 拆分为每题一个 subtopic。
   - 单题分析必须采用 `主结论 -> 解释/风险 -> 收束` 的报告式逻辑，不得写成程序化套话拼接。
   - 单题分析正文不得出现 `A/B/C/D`、`选项A/选项B` 等字母选项列举，应直接使用选项语义或归纳表达。
4. 用 `scripts/build_payload.py` 构建 `report_payload.json`。如果 AI 草稿中的 `4.x` 章节集合、标题或小标题与 `cluster_dimensions` 骨架冲突，必须直接报错，不做兼容。
   - `问卷结果分析` 导语中的维度名称和维度数量必须与 `4.x` 实际章节完全一致。
   - `用药体验与疗效反馈` 模板的 `问卷结果分析` 固定包含 `2` 张概览图：`维度占比饼状图` 与 `维度横向柱形图`。
   - 两张概览图的统计口径统一为各 `4.x` 维度归入的题目数量，维度顺序必须与 `4.1-4.7` 顺序一致。
   - 第一张概览图必须保持原生 Office 饼图；第二张概览图必须改为 PNG 图片形式的“横向柱状图 + 维度题目数量”输出，不能退化为饼图、占比图或依赖 `chart2.xml + Workbook2.xlsx` 的重算结果。
   - `5.1 问卷重点问题分析` 必须保留 `2` 张原生 `3D` 饼图，并与第 `4` 部分的两张概览图并存，不能因为修第 `4` 部分图表而删除 `5.1` 双图。
5. `5.1 问卷重点问题分析` 默认从模版内置的维度优先级中选题（依从性模版为 `4.2→4.3`，疗效模版为 `4.1→4.2`）。可通过 `report_content.md` front matter 中的 `key_issue_sections` 字段覆盖选题优先级：
   - 格式：JSON 数组字符串，`key_issue_sections: '["4.4", "4.2"]'`，或逗号分隔 `key_issue_sections: '4.4,4.2'`
   - 校验：声明的编号必须存在于实际 `4.x` 章节中，否则 pipeline 直接报错
   - 限制：最多取前 2 个有效值，程序仍负责标题映射和图表生成
   - 未声明时完全回退到内置默认顺序，零影响
6. 用 `scripts/render_from_template.py` 基于模板底稿做**对象级替换**输出最终 `docx`。
     - 这是主渲染器，在模板文档中定位锚点并替换内容，不清空正文。
     - 保留模板所有样式：section 断点、页眉、表格、图表、字体。
     - `scripts/render_report.py` 是旧版「清空重写」方案，已弃用。
     - `04_outputs/` 中若存在旧版 `report_content.md / report_final.md / render_report.py` 等文件，只能视为历史脏产物，不能作为当前结构参考。
     - `附件1` 必须按题目块整体重建：题干使用 `（1）（2）...` 顺序号，去掉原题号前缀如 `1.` `6.` `11.`，选项仍保留 `A./B./C./D.` 且必须按原始问卷顺序逐题展开。
7. 运行时优先使用 `scripts/run_report_pipeline.py`，它会为每次生成创建独立运行目录，避免复用固定的 `tmp/docs/content.md` 或 `generated.docx`。
8. `scripts/run_report_pipeline.py` 必须在渲染后调用 `scripts/final_validate_docx.py` 做最终 Word 验收；验收失败时必须直接失败并移除最终交付 docx，只保留运行目录诊断文件。
9. 检查输出必须保留目录、前言、页眉、标题层级、蓝底表格和模板图表风格。
10. `scripts/build_payload.py` 会对 `前言`、`项目背景` 和 `问卷结果分析` 单题正文执行硬校验；段数、字数、区域信息、结构或旧式百分比分析不达标时，pipeline 必须直接失败，不得静默降级。
11. 如果模板底稿缺少第 `4` 部分的第二张概览图位，或维度配置数量与实际 `4.x` 章节不一致，pipeline 必须直接失败。

## Supported Templates
- `用药体验与疗效反馈`（内置硬编码回退）
- `依从性与用药习惯`（内置硬编码回退）
- 任意新品种：通过 `report_content.md` front matter 中的 `dimensions_json` 字段动态声明维度，无需硬编码

## Chapter 4 Format Rules (must match reference template)
- `4.x` heading format: `4.1Title` (no space between number and title), font Hanyi Zhongsong 16pt Bold, `pageBreakBefore=true`, `keepNext=true`, line spacing `line=480`.
- Subtitle format: plain text only (no `(1)` index prefix), font Hanyi Zhongsong 12pt Bold, `pageBreakBefore=true`, `keepNext=true`, line spacing `line=600`, `firstLineChars=0` (no first-line indent).
- Analysis paragraph: Hanyi Zhongsong 12pt, justified alignment, first-line indent 2 chars (`firstLineChars=200`), line spacing `line=600`, `pageBreakBefore=true`, `keepNext=true`.
- Table style: `tblW=5000 pct` (full page width), header/first-column bg `4684D3`, data-cell bg `D9EAF7`, row height `atLeast 567 dxa`, font Hanyi Zhongsong 10pt, line spacing `line=240` (single), border color `FFFFFF`.
- Each question follows the fixed pattern: `Subtitle -> Statistics Table -> Analysis Paragraph`. One subtopic contains exactly one question; multi-question stacking under one subtitle is prohibited.

## Output Notes
- 图表必须按模板定义的点位和风格生成，不接受“近似即可”的退化方案。
- 正文版式、目录、页眉、标题层级和表格样式必须继承模板底稿，不沿用旧绿色标题系统。

## 5.3 建议 AI 写作规范

### 输出位置
在 `report_content.md` 的 `## 调研结果` 章节末尾，以 `### 5.3 建议` 为标题输出。前面必须已经输出 `### 5.1 问卷重点问题分析` 和 `### 5.2 调研结果分析`。

### 结构要求
1. **导语**（1 段，50-80 字）：以「基于/结合调研结果，为进一步……，提出以下建议：」开头，必须点名 `{product}` 和本次调研发现的核心问题方向（如依从性、认知偏差、监测缺失等）。
2. **分条建议**（2-4 条，不要写「（1）」这种中文编号，用阿拉伯数字+英文句点，如 `1. `）：
   - 每条建议前面**不加小标题**，直接以编号开头进入正文。
   - 每条 80-150 字，总字数控制在 300-600 字。

### 角度选择规则（动态池，不要求每次全选）
AI 必须从以下角度池中**动态挑选 2-4 个最贴合本文分析结果的角度**，每次输出不要固定同样的组合，优先匹配前文暴露的真实问题：

- **临床规范教育**：针对自行调剂量、监测缺失、联合用药未确认等问题，提出医生/药师/复诊层面的干预动作。
- **产品与服务体验优化**：针对包装不便、说明书不清晰、指导渠道不足等反馈，从取用便利性、说明表达、咨询触点等角度提出改进。
- **渠道/监管/供应**：针对非正规购药、缺货、价格敏感等问题，提出渠道引导或供应保障建议。
- **监测与随访机制**：针对血压监测不规律、随访脱节等问题，提出记录表、提醒工具、医患沟通群等手段。
- **健康教育创新**：针对认知误区、宣教形式单一等问题，提出短视频、口诀宣传册、讲座等多样化手段。
- **品牌价值/场景拓展**（中成药适用）：针对疗程认知不足、潜在患者挖掘等，提出差异化传播或场景适配。

### 每条建议的三层公式（缺一不可）
1. **靶子层**：针对什么问题/哪类人群（如「针对血压偶尔波动、自行调剂量的重点群体」）。
2. **手段层**：通过什么具体形式/载体/渠道（如「通过复诊一对一宣讲、线上科普短视频、患者手册等」）。
3. **目的层**：达到什么效果（如「引导患者养成规律监测习惯，精准优化治疗方案」）。

### 语言与风格约束
- 动词必须具体：强化、优化、完善、拓展、加强、引导、搭建、开展、普及、提升、补充、精准化、规范化、差异化。
- 必须出现可落地的载体或工具（如「易撕口包装」「用药口诀宣传册」「血压监测记录表」「医患沟通群」），禁止纯口号。
- 禁止出现「有关部门要高度重视」「加强管理」「提高认识」等没有执行路径的空话。
- 禁止重复前文具体百分比数据（不写「XX%的患者……」）。
- 禁止写成「总之」「综上所述」等总结词。
- 禁止对药品疗效本身下结论（如「该药疗效确切」），只提改进空间。
- 建议应从具体方面或机制展开，如临床规范、服务体验、渠道保障、随访监测、健康教育等；禁止使用「药企」「药企方面」作为建议主体或小标题，避免写成主体罗列。

### 多样性机制
每次生成时，AI 应在前文分析结果的基础上，自主选择切入角度和排列顺序。鼓励使用不同的具体载体示例（如这次写「线上科普短视频」，下次可写「门诊随访卡片」或「药店用药咨询角」），避免每次都输出同样的三条。
