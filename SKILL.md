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
- `问卷结果分析` 必须使用模板固定 `4.x + 小标题 + 表格 + 图表 + 分析` 结构，而不是旧 skill 的逐题小节结构。

## Workflow
1. 用 `scripts/parse_questionnaire.py` 解析附件为 `questionnaire.json`。
2. 用 `scripts/cluster_dimensions.py` 识别问卷属于哪类模板，并生成唯一允许的 `4.x / 标题 / 引言 / 小标题 / 图表点位 / 5.1重点题目` 骨架。
3. AI 按模板章节撰写 `report_content.md`，但只允许提供：
   - `前言` 正文
   - `项目背景` 正文
   - 各 `4.x` 小节的分析正文
   - `5.2 调研结果分析` 正文
   不能决定模板分类、`4.x` 标题、`4.x` 引言、小标题、`5.1` 重点题目或图表点位。
4. 用 `scripts/build_payload.py` 构建 `report_payload.json`。如果 AI 草稿中的 `4.x` 章节集合、标题或小标题与 `cluster_dimensions` 骨架冲突，必须直接报错，不做兼容。
5. 用 `scripts/render_from_template.py` 基于模板底稿做**对象级替换**输出最终 `docx`。
     - 这是主渲染器，在模板文档中定位锚点并替换内容，不清空正文。
     - 保留模板所有样式：section 断点、页眉、表格、图表、字体。
     - `scripts/render_report.py` 是旧版「清空重写」方案，已弃用。
     - `04_outputs/` 中若存在旧版 `report_content.md / report_final.md / render_report.py` 等文件，只能视为历史脏产物，不能作为当前结构参考。
6. 运行时优先使用 `scripts/run_report_pipeline.py`，它会为每次生成创建独立运行目录，避免复用固定的 `tmp/docs/content.md` 或 `generated.docx`。
7. 检查输出必须保留目录、前言、页眉、标题层级、蓝底表格和模板图表风格。

## Supported Templates
- `用药体验与疗效反馈`
- `依从性与用药习惯`

## Output Notes
- 图表必须按模板定义的点位和风格生成，不接受“近似即可”的退化方案。
- 正文版式、目录、页眉、标题层级和表格样式必须继承模板底稿，不沿用旧绿色标题系统。
