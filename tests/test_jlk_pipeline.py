from __future__ import annotations

import re
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from collections import Counter
from zipfile import ZipFile

from docx import Document

ROOT = Path(__file__).resolve().parents[1]

import sys

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.build_payload import build_payload, normalize_survey_period, parse_markdown_content, validate_payload
from scripts.cluster_dimensions import cluster_dimensions
from scripts.render_from_template import TemplateRenderer


class Namespace:
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)


def efficacy_questionnaire() -> dict:
    questions = [
        "您服用厄贝沙坦氢氯噻嗪片后，血压控制效果如何？",
        "您服用厄贝沙坦氢氯噻嗪片期间，是否出现过头晕、头痛症状？",
        "您服用厄贝沙坦氢氯噻嗪片期间，是否出现过口渴、多尿情况？",
        "您通常通过什么渠道购买厄贝沙坦氢氯噻嗪片？",
        "您认为厄贝沙坦氢氯噻嗪片的价格对您的用药负担影响如何？",
        "您服用厄贝沙坦氢氯噻嗪片的剂量是否严格按照医嘱？",
        "您对厄贝沙坦氢氯噻嗪片的服药频率满意度如何？",
        "您服用厄贝沙坦氢氯噻嗪片期间，是否同时服用其他降压药物？",
        "您是否咨询过医生，确认厄贝沙坦氢氯噻嗪片与其他药物可同时服用？",
        "您所在地区的厄贝沙坦氢氯噻嗪片供应是否稳定？",
        "您服用厄贝沙坦氢氯噻嗪片期间，是否定期监测血压？",
        "您对厄贝沙坦氢氯噻嗪片药品包装的便利性评价如何？",
        "您对厄贝沙坦氢氯噻嗪片说明书中 “用法用量”“不良反应” 的标注清晰度评价如何？",
        "您认为当前获取的用药指导是否足够详细准确？",
    ]
    options = [
        ("A", "选项A", "700", "39.13%"),
        ("B", "选项B", "600", "33.54%"),
        ("C", "选项C", "300", "16.77%"),
        ("D", "选项D", "189", "10.56%"),
    ]
    data = []
    for i, q in enumerate(questions, start=1):
        data.append(
            {
                "number": i,
                "question": q,
                "total": "1789",
                "options": [{"label": a, "text": b, "count": c, "pct": d} for a, b, c, d in options],
            }
        )
    return {"question_count": len(data), "questions": data}


def sample_markdown() -> str:
    return """---
product: 厄贝沙坦氢氯噻嗪片
region: 北京市
time: 2025.10
attachment_name: 厄贝沙坦氢氯噻嗪片用药体验与疗效反馈患者调查问卷
survey_period: 2025年10月01日——2025年10月31日
valid_count: 1789
disclaimer_unit: 北京玖麟空科技有限公司
---

## 前言

本次患者调研围绕厄贝沙坦氢氯噻嗪片的真实使用场景展开，重点收集患者在疗效感知、安全性体验以及长期管理行为等方面的反馈，为后续优化临床沟通和患者支持提供参考。

## 项目背景

本次调研围绕患者使用厄贝沙坦氢氯噻嗪片过程中的真实体验展开，重点观察疗效感知、安全性体验、行为习惯以及信息支持等关键环节。

## 项目开展情况

本项目采用患者问卷方式开展。

## 问卷说明

本次分析采用结构化处理方式。

## 问卷结果分析

本次问卷从7个维度展开统计分析。

### 4.1药品疗效

本维度用于观察患者对控压效果的主观感知。

#### 血压控制效果分析

多数患者在该题中的反馈集中在正向选项，说明控压结果的主流感知较为稳定。

### 4.2药品安全性

本维度用于观察患者对常见不适反应的主观感知。

#### 头晕头痛不良反应分析

多数患者并未将该症状视为主要负担。

#### 口渴多尿不良反应分析

患者对该类不适的整体反馈同样偏轻。

### 4.3用药行为与习惯

本维度用于观察患者的长期行为执行情况。

#### 购药渠道分析

渠道选择主要集中于正规渠道。

#### 剂量遵循情况分析

多数患者能够保持较稳定的剂量执行。

#### 联合用药情况分析

部分患者存在联合用药需求。

#### 血压监测频率分析

规律监测行为仍有继续提升空间。

### 4.4用药便利性

本维度用于观察频率与包装体验。

#### 服药频率满意度分析

大多数患者认为服药频率可以接受。

#### 药品包装便利性分析

包装便利性整体较好。

### 4.5药品经济性

本维度用于观察价格感知。

#### 价格负担影响分析

多数患者并未将价格视为主要障碍。

### 4.6药品可及性

本维度用于观察供应稳定性。

#### 药品供应稳定性分析

大多数患者认为购药较为顺畅。

### 4.7用药指导信息评价

本维度用于观察说明信息和指导支持。

#### 说明书信息清晰度分析

大多数患者能够理解说明书中的核心信息。

#### 用药指导详细准确性分析

现有指导信息总体能够满足基础使用需求。

## 调研结果

### 5.1问卷重点问题分析

本次问卷重点问题主要集中在用药行为与习惯及用药指导信息评价两个层面。

### 5.2调研结果分析

整体来看，患者对疗效、安全性和便利性的反馈基础较好。

### 5.3建议

1. 补充场景化患者教育材料。
"""


class PipelineTest(unittest.TestCase):
    def test_cluster_dimensions_uses_efficacy_template(self) -> None:
        grouped = cluster_dimensions(efficacy_questionnaire())
        self.assertEqual(grouped["template_type"], "用药体验与疗效反馈")
        self.assertEqual(grouped["sections"][0]["section_title"], "药品疗效")
        self.assertEqual(grouped["sections"][0]["section_intro"], "本维度用于观察患者对控压效果的主观感知。")
        self.assertGreater(len(grouped["sections"]), 0)
        section_1 = grouped["sections"][0]
        self.assertGreater(len(section_1["subtopics"]), 0)
        subtopic_1 = section_1["subtopics"][0]
        self.assertIn("subtopic_index", subtopic_1)
        self.assertIn("question_refs", subtopic_1)
        self.assertEqual(subtopic_1["subtitle"], "血压控制效果分析")
        self.assertGreater(len(subtopic_1["question_refs"]), 0)

    def test_build_payload_matches_template_slots(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            report_content = Path(temp_dir) / "content.md"
            report_content.write_text(sample_markdown(), encoding="utf-8")
            meta, content = parse_markdown_content(report_content)
            payload = build_payload(
                efficacy_questionnaire(),
                meta,
                content,
                Namespace(
                    product=None,
                    region=None,
                    time=None,
                    attachment_name=None,
                    survey_period=None,
                    sample_size=None,
                    valid_count=None,
                    disclaimer_unit=None,
                ),
            )
        validate_payload(payload)
        self.assertEqual(len(payload["preface"]), 1)
        self.assertEqual(len(payload["project_background"]), 1)
        self.assertNotRegex(payload["preface"][0], r"#+")
        self.assertNotRegex(payload["project_background"][0], r"#+")
        self.assertEqual(
            payload["project_execution"]["lines"],
            [
                "项目工具：调查问卷，14道选择题。其内容涵盖药品疗效、药品安全性、用药行为与习惯、用药便利性、药品经济性、药品可及性6大维度，全面覆盖患者用药全流程关键节点。",
                "样本采集范围：北京市",
                "样本采集数量：本次共收集筛选到有效问卷1789份。",
                "样本采集时间：2025年10月01日——2025年10月31日",
            ],
        )
        self.assertEqual(payload["questionnaire_note"]["intro"], "为提升数据可比性与分析结果的科学性，本研究采用以下标准化统计处理方法对原始问卷数据进行系统规整：")
        self.assertEqual(len(payload["questionnaire_note"]["items"]), 4)
        self.assertEqual(
            payload["result_analysis"]["intro"],
            ["本次调查采用问卷调查方式，共有14个问题，本报告从药品疗效、药品安全性、用药行为与习惯、用药便利性、药品经济性、药品可及性等6个维度展开统计分析。"],
        )
        self.assertEqual(payload["result_analysis"]["sections"][0]["section_title"], "药品疗效")
        self.assertEqual(payload["result_analysis"]["sections"][0]["section_intro"], ["本维度用于观察患者对控压效果的主观感知。"])
        self.assertEqual(payload["result_analysis"]["sections"][0]["subtopics"][0]["subtitle"], "血压控制效果分析")
        self.assertEqual(payload["result_analysis"]["sections"][0]["visual_groups"][0]["chart_type"], "pie")
        self.assertEqual(payload["report_title"], "问卷调研分析报告")
        self.assertEqual(payload["header_text"], "问卷调研分析报告")
        self.assertEqual(payload["meta"]["survey_period"], "2025年10月01日——2025年10月31日")
        self.assertEqual(payload["service"]["unit"], "北京玖麟空科技有限公司")
        self.assertRegex(payload["service"]["date"], r"^2025年11月\d{2}日$")
        service_dt = datetime.strptime(payload["service"]["date"], "%Y年%m月%d日").date()
        self.assertGreaterEqual(service_dt, date(2025, 11, 1))
        self.assertLessEqual(service_dt, date(2025, 11, 30))
        self.assertEqual(payload["disclaimer"]["unit"], payload["service"]["unit"])
        self.assertEqual(payload["disclaimer"]["date"], payload["service"]["date"])
        self.assertEqual(payload["summary"]["key_issue_analysis"], payload["summary"]["key_issue_analysis_programmatic"])
        self.assertEqual(len(payload["summary"]["key_issue_items"]), 2)
        self.assertEqual(payload["summary"]["key_issue_items"][0]["heading"], "1. 血压控制现状与特征")
        self.assertEqual(payload["summary"]["key_issue_items"][1]["heading"], "2. 药品价格接受度情况")
        self.assertGreaterEqual(len(payload["summary"]["overall_analysis"]), 4)
        self.assertLessEqual(sum(len(paragraph) for paragraph in payload["summary"]["overall_analysis"]), 700)
        self.assertEqual(payload["summary"]["overall_analysis"], payload["summary"]["overall_analysis_programmatic"])
        self.assertEqual(len(payload["summary"]["key_issue_analysis"]), 4)
        self.assertEqual(len(payload["summary"]["recommendations"]), 3)

    def test_build_payload_rejects_conflicting_draft_structure(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            report_content = Path(temp_dir) / "content.md"
            report_content.write_text(
                sample_markdown()
                .replace("### 4.1药品疗效", "### 4.1药物认知与信息获取")
                .replace("#### 血压控制效果分析", "#### 药物认知情况分析"),
                encoding="utf-8",
            )
            meta, content = parse_markdown_content(report_content)
            with self.assertRaisesRegex(ValueError, "AI draft"):
                build_payload(
                    efficacy_questionnaire(),
                    meta,
                    content,
                    Namespace(
                        product=None,
                        region=None,
                        time=None,
                        attachment_name=None,
                        survey_period=None,
                        sample_size=None,
                        valid_count=None,
                        disclaimer_unit=None,
                    ),
                )

    def test_normalize_survey_period_accepts_compact_chinese_range(self) -> None:
        self.assertEqual(
            normalize_survey_period("2025年9月1日至9月30日"),
            "2025年09月01日——2025年09月30日",
        )
        self.assertEqual(
            normalize_survey_period("2025年9月1日-10月2日"),
            "2025年09月01日——2025年10月02日",
        )

    def test_render_from_template_preserves_template_pages_and_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            report_content = Path(temp_dir) / "content.md"
            report_content.write_text(sample_markdown(), encoding="utf-8")
            meta, content = parse_markdown_content(report_content)
            payload = build_payload(
                efficacy_questionnaire(),
                meta,
                content,
                Namespace(
                    product=None,
                    region=None,
                    time=None,
                    attachment_name=None,
                    survey_period=None,
                    sample_size=None,
                    valid_count=None,
                    disclaimer_unit=None,
                ),
            )
            output_docx = Path(temp_dir) / "rendered.docx"
            TemplateRenderer(Path(payload["meta"]["template_doc"]), payload).render(output_docx)

            document = Document(output_docx)
            texts = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
            all_text = "\n".join(texts)
            expected_table_questions = [
                visual["table_data"]["question"]
                for section in payload["result_analysis"]["sections"]
                for visual in section["visual_groups"]
            ]
            for required in ["问卷调研服务结算", "目录", "前言", "项目背景", "项目开展情况", "问卷说明", "问卷结果分析", "调研结果", "免责申明"]:
                self.assertIn(required, texts)
            self.assertLess(texts.index("目录"), texts.index("前言"))
            self.assertIn(payload["report_title"], texts)
            self.assertIn("4.1药品疗效", texts)
            self.assertIn("血压控制效果分析", texts)
            self.assertIn("附件2：问卷调查明细表", texts)
            self.assertNotIn("E.E", all_text)
            self.assertNotRegex(all_text, r"\n#+")
            self.assertGreaterEqual(len(document.tables), len(expected_table_questions) + 1)
            self.assertEqual(len(document.sections), 5)
            self.assertEqual(document.sections[0].header.paragraphs[0].text, payload["header_text"])
            self.assertIn("项目工具：调查问卷，14道选择题。其内容涵盖药品疗效、药品安全性、用药行为与习惯、用药便利性、药品经济性、药品可及性6大维度，全面覆盖患者用药全流程关键节点。", texts)
            self.assertIn("样本采集时间：2025年10月01日——2025年10月31日", texts)
            self.assertIn(f"服务单位：{payload['service']['unit']}", texts)
            self.assertIn(f"日期：{payload['service']['date']}", texts)
            self.assertLess(texts.index(f"服务单位：{payload['service']['unit']}"), texts.index(f"日期：{payload['service']['date']}"))
            self.assertIn(f"服务提供单位:{payload['service']['unit']}", texts)
            self.assertIn(payload["service"]["date"], texts)
            service_para = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == f"服务单位：{payload['service']['unit']}")
            date_para = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == f"日期：{payload['service']['date']}")
            self.assertEqual(service_para.alignment, 2)
            self.assertEqual(date_para.alignment, 2)
            key_issue_idx = texts.index("5.1问卷重点问题分析")
            key_issue_end = texts.index("5.2调研结果分析")
            key_issue_body = texts[key_issue_idx + 1:key_issue_end]
            self.assertIn("1. 血压控制现状与特征", key_issue_body)
            self.assertIn("2. 药品价格接受度情况", key_issue_body)
            self.assertEqual(len(key_issue_body), 4)
            title_para_1 = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == "1. 血压控制现状与特征")
            title_para_2 = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == "2. 药品价格接受度情况")
            self.assertEqual(title_para_1.alignment, 0)
            self.assertEqual(title_para_2.alignment, 0)
            overall_idx = texts.index("5.2调研结果分析")
            overall_end = texts.index("5.3建议")
            overall_paragraphs = texts[overall_idx + 1:overall_end]
            self.assertGreaterEqual(len(overall_paragraphs), 4)
            self.assertLessEqual(sum(len(paragraph) for paragraph in overall_paragraphs), 700)
            first_key_issue_para = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == key_issue_body[1])
            self.assertEqual(first_key_issue_para.alignment, 3)
            self.assertEqual(first_key_issue_para.paragraph_format.line_spacing, 2.5)
            self.assertEqual(first_key_issue_para.paragraph_format.first_line_indent, 304800)
            disclaimer_heading = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == "免责申明")
            disclaimer_item = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip().startswith("（1）"))
            disclaimer_unit = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip() == f"服务提供单位:{payload['service']['unit']}")
            self.assertEqual(disclaimer_heading.alignment, 1)
            self.assertEqual(disclaimer_item.alignment, 3)
            self.assertEqual(disclaimer_item.paragraph_format.first_line_indent, 0)
            self.assertEqual(disclaimer_unit.alignment, 2)
            attachment_question = next(paragraph for paragraph in document.paragraphs if "您服用厄贝沙坦氢氯噻嗪片后，血压控制效果如何？" in paragraph.text.strip())
            attachment_option = next(paragraph for paragraph in document.paragraphs if paragraph.text.strip().startswith("A. 选项A"))
            self.assertEqual(attachment_question.runs[0].font.name, "汉仪中宋简")
            self.assertEqual(attachment_option.runs[0].font.name, "汉仪中宋简")

            analysis_table_questions = []
            for table in document.tables:
                first_cell_text = table.cell(0, 0).text.strip()
                for question in expected_table_questions:
                    if first_cell_text == question:
                        analysis_table_questions.append(first_cell_text)
            self.assertEqual(len(analysis_table_questions), len(expected_table_questions))
            self.assertEqual(Counter(analysis_table_questions), Counter(expected_table_questions))
            self.assertFalse(any(re.match(r"^\d+(?:\.\d+)*\.", question) for question in analysis_table_questions))

            with ZipFile(output_docx) as zipped:
                xml = zipped.read("word/document.xml").decode("utf-8", "ignore")
                settings_xml = zipped.read("word/settings.xml").decode("utf-8", "ignore")
                chart1_xml = zipped.read("word/charts/chart1.xml").decode("utf-8", "ignore")
                chart2_xml = zipped.read("word/charts/chart2.xml").decode("utf-8", "ignore")
            self.assertIn('TOC \\o "1-3" \\h \\u', xml)
            self.assertNotIn("w:sdt", xml)
            self.assertIn("updateFields", settings_xml)
            self.assertIn('w:val="true"', settings_xml)
            self.assertIn("调研时间：2025年10月01日——2025年10月31日", xml)
            self.assertNotIn("2025年11月1日-11月30日", xml)
            self.assertIn("患者血压控制效果", chart1_xml)
            self.assertIn("药品价格接受度情况", chart2_xml)

    def test_render_from_template_does_not_fallback_to_first_visual_group_on_unmatched_table(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            report_content = Path(temp_dir) / "content.md"
            report_content.write_text(sample_markdown(), encoding="utf-8")
            meta, content = parse_markdown_content(report_content)
            payload = build_payload(
                efficacy_questionnaire(),
                meta,
                content,
                Namespace(
                    product=None,
                    region=None,
                    time=None,
                    attachment_name=None,
                    survey_period=None,
                    sample_size=None,
                    valid_count=None,
                    disclaimer_unit=None,
                ),
            )

            tampered_template = Path(temp_dir) / "tampered-template.docx"
            with ZipFile(Path(payload["meta"]["template_doc"])) as src_zip, ZipFile(tampered_template, "w") as dst_zip:
                for item in src_zip.infolist():
                    data = src_zip.read(item.filename)
                    if item.filename == "word/document.xml":
                        text = data.decode("utf-8", "ignore").replace(
                            "1.您服用厄贝沙坦氢氯噻嗪片后，血压控制效果如何？",
                            "X.您服用厄贝沙坦氢氯噻嗪片后，血压控制效果如何？",
                            1,
                        )
                        data = text.encode("utf-8")
                    dst_zip.writestr(item, data)

            output_docx = Path(temp_dir) / "rendered.docx"
            TemplateRenderer(tampered_template, payload).render(output_docx)

            document = Document(output_docx)
            expected_table_questions = [
                visual["table_data"]["question"]
                for section in payload["result_analysis"]["sections"]
                for visual in section["visual_groups"]
            ]
            analysis_table_questions = []
            for table in document.tables:
                first_cell_text = table.cell(0, 0).text.strip()
                for question in expected_table_questions:
                    if first_cell_text == question:
                        analysis_table_questions.append(first_cell_text)
            counts = Counter(analysis_table_questions)
            self.assertLessEqual(counts[payload["result_analysis"]["sections"][0]["visual_groups"][0]["table_data"]["question"]], 1)


if __name__ == "__main__":
    unittest.main()
