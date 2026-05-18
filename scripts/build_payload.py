#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    from .cluster_dimensions import cluster_dimensions
except ImportError:
    from cluster_dimensions import cluster_dimensions


def normalize_space(text: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", str(text).strip())


def normalize_pct(text: str) -> str:
    value = str(text).strip()
    if not value:
        return ""
    if value.endswith("%"):
        value = value[:-1].strip()
    return f"{float(value):.2f}%"


def split_front_matter(text: str) -> tuple[dict, str]:
    stripped = text.lstrip("\ufeff")
    match = re.match(r"^\s*---\s*\n(.*?)\n\s*---\s*\n?(.*)$", stripped, flags=re.DOTALL)
    if not match:
        return {}, stripped
    meta_block = match.group(1)
    body = match.group(2)
    meta = {}
    for line in meta_block.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = value.strip().strip('"').strip("'")
    return meta, body


def parse_paragraphs(lines: list[str]) -> list[str]:
    paragraphs = []
    buf: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if buf:
                paragraphs.append(normalize_space(" ".join(buf)))
                buf = []
            continue
        buf.append(stripped)
    if buf:
        paragraphs.append(normalize_space(" ".join(buf)))
    return paragraphs


def strip_markdown_heading_prefix(text: str) -> str:
    return re.sub(r"^\s{0,3}#{1,6}\s*", "", str(text).strip())


def parse_markdown_content(path: Path) -> tuple[dict, dict]:
    meta, body = split_front_matter(path.read_text(encoding="utf-8"))
    lines = body.splitlines()
    current_section: str | None = None
    current_result_section: dict | None = None
    current_subtopic: dict | None = None

    sections = {
        "preface": [],
        "project_background": [],
        "project_execution": [],
        "questionnaire_note": [],
        "result_analysis_intro": [],
        "result_analysis": [],
        "key_issue_analysis": [],
        "overall_analysis": [],
        "recommendations": [],
    }

    def flush_subtopic() -> None:
        nonlocal current_subtopic, current_result_section
        if current_subtopic and current_result_section:
            current_subtopic["paragraphs"] = parse_paragraphs(current_subtopic["lines"])
            del current_subtopic["lines"]
            current_result_section["subtopics"].append(current_subtopic)
        current_subtopic = None

    def flush_result_section() -> None:
        nonlocal current_result_section
        flush_subtopic()
        if current_result_section:
            current_result_section["intro"] = parse_paragraphs(current_result_section["lines"])
            del current_result_section["lines"]
            sections["result_analysis"].append(current_result_section)
        current_result_section = None

    for line in lines:
        heading = re.match(r"^(#{2,5})\s+(.+?)\s*$", line.lstrip())
        if heading:
            level = len(heading.group(1))
            title = normalize_space(heading.group(2))
            if title in {"前言"}:
                flush_result_section()
                current_section = "preface"
                continue
            if title in {"项目背景", "一、项目背景"}:
                flush_result_section()
                current_section = "project_background"
                continue
            if title in {"项目开展情况", "二、项目开展情况"}:
                flush_result_section()
                current_section = "project_execution"
                continue
            if title in {"问卷说明", "三、问卷说明"}:
                flush_result_section()
                current_section = "questionnaire_note"
                continue
            if title in {"问卷结果分析", "四、问卷结果分析"}:
                flush_result_section()
                current_section = "result_analysis_intro"
                continue
            if title in {"调研结果", "五、调研结果"}:
                flush_result_section()
                current_section = "summary"
                continue
            if level <= 3 and re.match(r"4\.\d+", title):
                flush_result_section()
                current_section = "result_analysis"
                # Extract section_number from title (e.g. "4.1 药品疗效" → "4.1")
                sn_match = re.match(r"(4\.\d+)\s*", title)
                section_number = sn_match.group(1) if sn_match else title
                section_title = title[len(section_number):].strip()
                current_result_section = {
                    "title": title,
                    "section_number": section_number,
                    "section_title": section_title,
                    "lines": [],
                    "subtopics": [],
                }
                continue
            if level >= 4 and current_result_section:
                flush_subtopic()
                current_section = "result_analysis"
                current_subtopic = {"title": title, "lines": []}
                continue
            if re.match(r"5\.1", title):
                flush_result_section()
                current_section = "key_issue_analysis"
                continue
            if re.match(r"5\.2", title):
                flush_result_section()
                current_section = "overall_analysis"
                continue
            if re.match(r"5\.3", title):
                flush_result_section()
                current_section = "recommendations"
                continue

        if current_section in {"preface", "project_background", "project_execution", "questionnaire_note", "result_analysis_intro", "key_issue_analysis", "overall_analysis", "recommendations"}:
            sections[current_section].append(strip_markdown_heading_prefix(line))
        elif current_result_section and current_subtopic:
            current_subtopic["lines"].append(strip_markdown_heading_prefix(line))
        elif current_result_section:
            current_result_section["lines"].append(strip_markdown_heading_prefix(line))

    flush_result_section()
    return meta, {
        "preface": parse_paragraphs(sections["preface"]),
        "project_background": parse_paragraphs(sections["project_background"]),
        "project_execution": parse_paragraphs(sections["project_execution"]),
        "questionnaire_note": parse_paragraphs(sections["questionnaire_note"]),
        "result_analysis_intro": parse_paragraphs(sections["result_analysis_intro"]),
        "result_analysis": sections["result_analysis"],
        "summary": {
            "key_issue_analysis": parse_paragraphs(sections["key_issue_analysis"]),
            "overall_analysis": parse_paragraphs(sections["overall_analysis"]),
            "recommendations": parse_paragraphs(sections["recommendations"]),
        },
    }


def build_attachment_questions(questionnaire: dict) -> list[dict]:
    items = []
    for question in questionnaire.get("questions", []):
        options = []
        for option in question.get("options", []):
            code = str(option["label"]).strip()
            text = normalize_space(option["text"])
            if not text or text == code:
                continue
            options.append(
                {
                    "code": code,
                    "text": text,
                    "count": str(option.get("count") or "").strip(),
                    "pct": normalize_pct(option["pct"]),
                }
            )
        items.append(
            {
                "question_ref": f"q{int(question['number']):02d}",
                "number": int(question["number"]),
                "question": normalize_space(question["question"]),
                "options": options,
            }
        )
    return items


def first_sample_size(questionnaire: dict) -> str | None:
    totals = [str(item.get("total")).strip() for item in questionnaire.get("questions", []) if str(item.get("total", "")).strip()]
    return totals[0] if totals and len(set(totals)) == 1 else None


def format_today() -> str:
    today = date.today()
    return f"{today.year}年{today.month:02d}月{today.day:02d}日"


def normalize_survey_period(text: str) -> str:
    cleaned = normalize_space(text).replace("—", "-").replace("–", "-").replace("至", "-").replace("~", "-")
    parts = [part.strip() for part in re.split(r"\s*-\s*", cleaned, maxsplit=1) if part.strip()]
    if len(parts) != 2:
        raise ValueError(f"Invalid survey_period: {text}")

    def _parse_fragment(fragment: str, default_year: int | None = None, default_month: int | None = None) -> date:
        numbers = [int(x) for x in re.findall(r"\d+", fragment)]
        if len(numbers) == 3:
            year, month, day = numbers
        elif len(numbers) == 2 and default_year is not None:
            year, month, day = default_year, numbers[0], numbers[1]
        elif len(numbers) == 1 and default_year is not None and default_month is not None:
            year, month, day = default_year, default_month, numbers[0]
        else:
            raise ValueError(f"Invalid survey_period: {text}")
        return date(year, month, day)

    start = _parse_fragment(parts[0])
    end = _parse_fragment(parts[1], default_year=start.year, default_month=start.month)
    return f"{start.year}年{start.month:02d}月{start.day:02d}日——{end.year}年{end.month:02d}月{end.day:02d}日"


def derive_service_date(survey_period: str) -> str:
    normalized = normalize_survey_period(survey_period)
    numbers = [int(x) for x in re.findall(r"\d+", normalized)]
    end = date(numbers[3], numbers[4], numbers[5])
    next_month = date(end.year + (1 if end.month == 12 else 0), 1 if end.month == 12 else end.month + 1, 1)
    after_next_month = date(next_month.year + (1 if next_month.month == 12 else 0), 1 if next_month.month == 12 else next_month.month + 1, 1)
    last_day = after_next_month - timedelta(days=1)
    day_span = (last_day - next_month).days
    offset = min(day_span, max(0, (end.day * 3 + end.month) % (day_span + 1)))
    service_dt = next_month + timedelta(days=offset)
    return service_dt.strftime("%Y年%m月%d日")


def as_single_paragraph(content: list[str], fallback: str) -> list[str]:
    chosen = content if content else [fallback]
    merged = normalize_space(" ".join(strip_markdown_heading_prefix(item) for item in chosen if str(item).strip()))
    return [merged] if merged else [fallback]


def choose_paragraphs(content: list[str], fallback: list[str]) -> list[str]:
    normalized = [normalize_space(strip_markdown_heading_prefix(item)) for item in content if str(item).strip()]
    return normalized if normalized else fallback


def sanitize_body_paragraphs(content: list[str]) -> list[str]:
    blocked_patterns = [
        r"^(?:前言|项目背景|项目开展情况|问卷说明|问卷结果分析|调研结果|免责申明)\s*$",
        r"^(?:4|5)\.\d+",
        r"^附件\d+",
    ]
    sanitized = []
    for item in content:
        text = normalize_space(strip_markdown_heading_prefix(item))
        if not text:
            continue
        if any(re.match(pattern, text) for pattern in blocked_patterns):
            continue
        sanitized.append(text)
    return sanitized


def choose_body_paragraphs(content: list[str], fallback: list[str]) -> list[str]:
    normalized = sanitize_body_paragraphs(content)
    return normalized if normalized else fallback


def build_project_execution(
    region: str,
    sample_size: str | None,
    survey_period: str,
    question_count: int,
    project_dimensions: list[str],
) -> dict:
    dimensions_text = "、".join(project_dimensions)
    lines = [
        f"项目工具：调查问卷，{question_count}道选择题。其内容涵盖{dimensions_text}{len(project_dimensions)}大维度，全面覆盖患者用药全流程关键节点。",
        f"样本采集范围：{region}",
        f"样本采集数量：本次共收集筛选到有效问卷{sample_size or '问卷未提供'}份。",
        f"样本采集时间：{survey_period}",
    ]
    return {"lines": lines}


def build_questionnaire_note(question_count: int) -> dict:
    intro = "为提升数据可比性与分析结果的科学性，本研究采用以下标准化统计处理方法对原始问卷数据进行系统规整："
    items = [
        "1．数据清洗：剔除无效作答、空白选项及关键信息缺失数据，保障分析数据集的完整性与可靠性。",
        "2．数据转换：按照问卷分析需求对原始选项、样本量和占比进行统一格式化处理，使其适配后续统计逻辑。",
        "3．数据聚合：根据题目主题和分析维度对数据进行归类整合，形成结构化分析单元，降低交叉解读难度。",
        "4．数据可视化：针对重点问题配置结果表格与对应图表，以直观呈现样本分布特征及不同选项之间的差异。",
    ]
    closing = f"通过上述标准化数据处理流程，可在{question_count}道题目的分析中有效强化不同维度数据间的可比性，确保研究结论具备更好的客观性与可信度。"
    return {"intro": intro, "items": items, "closing": closing}


def derive_report_title(product: str, grouped: dict) -> str:
    return "问卷调研分析报告"


def derive_project_background(product: str, region: str = "") -> str:
    return (
        f"高血压作为常见慢性心血管疾病，长期规范用药是控制病情、降低并发症风险的关键措施。"
        f"{product}作为临床常用复方降压药物，兼具稳定降压与改善长期管理体验的应用价值，"
        f"在{region or '目标地区'}患者群体中具有持续且稳定的使用基础。围绕患者真实用药场景开展本次调研，"
        "有助于从主观疗效感知、安全性体验、行为习惯和信息支持等层面系统识别患者反馈，"
        "为后续优化临床沟通、完善患者教育内容以及提升药品服务支持提供更具针对性的依据。"
    )


def derive_preface(product: str, region: str = "", sample_size: str = "") -> str:
    return (
        f"{product}作为临床常用复方降压药物，在高血压患者长期管理中兼具明确的疗效基础与较好的耐受性表现，"
        f"其真实用药体验对提升慢病管理质量具有重要参考价值。本次调研面向{region or '目标地区'}患者群体开展，"
        f"共收集有效问卷{sample_size or '若干'}份，重点围绕患者在疗效感知、安全性体验、用药行为、便利性和信息支持等方面的实际反馈进行整理分析，"
        "旨在从患者视角进一步识别当前用药过程中值得巩固的积极体验与仍待优化的关键环节，为后续临床沟通、患者教育和药品服务支持提供依据。"
    )


def question_map_by_ref(questionnaire: dict) -> dict:
    mapping = {}
    for question in build_attachment_questions(questionnaire):
        mapping[question["question_ref"]] = question
    return mapping


def default_subtopic_paragraph(question: dict) -> list[str]:
    """Generate analysis paragraphs using expression modules from expression_data.py."""
    try:
        from .expression_data import generate_analysis_paragraphs
    except ImportError:
        from expression_data import generate_analysis_paragraphs

    options = question.get("options", [])
    if not options:
        return ["原始问卷未提供足够的选项数据，无法形成有效统计解释。"]

    return generate_analysis_paragraphs(options)


def build_key_issue_analysis(grouped: dict, section_titles: list[str], region: str, product: str) -> list[str]:
    dim_text = "、".join(title for title in section_titles if title)
    return [
        f"本次问卷重点问题主要集中在{dim_text}等维度，整体反馈呈现出积极体验较为集中、个别环节仍有优化空间的特点。",
        f"结合{region or '目标地区'}患者的整体反馈来看，{product}在多数核心维度上已经形成较稳定的正向感知，但在部分行为管理与信息支持环节仍需持续强化沟通和引导。",
    ]


def key_issue_heading(question: dict) -> str:
    text = question.get("question", "")
    rules = [
        (r"血压控制|控压", "1. 血压控制现状与特征"),
        (r"价格|负担|承受", "2. 药品价格接受度情况"),
        (r"剂量|医嘱", "2. 剂量执行规范性表现"),
        (r"监测血压|监测频率", "2. 血压监测行为特征"),
    ]
    for pattern, heading in rules:
        if re.search(pattern, text):
            return heading
    return "重点问题分析"


def key_issue_chart_title(question: dict) -> str:
    text = question.get("question", "")
    rules = [
        (r"血压控制|控压", "患者血压控制效果"),
        (r"价格|负担|承受", "药品价格接受度情况"),
        (r"剂量|医嘱", "患者剂量遵循情况"),
        (r"监测血压|监测频率", "患者血压监测频率"),
    ]
    for pattern, title in rules:
        if re.search(pattern, text):
            return title
    return question.get("question", "重点问题")


def build_key_issue_paragraph(question: dict, region: str, product: str) -> str:
    options = sorted(
        question.get("options", []),
        key=lambda item: float(str(item.get("pct", "0")).rstrip("%") or 0),
        reverse=True,
    )
    if not options:
        return f"围绕{product}在{region or '目标地区'}患者中的这一重点问题，目前问卷样本不足以支撑稳定结论，建议后续结合补充样本继续观察。"

    lead = options[0]
    second = options[1] if len(options) > 1 else None
    tail = options[-1]
    question_text = question.get("question", "")

    if re.search(r"血压控制|控压", question_text):
        return normalize_space(
            f"在{region or '目标地区'}患者样本中，围绕{product}使用后的血压控制表现，患者反馈主要集中在“{lead.get('text', '')}”"
            f"{'和“' + second.get('text', '') + '”' if second else ''}等正向感知选项，说明多数患者能够较清晰地感受到产品在长期控压中的稳定价值。"
            f"这意味着产品的疗效体验并非停留在抽象认知层面，而是已经在真实用药过程中被患者转化为可感知、可复述的主观评价。"
            f"与此同时，仍有少量反馈落在“{tail.get('text', '')}”等相对谨慎的选项上，提示个别患者在血压管理结果上仍存在波动感受。"
            "从重点问题角度看，这一题反映出的核心并不是产品是否具备降压价值，而是后续如何通过复诊沟通、监测教育与规范用药提醒，让已经形成的正向疗效体验进一步稳定下来，并减少少数患者在血压波动阶段的疑虑。"
        )

    if re.search(r"价格|负担|承受", question_text):
        return normalize_space(
            f"在价格接受度这一重点问题上，患者反馈主要集中在“{lead.get('text', '')}”"
            f"{'和“' + second.get('text', '') + '”' if second else ''}等可接受区间，说明{product}在{region or '当前区域'}的长期使用成本并未普遍构成障碍。"
            "这类结果的意义在于，患者对产品价值的认可并没有因为价格而被明显削弱，长期治疗的连续性因此具备较好的现实基础。"
            f"同时，仍有少量患者落在“{tail.get('text', '')}”等相对保守的选项上，提示不同患者在长期负担感知上仍存在差异。"
            "对于慢病用药而言，价格不是单独存在的判断，而会直接影响患者是否愿意持续购药、是否能够长期坚持治疗。"
            "因此，这一重点题目所反映出的结论不仅是“价格总体可接受”，更是产品在长期管理场景中具备较强的经济可持续性，为稳定依从行为提供了重要支撑。"
        )

    return normalize_space(
        f"围绕“{question.get('question', '')}”这一重点问题，患者反馈主要集中在“{lead.get('text', '')}”"
        f"{'和“' + second.get('text', '') + '”' if second else ''}等主流选项上，说明该问题已经能够形成较明确的患者判断。"
        f"少量反馈落在“{tail.get('text', '')}”等相对谨慎的选项上，提示仍需针对这一环节继续加强说明与管理支持。"
        f"从{region or '当前区域'}样本看，这一重点问题更多体现的是患者在真实使用场景中的长期执行差异，而非单纯的一次性态度表达，后续应继续结合随访与场景化教育强化改善。"
    )


def select_key_issue_questions(result_sections: list[dict], qmap: dict, preferred_section_numbers: list[str]) -> list[dict]:
    selected = []
    used_refs = set()
    for section_number in preferred_section_numbers:
        section = next((item for item in result_sections if item.get("section_number") == section_number), None)
        if not section:
            continue
        for subtopic in section.get("subtopics", []):
            refs = subtopic.get("question_refs", [])
            if refs and refs[0] in qmap and refs[0] not in used_refs:
                question = qmap[refs[0]]
                selected.append(
                    {
                        "question_ref": refs[0],
                        "heading": key_issue_heading(question),
                        "chart_title": key_issue_chart_title(question),
                        "question": question,
                    }
                )
                used_refs.add(refs[0])
                break
        if len(selected) >= 2:
            break
    return selected[:2]


def validate_content_structure(content: dict, grouped: dict) -> None:
    expected_sections = grouped.get("sections", [])
    expected_by_number = {section["section_number"]: section for section in expected_sections}
    expected_numbers = [section["section_number"] for section in expected_sections]
    drafted_sections = content.get("result_analysis", [])
    if not drafted_sections:
        return

    drafted_numbers = [section.get("section_number", "") for section in drafted_sections]
    if any(number not in expected_by_number for number in drafted_numbers):
        raise ValueError("AI draft contains result-analysis sections outside the detected template structure.")

    expected_order = [number for number in expected_numbers if number in drafted_numbers]
    if drafted_numbers != expected_order:
        raise ValueError("AI draft result-analysis section order conflicts with the detected template structure.")

    for section in drafted_sections:
        expected = expected_by_number[section["section_number"]]
        drafted_title = normalize_space(section.get("section_title", ""))
        if drafted_title and drafted_title != expected["section_title"]:
            raise ValueError(
                f"AI draft section title mismatch for {section['section_number']}: "
                f"expected {expected['section_title']}, got {drafted_title}."
            )

        drafted_subtopics = section.get("subtopics", [])
        expected_subtopics = expected.get("subtopics", [])
        if len(drafted_subtopics) > len(expected_subtopics):
            raise ValueError(
                f"AI draft subtopic count exceeds detected template slots for {section['section_number']}."
            )
        for index, drafted_subtopic in enumerate(drafted_subtopics):
            drafted_subtopic_title = normalize_space(drafted_subtopic.get("title", ""))
            if drafted_subtopic_title and drafted_subtopic_title != expected_subtopics[index]["subtitle"]:
                raise ValueError(
                    f"AI draft subtitle mismatch for {section['section_number']} slot {index + 1}: "
                    f"expected {expected_subtopics[index]['subtitle']}, got {drafted_subtopic_title}."
                )


def build_programmatic_recommendations(product: str, region: str) -> list[str]:
    return [
        normalize_space(
            f"围绕{product}在{region or '当前调研区域'}患者中的真实使用反馈，建议持续强化规律服药、血压监测和复诊沟通等基础教育内容，帮助患者将已有的正向体验稳定转化为长期管理习惯。"
        ),
        normalize_space(
            "针对联合用药确认、信息理解和少数执行薄弱环节，可通过药师提示、随访提醒和场景化问答材料补强关键知识点，减少患者在日常决策中的理解偏差。"
        ),
        normalize_space(
            "后续患者支持可继续围绕重点问题开展分层服务，对反馈积极的维度沉淀正向案例，对存在波动的维度配置更有针对性的提醒与解释机制。"
        ),
    ]


def _question_option_snapshot(question: dict) -> str:
    options = question.get("options", [])
    if not options:
        return "问卷原始分布不足，需结合后续回访进一步确认。"
    ranked = sorted(
        options,
        key=lambda item: float(str(item.get("pct", "0")).rstrip("%") or 0),
        reverse=True,
    )
    lead = ranked[0]
    follow = ranked[1] if len(ranked) > 1 else None
    if follow:
        return f"反馈主要集中在“{lead.get('text', '')}”与“{follow.get('text', '')}”两类选项"
    return f"反馈主要集中在“{lead.get('text', '')}”这一选项"


def build_programmatic_overall_analysis(
    result_sections: list[dict],
    qmap: dict,
    region: str,
    product: str,
    sample_size: str | None,
) -> list[str]:
    section_lookup = {section["section_number"]: section for section in result_sections}

    def first_snapshot(section_number: str) -> str:
        section = section_lookup.get(section_number, {})
        for subtopic in section.get("subtopics", []):
            refs = subtopic.get("question_refs", [])
            if refs and refs[0] in qmap:
                return _question_option_snapshot(qmap[refs[0]])
        return "当前维度的问卷反馈能够形成相对稳定的结论，但仍需结合后续随访持续观察。"

    paragraphs = [
        (
            f"本次调研基于{region or '目标地区'}患者样本开展，结合{sample_size or '当前收集到的有效问卷'}份有效问卷结果，"
            f"对{product}在真实使用过程中的疗效感知、安全性体验、用药行为、便利性、经济性、可及性以及信息支持情况进行了复盘。"
            "整体来看，患者反馈以正向、稳定为主，说明产品在长期慢病管理场景中已经形成较成熟的使用基础，但行为执行和信息确认环节仍存在继续优化空间。"
        ),
        (
            f"在疗效与安全性层面，{first_snapshot('4.1')}，说明患者对血压控制结果的感知总体较积极。"
            f"同时，{first_snapshot('4.2')}，意味着患者并未将常见不适反应视为主要阻碍，产品在真实使用场景中的耐受性基础较好。"
        ),
        (
            f"在用药行为与便利性层面，{first_snapshot('4.3')}，说明多数患者已经具备一定的规范意识，但在联合用药确认、监测频率和主动咨询等环节上仍可能存在执行松动。"
            f"在便利性方面，{first_snapshot('4.4')}，反映出服药频率与包装设计整体上没有构成普遍阻力。"
        ),
        (
            f"在经济性、可及性与信息支持层面，{first_snapshot('4.5')}，说明价格感知并未普遍构成持续治疗障碍；{first_snapshot('4.6')}，表明供应稳定性总体可控；{first_snapshot('4.7')}，则显示现有说明书与指导支持基本能够覆盖患者的基础理解需求。"
            f"综合判断，{product}在{region or '当前调研区域'}已经具备较完整的正向使用基础，后续重点应放在加强行为执行、提升细分解释能力，以及补充联合用药和监测管理支持等服务动作上。"
        ),
    ]
    return [normalize_space(paragraph) for paragraph in paragraphs]


def choose_overall_analysis(ai_paragraphs: list[str], fallback: list[str]) -> list[str]:
    normalized = choose_body_paragraphs(ai_paragraphs, [])
    total_len = sum(len(paragraph) for paragraph in normalized)
    if len(normalized) >= 4 and total_len <= 700:
        return normalized
    return fallback


def build_payload(questionnaire: dict, meta: dict, content: dict, cli_args: argparse.Namespace) -> dict:
    grouped = cluster_dimensions(questionnaire)
    validate_content_structure(content, grouped)
    attachment_questions = build_attachment_questions(questionnaire)
    qmap = {item["question_ref"]: item for item in attachment_questions}
    sample_size = cli_args.sample_size or first_sample_size(questionnaire)
    survey_period_raw = cli_args.survey_period or meta.get("survey_period")
    if not survey_period_raw:
        raise ValueError("survey_period is required.")
    survey_period = normalize_survey_period(survey_period_raw)
    product = cli_args.product or meta.get("product") or meta.get("品种")
    region = cli_args.region or meta.get("region") or meta.get("地区")
    question_count = questionnaire.get("question_count", len(questionnaire.get("questions", [])))
    report_title = derive_report_title(product, grouped)

    ai_sections = content.get("result_analysis", [])
    ai_by_sn = {}
    for item in ai_sections:
        sn = item.get("section_number")
        if sn:
            ai_by_sn[sn] = item

    dimension_names = grouped.get("project_dimensions", [])

    # Build result_intro from template
    result_intro_template = grouped.get("result_intro_template", grouped.get("result_intro", ""))
    result_intro = result_intro_template.format(
        question_count=question_count,
        dimensions_list="、".join(dimension_names),
        dimension_count=len(dimension_names),
    )

    result_sections = []
    chart_map = {}
    for planned in grouped["sections"]:
        section_number = planned["section_number"]
        drafted = ai_by_sn.get(section_number, {})

        section_title = planned["section_title"]
        section_intro = [planned["section_intro"]]

        ai_subtopics = drafted.get("subtopics", [])
        ai_st_by_idx = {}
        for idx, st in enumerate(ai_subtopics):
            ai_st_by_idx[idx] = st

        subtopics = []
        for st in planned["subtopics"]:
            refs = st["question_refs"]
            st_idx = st["subtopic_index"]
            ai_st = ai_st_by_idx.get(st_idx, {})

            subtitle = st["subtitle"]
            analysis = choose_body_paragraphs(
                ai_st.get("paragraphs", []),
                default_subtopic_paragraph(qmap[refs[0]]) if refs else [],
            )

            subtopics.append({
                "subtitle": subtitle,
                "question_refs": refs,
                "paragraphs": analysis,
            })

        # Visual groups
        visuals = []
        for visual in planned["visual_groups"]:
            question = qmap[visual["question_ref"]]
            chart_map[visual["chart_ref"]] = question
            visuals.append({
                "question_ref": visual["question_ref"],
                "chart_ref": visual["chart_ref"],
                "chart_type": visual["chart_type"],
                "chart_style_profile": visual["chart_style_profile"],
                "table_data": question,
            })

        result_sections.append({
            "section_number": section_number,
            "section_title": section_title,
            "section_intro": section_intro if section_intro else [],
            "visual_groups": visuals,
            "subtopics": subtopics,
        })

    key_issue_questions = select_key_issue_questions(
        result_sections,
        qmap,
        grouped.get("key_issue_preferred_sections", []),
    )
    key_issue_items = [
        {
            "question_ref": item["question_ref"],
            "heading": item["heading"],
            "chart_title": item["chart_title"],
            "paragraph": build_key_issue_paragraph(item["question"], region, product),
        }
        for item in key_issue_questions
    ]

    programmatic_overall_analysis = build_programmatic_overall_analysis(
        result_sections,
        qmap,
        region,
        product,
        str(sample_size) if sample_size else None,
    )

    return {
        "meta": {
            "product": product,
            "region": region,
            "time": cli_args.time or meta.get("time") or meta.get("时间"),
            "survey_period": survey_period,
            "valid_count": cli_args.valid_count or meta.get("valid_count") or sample_size,
            "sample_size": sample_size,
            "template_type": grouped["template_type"],
            "template_doc": grouped["template_doc"],
            "report_date": format_today(),
        },
        "red_font_replacements": {
            "drug_name": product,
            "region": region,
            "sample_size": str(sample_size) if sample_size else "",
            "survey_period": survey_period or "",
            "report_date": format_today(),
        },
        "header_text": report_title,
        "report_title": report_title,
        "service": {
            "unit": cli_args.disclaimer_unit or meta.get("disclaimer_unit") or "北京玖麟空科技有限公司",
            "date": derive_service_date(survey_period),
        },
        "preface": as_single_paragraph(content["preface"], derive_preface(product, region, str(sample_size) if sample_size else "")),
        "project_background": as_single_paragraph(content["project_background"], derive_project_background(product, region)),
        "project_execution": build_project_execution(
            region,
            sample_size,
            survey_period,
            question_count,
            dimension_names,
        ),
        "questionnaire_note": build_questionnaire_note(question_count),
        "result_analysis": {
            "intro": [result_intro],
            "sections": result_sections,
        },
        "summary": {
            "key_issue_analysis": [
                value
                for item in key_issue_items
                for value in [item["heading"], item["paragraph"]]
            ],
            "key_issue_analysis_programmatic": [
                value
                for item in key_issue_items
                for value in [item["heading"], item["paragraph"]]
            ],
            "key_issue_items": key_issue_items,
            "overall_analysis": choose_overall_analysis(content["summary"]["overall_analysis"], programmatic_overall_analysis),
            "overall_analysis_programmatic": programmatic_overall_analysis,
            "recommendations": build_programmatic_recommendations(product, region),
        },
        "attachments": {
            "attachment1_name": cli_args.attachment_name or meta.get("attachment_name") or "问卷调查附件",
            "attachment1_questions": attachment_questions,
            "attachment2_name": "问卷调查明细表",
        },
        "disclaimer": {
            "title": "免责申明",
            "items": [
                "（1）本次调研项目以随机选取对象进行面对面调研，本次调研只对本次样本数据负责。",
                f"（2）承接单位调研项目，是针对{product}这一品种调研，并非指定厂家指定品种。",
                "（3）本次调研只针对调研区域数据负责，不代表全国调研数据。",
            ],
            "unit": cli_args.disclaimer_unit or meta.get("disclaimer_unit") or "北京玖麟空科技有限公司",
            "date": derive_service_date(survey_period),
        },
        "chart_map": chart_map,
    }


def validate_payload(payload: dict) -> None:
    for key in ["meta", "header_text", "report_title", "preface", "project_background", "project_execution", "questionnaire_note", "result_analysis", "summary", "attachments", "disclaimer"]:
        if key not in payload:
            raise ValueError(f"Missing payload key: {key}")
    if not payload["preface"]:
        raise ValueError("Preface must not be empty.")
    if not payload["result_analysis"]["sections"]:
        raise ValueError("Result analysis sections must not be empty.")
    for section in payload["result_analysis"]["sections"]:
        if not section["subtopics"]:
            raise ValueError(f"Section {section['section_number']}{section['section_title']} must contain subtopics.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build template-aligned patient report payload.")
    parser.add_argument("questionnaire_json")
    parser.add_argument("report_content")
    parser.add_argument("-o", "--output", required=True)
    parser.add_argument("--product")
    parser.add_argument("--region")
    parser.add_argument("--time")
    parser.add_argument("--attachment-name")
    parser.add_argument("--survey-period")
    parser.add_argument("--sample-size")
    parser.add_argument("--valid-count")
    parser.add_argument("--disclaimer-unit")
    args = parser.parse_args()

    questionnaire = json.loads(Path(args.questionnaire_json).read_text(encoding="utf-8"))
    meta, content = parse_markdown_content(Path(args.report_content))
    payload = build_payload(questionnaire, meta, content, args)
    validate_payload(payload)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
