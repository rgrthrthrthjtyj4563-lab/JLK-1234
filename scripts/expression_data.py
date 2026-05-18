#!/usr/bin/env python3
"""
Expression data extracted from expression-modules.md for programmatic use.

These are interchangeable building blocks for chapter 2 analysis paragraphs.
The default_subtopic_paragraph() function in build_payload.py randomly selects
one analysis angle and fills in actual option data.
"""

from __future__ import annotations

import random

# ─── Analysis Angles ─────────────────────────────────────────────────────────
# Each angle has: openings (list[str]), and angle-specific modules.

ANGLES = {
    "conclusion-and-action": {
        "openings": [
            "从本题选项占比看，当前患者反馈已经呈现出较明确的主导方向，可以据此得出……",
            "结合各选项占比结构，这一题反映出的核心结论比较清晰，重点在于……",
            "从问卷分布结果判断，该维度的患者状态已有较强指向性，说明……",
            "根据主要选项的占比高低，可以判断该问题当前最值得关注的结论是……",
            "本题的数据结构并不分散，说明患者在这一维度上的主要问题或主要优势已经较为明确。",
        ],
        "suggestions": [
            "因此，后续可优先从……入手，通过……这一类可执行方式进行改善或巩固。",
            "基于这一结果，更可行的做法是围绕……提供针对性支持，而不是泛泛增加信息量。",
            "从实际管理角度看，建议优先加强……，这样更有可能对应到当前占比最高的患者需求。",
            "若希望进一步改善这一维度，较现实的路径是通过……来回应当前主要反馈方向。",
            "围绕这一结论，后续可将患者教育、提醒或随访重点放在……，可操作性更强。",
        ],
    },
    "option-by-option-cause": {
        "openings": [
            "从各选项占比看，患者在本题上的选择并非偶然，不同选项背后对应着不同的现实原因。",
            "本题的分布结构提示，各主要选项之所以形成当前占比，与患者的认知、习惯和生活场景密切相关。",
            "若逐项看本题选项占比，可以发现每一类高占比选择背后都有相对清晰的行为逻辑。",
            "本题更适合逐项理解其占比形成原因，因为不同选项对应的并不是简单好坏，而是不同类型的患者状态。",
            "从选项分布本身出发，本题需要解释的重点不是谁高谁低，而是这些比例为什么会这样形成。",
        ],
        "cause_templates": [
            '选择"{opt_text}"的患者占{opt_pct}，这通常说明……',
            '"{opt_text}"对应{opt_pct}，往往反映出患者在……方面受到……影响。',
            '另一部分患者集中在"{opt_text}"（{opt_pct}），其背后更可能是……',
            '相比之下，"{opt_text}"占{opt_pct}，提示这类患者在……方面已有较稳定习惯或明确认知。',
            '从原因层面看，"{opt_text}"之所以形成当前比例，往往与……这一现实因素有关。',
        ],
    },
    "overall-structure-analysis": {
        "openings": [
            "从整体占比结构看，本题反映出的并不是单一意见，而是患者群体在这一维度上的整体状态分布。",
            "如果把本题各选项作为一个整体来看，可以更清楚地看到患者在该维度上的总体特征。",
            "本题更值得关注的是整体分布结构，因为它反映的是患者群体的共性状态，而不仅是某一个选项本身。",
            "从总体占比格局判断，本题呈现出的核心信息在于患者整体处于怎样的结构状态。",
            "各选项合并起来观察后，本题更像是在描述患者群体的整体管理面貌，而不是单点问题。",
        ],
        "structure_interpretations": [
            "这一结构说明，患者整体上更偏向……，但同时仍保留一定比例的……",
            '从群体层面看，该分布呈现出"……"的总体格局，说明……',
            "这一占比结构反映出患者并未完全走向单一模式，而是形成了以……为主、以……为辅的整体状态。",
            "综合各项比例后可以看出，当前患者群体在这一维度上的主流状态是……，而边缘问题主要集中在……",
            "从整体结构意义上说，本题显示出患者在……方面已经形成基本趋势，但距离更理想状态仍有差距。",
        ],
    },
}

# ─── Evidence Embedding Templates ────────────────────────────────────────────
# {opt_text} and {opt_pct} are replaced with actual data.
# For multi-option: {top_text}, {top_pct}, {second_text}, {second_pct}

EVIDENCE_TEMPLATES = {
    "single": [
        '其中，选择"{opt_text}"的患者占{opt_pct}，另有{second_pct}的患者倾向于"{second_text}"。',
        '具体来看，{opt_pct}的受访患者选择"{opt_text}"，{second_pct}选择"{second_text}"。',
        '在各项反馈中，占比相对更高的两类意见分别为"{opt_text}"（{opt_pct}）和"{second_text}"（{second_pct}）。',
        '从比例分布看，患者反馈主要集中在"{opt_text}"和"{second_text}"两类感受，分别占{opt_pct}和{second_pct}。',
        '就主流反馈而言，{opt_pct}的患者倾向于{opt_text}，同时{second_pct}的患者认为{second_text}。',
        '数据层面，占比最高的两个方向依次为"{opt_text}"（{opt_pct}）和"{second_text}"（{second_pct}）。',
        '按选项占比排序，前两位分别为"{opt_text}"（{opt_pct}）与"{second_text}"（{second_pct}），合计占比{combined_pct}。',
        '从百分比结构可观察到，"{opt_text}"以{opt_pct}居于首位，"{second_text}"以{second_pct}紧随其后。',
        '占比分布显示，{opt_text}和{second_text}两类反馈合计构成该维度判断的主体（共{combined_pct}）。',
        '在有效反馈中，{opt_pct}的患者选择{opt_text}，{second_pct}选择{second_text}，构成该维度的主要反馈走向。',
    ],
}

# ─── Closing Reminder Modules ────────────────────────────────────────────────

CLOSING_MODULES = [
    "日常用药过程中，仍建议结合自身感受、规律监测和医生指导持续关注这一维度。",
    "在后续管理中，可进一步通过提醒、宣教或随访支持提升该维度的执行效果。",
    "对于存在明显困扰或执行困难的患者，后续仍需在日常管理中加强关注和支持。",
    "该结果提示，患者端管理中仍需兼顾主流正向体验与少数场景下的实际痛点。",
    "综合来看，该维度的反馈为理解患者的真实管理状态提供了重要参考，后续可据此优化患者教育内容。",
    "从管理角度看，该维度的反馈结构支持将现有的正向模式推广到更多患者，同时为困难患者提供差异化支持。",
    "患者反馈的主流方向为日常管理提供了信心依据，但少数群体的困难声音同样值得在后续工作中重点回应。",
    "在实际工作中，可借助该维度已形成的患者经验开展同伴教育，同时为执行困难者建立更精准的帮扶路径。",
    "该维度虽有较清晰的反馈方向，但慢性病管理不是单一维度决策，需在整体自我管理框架中统筹考量。",
    "建议将本维度的正向经验转化为可操作的患者教育素材，将担忧和困难转化为后续支持服务的优先方向。",
]


def _format_evidence(options: list[dict]) -> str:
    """Format top 2 options into an evidence sentence using a random template."""
    if len(options) < 2:
        return ""
    top = options[0]
    second = options[1]
    combined = float(top["pct"].rstrip("%")) + float(second["pct"].rstrip("%"))
    template = random.choice(EVIDENCE_TEMPLATES["single"])
    return template.format(
        opt_text=f"{top['code']}.{top['text']}",
        opt_pct=top["pct"],
        second_text=f"{second['code']}.{second['text']}",
        second_pct=second["pct"],
        combined_pct=f"{combined:.1f}%",
    )


def generate_analysis_paragraphs(options: list[dict]) -> list[str]:
    """Generate 2-3 analysis paragraphs for a single question's options.

    Args:
        options: List of {code, text, count, pct} dicts, sorted by pct descending.

    Returns:
        List of paragraph strings (2-3 items).
    """
    if not options:
        return ["原始问卷未提供足够的选项数据，无法形成有效统计解释。"]

    # Sort by percentage descending
    sorted_opts = sorted(
        options,
        key=lambda o: float(o.get("pct", "0").rstrip("%")),
        reverse=True,
    )

    # Randomly pick an analysis angle
    angle_name, angle = random.choice(list(ANGLES.items()))
    opening = random.choice(angle["openings"])
    # Strip trailing "……" from opening — evidence will provide concrete data
    opening = opening.rstrip("……")

    # Helper: fill "……" placeholders with generic completions
    def _fill_placeholder(text: str) -> str:
        """Replace "……" with data-based completions where possible."""
        text = text.replace("……", "……")
        return text

    paragraphs = []

    if angle_name == "conclusion-and-action":
        # P1: opening with evidence
        evidence = _format_evidence(sorted_opts)
        p1_parts = [opening]
        if evidence:
            p1_parts.append(evidence)
        paragraphs.append("".join(p1_parts))

        # P2: suggestion — replace "……" with top option reference
        suggestion = random.choice(angle["suggestions"])
        top_opt = sorted_opts[0]
        top_text = f'{top_opt["code"]}.{top_opt["text"]}'
        suggestion = suggestion.replace("……", top_text, 1)
        # Replace remaining "……" with generic fill
        while "……" in suggestion:
            suggestion = suggestion.replace("……", "该方向", 1)
        paragraphs.append(suggestion)

    elif angle_name == "option-by-option-cause":
        # P1: opening
        paragraphs.append(opening)

        # P2: evidence embedding + cause analysis
        evidence = _format_evidence(sorted_opts)
        cause_templates = random.sample(angle["cause_templates"], min(2, len(sorted_opts)))
        cause_lines = []
        for i, tmpl in enumerate(cause_templates):
            if i >= len(sorted_opts):
                break
            opt = sorted_opts[i]
            line = tmpl.format(
                opt_text=f'{opt["code"]}.{opt["text"]}',
                opt_pct=opt["pct"],
            )
            # Fill remaining "……" placeholders
            while "……" in line:
                line = line.replace("……", f'{opt["code"]}.{opt["text"]}')
            cause_lines.append(line)
        combined = evidence + "".join(cause_lines) if evidence else "".join(cause_lines)
        paragraphs.append(combined)

    elif angle_name == "overall-structure-analysis":
        # P1: opening with evidence
        evidence = _format_evidence(sorted_opts)
        p1 = opening
        if evidence:
            p1 += evidence
        paragraphs.append(p1)

        # P2: structure interpretation — fill placeholders with top options
        interpretation = random.choice(angle["structure_interpretations"])
        top_opt = sorted_opts[0]
        top_text = f'{top_opt["code"]}.{top_opt["text"]}'
        while "……" in interpretation:
            interpretation = interpretation.replace("……", top_text, 1)
        paragraphs.append(interpretation)

    # Closing reminder
    closing = random.choice(CLOSING_MODULES)
    paragraphs.append(closing)

    return paragraphs
