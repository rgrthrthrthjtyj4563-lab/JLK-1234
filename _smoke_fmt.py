"""End-to-end smoke test for the template renderer."""

from __future__ import annotations

import tempfile
from pathlib import Path

from docx import Document

from scripts.build_payload import build_payload, parse_markdown_content
from scripts.render_from_template import TemplateRenderer
from tests.test_jlk_pipeline import Namespace, efficacy_questionnaire, sample_markdown


def main() -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        report_content = temp_dir / "content.md"
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

        output_docx = temp_dir / "smoke.docx"
        TemplateRenderer(Path(payload["meta"]["template_doc"]), payload).render(output_docx)

        document = Document(str(output_docx))
        texts = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
        required = {"问卷调研服务结算", "目录", "前言", "项目背景", "问卷结果分析", "调研结果", "免责申明"}
        missing = sorted(required.difference(texts))
        if missing:
            raise AssertionError(f"Missing required template headings: {missing}")

        if len(document.sections) != 5:
            raise AssertionError(f"Expected 5 sections, got {len(document.sections)}")

    print("template-render smoke test: OK")


if __name__ == "__main__":
    main()
