#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime
from pathlib import Path

try:
    from .build_payload import build_payload, parse_markdown_content, validate_payload
    from .parse_questionnaire import parse_sheet
    from .render_from_template import TemplateRenderer
except ImportError:
    from build_payload import build_payload, parse_markdown_content, validate_payload
    from parse_questionnaire import parse_sheet
    from render_from_template import TemplateRenderer


ROOT = Path(__file__).resolve().parents[1]


def slugify(text: str) -> str:
    cleaned = re.sub(r"[^\w\u4e00-\u9fff-]+", "-", text.strip(), flags=re.UNICODE)
    cleaned = re.sub(r"-{2,}", "-", cleaned).strip("-")
    return cleaned or "report"


def default_run_dir(questionnaire_path: Path, content_path: Path) -> Path:
    seed = f"{questionnaire_path.resolve()}|{content_path.resolve()}|{datetime.now().isoformat()}"
    digest = hashlib.sha1(seed.encode("utf-8")).hexdigest()[:8]
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    stem = slugify(questionnaire_path.stem)
    return ROOT / "tmp" / "runs" / f"{stamp}-{stem}-{digest}"


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the patient report pipeline in an isolated per-run directory.")
    parser.add_argument("questionnaire_xlsx")
    parser.add_argument("report_content")
    parser.add_argument("--run-dir")
    parser.add_argument("--output-docx")
    parser.add_argument("--product")
    parser.add_argument("--region")
    parser.add_argument("--time")
    parser.add_argument("--attachment-name")
    parser.add_argument("--survey-period")
    parser.add_argument("--sample-size")
    parser.add_argument("--valid-count")
    parser.add_argument("--disclaimer-unit")
    args = parser.parse_args()

    questionnaire_path = Path(args.questionnaire_xlsx)
    report_content_path = Path(args.report_content)
    run_dir = Path(args.run_dir) if args.run_dir else default_run_dir(questionnaire_path, report_content_path)
    run_dir.mkdir(parents=True, exist_ok=False)

    questionnaire = parse_sheet(questionnaire_path)
    questionnaire_json = run_dir / "questionnaire.json"
    questionnaire_json.write_text(json.dumps(questionnaire, ensure_ascii=False, indent=2), encoding="utf-8")

    meta, content = parse_markdown_content(report_content_path)
    payload = build_payload(questionnaire, meta, content, args)
    validate_payload(payload)

    payload_json = run_dir / "report_payload.json"
    payload_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    output_docx = Path(args.output_docx) if args.output_docx else run_dir / "report.docx"
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    TemplateRenderer(Path(payload["meta"]["template_doc"]), payload).render(output_docx)

    print(output_docx.resolve())


if __name__ == "__main__":
    main()
