# Report Payload Schema

```json
{
  "meta": {
    "product": "厄贝沙坦氢氯噻嗪片",
    "region": "北京市",
    "time": "2025.10",
    "template_type": "用药体验与疗效反馈"
  },
  "project_background": ["..."],
  "project_execution": ["..."],
  "questionnaire_note": ["..."],
  "result_analysis": {
    "sections": [
      {
        "section_number": "4.1",
        "section_title": "药品疗效",
        "paragraphs": ["..."],
        "question_refs": ["q01"],
        "chart_refs": ["chart_01"]
      }
    ]
  },
  "summary": {
    "key_issue_analysis": ["..."],
    "overall_analysis": ["..."],
    "recommendations": ["..."]
  },
  "attachments": {
    "attachment1_name": "厄贝沙坦氢氯噻嗪片用药体验与疗效反馈患者调查问卷",
    "attachment1_questions": [],
    "attachment2_name": "问卷调查明细表"
  },
  "disclaimer": {
    "title": "免责申明",
    "items": [],
    "unit": "北京玖麟空科技有限公司"
  }
}
```
