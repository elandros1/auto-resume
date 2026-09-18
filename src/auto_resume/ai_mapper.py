"""AI-powered field mapping using LLM APIs.

This module provides the "ultimate" solution for universal resume filling:
instead of relying on hardcoded regex patterns, it sends the field labels
from a Word template to an LLM (like DeepSeek, OpenAI, etc.) and asks it
to map them to the user's resume data fields.

Privacy: Only field LABELS (e.g. "姓名", "联系电话") and resume data
KEYS (e.g. "name", "phone") are sent to the AI. Personal data values
are NEVER sent — the AI only does the mapping, filling happens locally.

Usage:
    mapper = AIFieldMapper(api_key="sk-...", model="deepseek-chat")
    mapping = mapper.map_fields(doc_path, resume_keys)
    # mapping = {"姓名": "name", "联系电话": "phone", ...}
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from docx import Document


class AIFieldMapper:
    """Maps Word template field labels to resume data keys using LLM.

    Supports any OpenAI-compatible API (DeepSeek, OpenAI, Moonshot, etc.)

    Privacy: Only labels and keys are sent to the AI, never personal data.
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str = "deepseek-chat",
        base_url: str = "https://api.deepseek.com/v1",
    ):
        self.api_key = api_key or os.environ.get("AI_API_KEY", "")
        self.model = model
        self.base_url = base_url

    def is_available(self) -> bool:
        """Check if AI mapping is available (has API key)."""
        return bool(self.api_key)

    def extract_labels(self, doc_path: str | Path) -> list[dict[str, str]]:
        """Extract all field labels from a Word document.

        Returns a list of {label, location} dicts.
        Location format: "P:0" (paragraph 0) or "T:0:R:1:C:2" (table/row/col)
        """
        doc_path = Path(doc_path)
        doc = Document(str(doc_path))

        labels: list[dict[str, str]] = []
        seen: set[str] = set()

        # Extract from paragraphs
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if text and text not in seen:
                # Skip very long paragraphs (likely content, not labels)
                if len(text) <= 30:
                    seen.add(text)
                    labels.append({"label": text, "location": f"P:{i}"})

        # Extract from tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip().replace("\n", " ")
                    if text and text not in seen:
                        if len(text) <= 50:
                            seen.add(text)
                            labels.append({
                                "label": text,
                                "location": f"T:{t_idx}:R:{r_idx}:C:{c_idx}",
                            })

        return labels

    def extract_table_headers(
        self, doc_path: str | Path
    ) -> list[dict[str, Any]]:
        """Extract table header rows for column mapping.

        Returns list of {
            table_idx, row_idx, headers: [{col_idx, text}],
            location_hint (e.g. "education", "work", "family")
        }
        """
        doc_path = Path(doc_path)
        doc = Document(str(doc_path))

        import re

        section_hints = {
            r"教育|学习": "education",
            r"工作|职业": "work",
            r"项目|科研": "project",
            r"论文|发表": "publication",
            r"获奖|荣誉": "award",
            r"家庭|社会关系": "family",
        }

        header_tables: list[dict[str, Any]] = []

        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                cells_text = [c.text.strip().replace("\n", " ") for c in row.cells]

                # Check if this looks like a header row
                # (has 3+ non-empty cells with short text)
                non_empty = [c for c in cells_text if c and len(c) <= 20]
                if len(non_empty) >= 3:
                    # Check if it's in a section
                    hint = "unknown"
                    # Look at rows above for section headers
                    for prev_r in range(max(0, r_idx - 3), r_idx):
                        prev_text = " ".join(
                            c.text.strip() for c in table.rows[prev_r].cells
                        )
                        for pattern, section_type in section_hints.items():
                            if re.search(pattern, prev_text):
                                hint = section_type
                                break

                    # Also check current row
                    row_text = " ".join(cells_text)
                    for pattern, section_type in section_hints.items():
                        if re.search(pattern, row_text):
                            hint = section_type
                            break

                    headers = []
                    for c_idx, text in enumerate(cells_text):
                        if text and len(text) <= 20:
                            headers.append({"col_idx": c_idx, "text": text})

                    if headers:
                        header_tables.append({
                            "table_idx": t_idx,
                            "row_idx": r_idx,
                            "section": hint,
                            "headers": headers,
                        })

        return header_tables

    def build_prompt(
        self,
        labels: list[dict[str, str]],
        resume_keys: list[str],
        table_headers: list[dict[str, Any]] | None = None,
    ) -> str:
        """Build the prompt for the LLM.

        Only sends labels and keys — NO personal data values.
        """
        labels_text = "\n".join(
            f"  {i+1}. \"{item['label']}\" (位置: {item['location']})"
            for i, item in enumerate(labels)
        )

        keys_text = "\n".join(f"  - {k}" for k in resume_keys)

        headers_text = ""
        if table_headers:
            headers_text = "\n\n## 表格列头（需要映射到列表字段后缀）\n"
            for h in table_headers:
                section = h["section"]
                cols = ", ".join(
                    f'列{c["col_idx"]}="{c["text"]}"' for c in h["headers"]
                )
                headers_text += f"  区域={section}, {cols}\n"

        prompt = f"""你是一个简历表格字段映射专家。请将以下Word模板中的字段标签映射到简历数据字段。

## 简历数据可用字段（key列表）
{keys_text}

## Word模板中的字段标签
{labels_text}
{headers_text}

## 映射规则
1. 每个标签映射到一个最合适的简历字段key
2. 如果标签明显是表格列头（如"起止时间"、"学校名称"），映射到对应列表项的后缀
   - 教育经历列头: school, major, degree, start_date, end_date, duration, \
research_direction, education_level, education_form, reference_person, \
reference_phone
   - 工作经历列头: company, position, start_date, end_date, department, \
description, reference_person, reference_phone, professional_title
   - 家庭成员列头: name, relationship, gender, birth_date, phone, work_unit, \
political_status
3. 如果标签是区域标题（如"学习经历"、"工作经历"），映射到 "section:education" 等格式
4. 如果标签无法映射到任何字段，返回 "unmatched"
5. 注意上下文：如果标签在"配偶"区域内，应映射到spouse_前缀字段

## 输出格式
请返回JSON格式，key是标签文字，value是字段名：
```json
{{
  "姓名": "name",
  "联系电话": "phone",
  "起止时间": "education.date_range",
  "学习经历": "section:education"
}}
```

只返回JSON，不要其他文字。"""

        return prompt

    def map_fields(
        self,
        doc_path: str | Path,
        resume_keys: list[str],
    ) -> dict[str, str]:
        """Map all field labels in a Word document to resume keys.

        This is the main entry point. Sends only labels and keys to AI.
        Returns: {label_text: resume_key_or_section}
        """
        if not self.is_available():
            raise RuntimeError(
                "AI API key not configured. Set AI_API_KEY environment "
                "variable or pass api_key parameter."
            )

        # Extract labels and table headers
        labels = self.extract_labels(doc_path)
        table_headers = self.extract_table_headers(doc_path)

        if not labels:
            return {}

        # Build prompt
        prompt = self.build_prompt(labels, resume_keys, table_headers)

        # Call LLM API
        response = self._call_llm(prompt)

        # Parse JSON response
        mapping = self._parse_response(response)

        return mapping

    def _call_llm(self, prompt: str) -> str:
        """Call the LLM API and return the response text."""
        try:
            import urllib.request

            url = f"{self.base_url}/chat/completions"
            payload = json.dumps({
                "model": self.model,
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "你是一个专业的简历表格字段映射专家。"
                            "你只做字段映射，不涉及任何个人数据。"
                            "只返回JSON格式结果。"
                        ),
                    },
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.1,
                "max_tokens": 4000,
            }).encode("utf-8")

            req = urllib.request.Request(
                url,
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}",
                },
                method="POST",
            )

            # Use proxy if available
            import urllib.request as ur
            proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
            if proxy_url:
                proxy_handler = ur.ProxyHandler({
                    "https": proxy_url,
                    "http": proxy_url,
                })
                opener = ur.build_opener(proxy_handler)
                response = opener.open(req, timeout=60)
            else:
                response = ur.urlopen(req, timeout=60)

            result = json.loads(response.read().decode("utf-8"))
            return result["choices"][0]["message"]["content"]

        except Exception as e:
            raise RuntimeError(f"LLM API call failed: {e}") from e

    def _parse_response(self, response: str) -> dict[str, str]:
        """Parse the LLM response into a {label: key} mapping."""
        # Extract JSON from response (may be wrapped in markdown)
        import re

        # Try to find JSON block
        json_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", response, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            # Try raw JSON
            json_str = response.strip()
            # Remove any leading/trailing text
            start = json_str.find("{")
            end = json_str.rfind("}")
            if start != -1 and end != -1:
                json_str = json_str[start : end + 1]

        try:
            mapping = json.loads(json_str)
            # Clean up: remove "unmatched" entries
            return {
                k: v for k, v in mapping.items()
                if v and v != "unmatched"
            }
        except json.JSONDecodeError:
            return {}

    def map_with_context(
        self,
        doc_path: str | Path,
        resume_keys: list[str],
    ) -> dict[str, str]:
        """Full mapping with context-aware section detection.

        Returns a comprehensive mapping that includes:
        - Single-value field mappings: {label: key}
        - Section markers: {label: "section:education"}
        - Column mappings: {header: "col:school"}

        This can be used by the detector to fill everything.
        """
        return self.map_fields(doc_path, resume_keys)
