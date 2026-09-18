"""Smart field detection for Word templates.

Automatically detects Chinese field labels in Word documents and maps them
to resume data keys, so users don't have to manually mark placeholders.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document

# Mapping from Chinese field labels (regex patterns) to resume keys
FIELD_MAPPINGS: list[tuple[str, str]] = [
    # Basic info
    (r"姓\s*名", "name"),
    (r"性\s*别", "gender"),
    (r"出生年月|出生日期|生日", "birth_date"),
    (r"联系电话|手机|电话", "phone"),
    (r"电子邮箱|邮箱|Email|E-mail|电子邮件", "email"),
    (r"籍\s*贯", "hometown"),
    (r"政治面貌", "political_status"),
    (r"身份证号|证件号码", "id_number"),
    (r"民\s*族", "nationality"),
    (r"婚姻状况", "marital_status"),
    (r"通讯地址|联系地址|地址", "address"),
    (r"照片|相片", "photo_path"),

    # Job intent
    (r"求职意向|应聘岗位|期望岗位|意向岗位", "expected_position"),
    (r"期望薪资|期望薪酬|薪资要求", "expected_salary"),
    (r"期望城市|意向城市|期望工作地点", "expected_city"),
    (r"到岗时间|可入职时间", "availability"),

    # Self evaluation
    (r"自我评价|自我介绍|个人简介", "self_evaluation"),

    # Education
    (r"毕业院校|学校名称|院校", "education_summary"),
    (r"专\s*业", "education_1_major"),
    (r"学\s*历|学位", "education_1_degree"),

    # Work
    (r"工作经历|工作经验", "work_experience_summary"),
    (r"工作单位|单位名称", "work_1_company"),
    (r"职\s*位|职务", "work_1_position"),

    # Skills
    (r"专业技能|技能特长|技能", "skills_summary"),
    (r"证书|资格证书|职业证书", "certificates_summary"),
    (r"语言能力|外语水平|语言", "languages_summary"),
    (r"兴趣爱好|爱好|特长", "hobbies_summary"),

    # Projects
    (r"项目经验|项目经历", "projects_summary"),
]


class FieldDetector:
    """Detect form fields in a Word template and map them to resume data."""

    def __init__(self, mappings: list[tuple[str, str]] | None = None):
        self.mappings = mappings or FIELD_MAPPINGS
        self._compiled = [(re.compile(p, re.IGNORECASE), k) for p, k in self.mappings]

    def detect_fields(self, doc_path: str | Path) -> dict[str, list[str]]:
        """Scan a Word document for recognizable field labels.

        Returns a dict: {resume_key: [list of cell/paragraph references]}
        Each reference is "P:0" (paragraph index) or "T:0:R:1:C:2" (table, row, cell).
        """
        doc_path = Path(doc_path)
        doc = Document(str(doc_path))

        results: dict[str, list[str]] = {}

        # Scan paragraphs
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            for pattern, key in self._compiled:
                if pattern.search(text):
                    if key not in results:
                        results[key] = []
                    results[key].append(f"P:{i}")

        # Scan tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if not text:
                        continue
                    for pattern, key in self._compiled:
                        if pattern.search(text):
                            if key not in results:
                                results[key] = []
                            ref = f"T:{t_idx}:R:{r_idx}:C:{c_idx}"
                            if ref not in results[key]:
                                results[key].append(ref)

        return results

    def get_blank_cells_near_labels(
        self, doc_path: str | Path
    ) -> dict[str, str]:
        """Find blank cells next to label cells in tables.

        When a table cell contains a label like "姓名", the adjacent cell
        is usually where the value should go.

        Returns: {resume_key: cell_reference}
        """
        doc_path = Path(doc_path)
        doc = Document(str(doc_path))

        results: dict[str, str] = {}

        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                cells = row.cells
                for c_idx, cell in enumerate(cells):
                    text = cell.text.strip()
                    if not text:
                        continue

                    for pattern, key in self._compiled:
                        if pattern.search(text):
                            # Look for adjacent blank cell (right side)
                            if c_idx + 1 < len(cells):
                                next_cell = cells[c_idx + 1]
                                if not next_cell.text.strip():
                                    ref = f"T:{t_idx}:R:{r_idx}:C:{c_idx + 1}"
                                    if key not in results:
                                        results[key] = ref
                                    break

                            # Also check if label and blank are merged
                            # (label:text: format in same cell)
                            elif ":" in text or "：" in text:
                                # Label cell itself has space for value
                                ref = f"T:{t_idx}:R:{r_idx}:C:{c_idx}"
                                if key not in results:
                                    results[key] = ref
                                break

        return results

    def auto_fill(
        self,
        doc_path: str | Path,
        data: dict[str, str],
        output_path: str | Path,
    ) -> Path:
        """Auto-fill a Word document by detecting fields and filling values.

        This is the smart mode: it scans the document for Chinese labels,
        finds the blank cells next to them, and fills in the values.

        Args:
            doc_path: Path to the original Word document.
            data: Flat dict of resume data (from ResumeData.to_flat_dict()).
            output_path: Where to save the filled document.

        Returns:
            Path to the saved file.
        """
        doc_path = Path(doc_path)
        output_path = Path(output_path)
        doc = Document(str(doc_path))

        # First: fill existing {{placeholder}} style markers
        for para in doc.paragraphs:
            self._fill_placeholders_in_paragraph(para, data)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._fill_placeholders_in_paragraph(para, data)

        # Second: smart-detect label cells and fill adjacent blanks
        filled_count = 0
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                cells = row.cells
                for c_idx, cell in enumerate(cells):
                    text = cell.text.strip()

                    # Skip cells that already have content (already filled or not a label)
                    # Check if this is a label cell with a blank neighbor
                    for pattern, key in self._compiled:
                        if pattern.search(text):
                            value = data.get(key, "")
                            if not value:
                                continue

                            # Try adjacent right cell
                            if c_idx + 1 < len(cells):
                                next_cell = cells[c_idx + 1]
                                if not next_cell.text.strip():
                                    self._set_cell_text(next_cell, value)
                                    filled_count += 1
                                    break

                            # Try the same cell if it has "label:___" format
                            # Replace trailing colons/underscores with value
                            elif text.endswith(":") or text.endswith("："):
                                new_text = text + " " + value
                                self._set_cell_text(cell, new_text)
                                filled_count += 1
                                break

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
        return output_path

    def _fill_placeholders_in_paragraph(self, para, data: dict[str, str]) -> None:
        """Fill {{placeholder}} style markers in a paragraph."""
        text = para.text
        if "{{" not in text and "${" not in text:
            return

        def replacer(m):
            key = m.group(1) or m.group(2)
            return data.get(key, m.group(0))

        new_text = re.sub(
            r"\{\{(\w+)\}\}|\$\{(\w+)\}",
            replacer,
            text,
        )

        if new_text != text and para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""

    def _set_cell_text(self, cell, text: str) -> None:
        """Set text in a table cell, preserving first paragraph formatting."""
        if cell.paragraphs:
            para = cell.paragraphs[0]
            if para.runs:
                para.runs[0].text = text
            else:
                para.add_run(text)
            # Clear extra paragraphs
            for extra in cell.paragraphs[1:]:
                for run in extra.runs:
                    run.text = ""
        else:
            cell.text = text

    def preview_mapping(
        self, doc_path: str | Path
    ) -> list[dict[str, str]]:
        """Preview what fields would be detected and mapped.

        Returns a list of dicts with: label, resume_key, location
        Useful for showing the user what will be auto-filled.
        """
        doc_path = Path(doc_path)
        doc = Document(str(doc_path))

        previews: list[dict[str, str]] = []

        # Scan paragraphs
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            for pattern, key in self._compiled:
                if pattern.search(text):
                    previews.append({
                        "label": text[:50],
                        "resume_key": key,
                        "location": f"段落 {i}",
                    })

        # Scan tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    for pattern, key in self._compiled:
                        if pattern.search(text):
                            previews.append({
                                "label": text[:50],
                                "resume_key": key,
                                "location": f"表格{t_idx} 行{r_idx} 列{c_idx}",
                            })

        return previews
