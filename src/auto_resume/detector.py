"""Smart field detection for Word templates.

Automatically detects Chinese field labels in Word documents and maps them
to resume data keys, so users don't have to manually mark placeholders.

Upgraded features:
- Table header recognition + column-to-field mapping
- Multi-row batch filling for education/work/project sections
- Multi-direction blank cell lookup (right / below / same cell)
- Merged cell awareness
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

# ──────────────────── Field Mapping Tables ────────────────────

# Single-value field mappings: (regex_pattern, resume_key)
FIELD_MAPPINGS: list[tuple[str, str]] = [
    # ── Basic personal info ──
    (r"姓\s*名", "name"),
    (r"性\s*别", "gender"),
    (r"出生年月|出生日期|生日", "birth_date"),
    (r"年\s*龄", "age"),
    (r"联系电话|手机|电话", "phone"),
    (r"电子邮箱|邮箱|Email|E-mail|电子邮件", "email"),
    (r"籍\s*贯", "hometown"),
    (r"政治面貌", "political_status"),
    (r"身份证号|证件号码|身份证号码", "id_number"),
    (r"民\s*族", "nationality"),
    (r"婚姻状况|是否已婚|婚否", "marital_status"),
    (r"通讯地址|联系地址|地址", "address"),
    (r"邮\s*编|邮政编码", "postal_code"),
    (r"户口所在地|户籍地|户口", "hukou_location"),
    (r"身\s*高", "height"),
    (r"视\s*力", "vision"),
    (r"照片|相片", "photo_path"),
    # ── Professional info ──
    (r"专业特长|特长", "specialty"),
    (r"计算机熟练程度|计算机水平|计算机能力", "computer_proficiency"),
    (r"外语熟练程度|外语水平|外语能力|外语程度", "foreign_language"),
    (r"原单位|原工作单位", "previous_employer"),
    (r"从事的岗位|现任岗位|现任职务", "current_position"),
    (r"职\s*称", "professional_title"),
    (r"社会经历|社会实践", "social_experience"),
    (r"所学专业", "education_1_major"),
    (r"学\s*历", "education_1_degree"),
    (r"学\s*位", "education_1_degree"),
    # ── Job seeking ──
    (r"求职意向|意向岗位", "expected_position"),
    (r"应聘岗位|应聘的岗位|应聘岗位", "applied_position"),
    (r"期望岗位|期望职位", "expected_position"),
    (r"期望薪资|期望薪酬|薪资要求", "expected_salary"),
    (r"期望城市|意向城市|期望工作地点", "expected_city"),
    (r"到岗时间|可入职时间", "availability"),
    # ── Self evaluation ──
    (r"自我评价|自我介绍|个人简介", "self_evaluation"),
    # ── Spouse info ──
    (r"配偶姓名|爱人姓名", "spouse_name"),
    (r"配偶出生年月|配偶生日", "spouse_birth_date"),
    (r"配偶工作单位|配偶单位", "spouse_work_unit"),
    (r"配偶电话", "spouse_phone"),
    # ── Summary fields (for paragraph-style templates) ──
    (r"毕业院校|学校名称|院校", "education_summary"),
    (r"专业技能|技能特长|技能", "skills_summary"),
    (r"证书|资格证书|职业证书", "certificates_summary"),
    (r"语言能力|外语水平|语言", "languages_summary"),
    (r"兴趣爱好|爱好|特长", "hobbies_summary"),
]

# Section header patterns: identify multi-row table sections
# Maps section header regex -> data prefix used in flat dict
# e.g. "education" prefix produces keys: education_1_school, education_1_major, etc.
SECTION_PATTERNS: dict[str, str] = {
    r"教育经历|学习经历|教育背景|学历背景": "education",
    r"工作经历|工作经验|工作背景|职业经历": "work",
    r"项目经验|项目经历|科研项目|主持的主要科研项目": "project",
    r"发表论文|论文列表|学术成果|发表的论文": "publication",
    r"获奖情况|荣誉奖项|获奖经历": "award",
    r"家庭成员|家庭情况|主要社会关系|家庭关系": "family",
    r"社会经历|社会实践": "social",
}

# Column header patterns: map column header text to field suffix
COLUMN_MAPPINGS: dict[str, dict[str, str]] = {
    "education": {
        r"起止时间|时间|起讫时间|在校时间": "date_range",
        r"院校名称|学校名称|院校|毕业院校|学校": "school",
        r"专\s*业": "major",
        r"学\s*历": "degree",
        r"学\s*位": "degree",
        r"学习层次|层次": "education_level",
        r"办学形式|培养方式": "education_form",
        r"证明人": "reference_person",
        r"证明人电话|证明人电话|联系电话": "reference_phone",
        r"GPA|成绩": "gpa",
        r"备注|说明|描述": "description",
    },
    "work": {
        r"起止时间|时间|起讫时间|工作时间": "date_range",
        r"工作单位|单位名称|单位|公司|任何单位": "company",
        r"职\s*位|职务|岗位|担任的职务": "position",
        r"职称|专业技术职务": "professional_title",
        r"部门|院系": "department",
        r"证明人": "reference_person",
        r"证明人电话|联系电话": "reference_phone",
        r"工作内容|职责|描述|备注": "description",
    },
    "project": {
        r"起止时间|时间|项目时间": "date_range",
        r"项目名称|名称": "name",
        r"角\s*色|职务|承担工作|主持|参与": "role",
        r"技术|技术栈|使用技术": "technologies",
        r"经费来源|来源": "funding_source",
        r"经费|金额|经费金额": "funding_amount",
        r"描述|内容|备注": "description",
    },
    "publication": {
        r"序号|编号": "index",
        r"论文题目|题目|名称": "title",
        r"期刊|发表刊物|杂志|发表期刊": "journal",
        r"时间|日期|发表时间": "date",
        r"作者|作者顺序|署名": "authors",
        r"收录|收录情况|SCI|EI|索引": "index",
    },
    "award": {
        r"时间|日期|获奖时间": "date",
        r"奖项名称|名称|获奖": "title",
        r"级别|等级": "level",
        r"颁发单位|授予单位": "issuer",
    },
    "family": {
        r"姓名|称谓": "name",
        r"关系|与本人关系": "relationship",
        r"性\s*别": "gender",
        r"出生年月|出生日期|年龄": "birth_date",
        r"电话|联系电话|手机": "phone",
        r"工作单位|单位": "work_unit",
        r"职务|岗位": "position",
        r"地址": "address",
        r"政治面貌": "political_status",
    },
    "social": {
        r"起止时间|时间": "date_range",
        r"内容|经历|描述": "description",
        r"地点|单位": "location",
    },
}


class FieldDetector:
    """Detect form fields in a Word template and map them to resume data."""

    def __init__(self, mappings: list[tuple[str, str]] | None = None):
        self.mappings = mappings or FIELD_MAPPINGS
        self._compiled = [
            (re.compile(p, re.IGNORECASE), k) for p, k in self.mappings
        ]
        self._section_compiled = {
            re.compile(p, re.IGNORECASE): prefix
            for p, prefix in SECTION_PATTERNS.items()
        }
        self._column_compiled = {
            prefix: [
                (re.compile(p, re.IGNORECASE), suffix)
                for p, suffix in cols.items()
            ]
            for prefix, cols in COLUMN_MAPPINGS.items()
        }

    # ──────────────────── Public API ────────────────────

    def detect_fields(self, doc_path: str | Path) -> dict[str, list[str]]:
        """Scan a Word document for recognizable field labels.

        Returns a dict: {resume_key: [list of cell/paragraph references]}
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
                    results.setdefault(key, []).append(f"P:{i}")

        # Scan tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if not text:
                        continue
                    for pattern, key in self._compiled:
                        if pattern.search(text):
                            ref = f"T:{t_idx}:R:{r_idx}:C:{c_idx}"
                            if ref not in results.setdefault(key, []):
                                results[key].append(ref)

        return results

    def preview_mapping(self, doc_path: str | Path) -> list[dict[str, str]]:
        """Preview what fields would be detected and mapped."""
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

        # Scan tables — include section header detection
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if not text:
                        continue

                    # Check section headers
                    for pattern, prefix in self._section_compiled.items():
                        if pattern.search(text):
                            previews.append({
                                "label": text[:50],
                                "resume_key": f"{prefix}_section",
                                "location": f"表格{t_idx} 行{r_idx} 列{c_idx}",
                            })

                    # Check individual field labels
                    for pattern, key in self._compiled:
                        if pattern.search(text):
                            previews.append({
                                "label": text[:50],
                                "resume_key": key,
                                "location": f"表格{t_idx} 行{r_idx} 列{c_idx}",
                            })

        return previews

    def auto_fill(
        self,
        doc_path: str | Path,
        data: dict[str, str],
        output_path: str | Path,
    ) -> Path:
        """Auto-fill a Word document by detecting fields and filling values.

        Strategy (in order):
        1. Fill {{placeholder}} and ${placeholder} style markers
        2. Detect multi-row sections (education/work/projects) and fill by column
        3. Smart-detect single-value label cells and fill adjacent blanks

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

        # Phase 1: Fill placeholder markers
        self._fill_all_placeholders(doc, data)

        # Phase 2: Detect and fill multi-row sections
        self._fill_section_tables(doc, data)

        # Phase 3: Smart-detect remaining single-value labels
        self._fill_single_value_labels(doc, data)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
        return output_path

    # ──────────────────── Phase 1: Placeholders ────────────────────

    def _fill_all_placeholders(self, doc: Document, data: dict[str, str]) -> None:
        """Fill {{placeholder}} and ${placeholder} style markers everywhere."""
        for para in doc.paragraphs:
            self._fill_placeholders_in_paragraph(para, data)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._fill_placeholders_in_paragraph(para, data)

    def _fill_placeholders_in_paragraph(
        self, para: Paragraph, data: dict[str, str]
    ) -> None:
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

    # ──────────────────── Phase 2: Section Tables ────────────────────

    def _fill_section_tables(self, doc: Document, data: dict[str, str]) -> None:
        """Detect multi-row section tables and fill them by column mapping.

        This handles tables like:
        | 教育经历 (merged header) |
        | 起止时间 | 院校名称 | 专业 | 学位 |
        | (blank)  | (blank)  | (blank) | (blank) |
        | (blank)  | (blank)  | (blank) | (blank) |

        The blank rows are filled with education_1_*, education_2_*, etc.
        """
        for table in doc.tables:
            self._process_single_table_sections(table, data)

    def _process_single_table(self, table: Table, data: dict[str, str]) -> None:
        """Process a single table for section detection and filling."""
        # First: detect if any cell is a section header
        section_prefix = None
        header_row_idx = None

        for r_idx, row in enumerate(table.rows):
            row_text = " ".join(cell.text.strip() for cell in row.cells).strip()
            for pattern, prefix in self._section_compiled.items():
                if pattern.search(row_text):
                    section_prefix = prefix
                    header_row_idx = r_idx
                    break
            if section_prefix:
                break

        if not section_prefix:
            return

        # Find the column header row (usually the row after the section header)
        col_header_row_idx = header_row_idx + 1
        if col_header_row_idx >= len(table.rows):
            return

        col_header_row = table.rows[col_header_row_idx]
        column_map = self._detect_column_mappings(col_header_row, section_prefix)

        if not column_map:
            # Maybe the section header row IS the column header row
            # Try the header row itself
            column_map = self._detect_column_mappings(
                table.rows[header_row_idx], section_prefix
            )
            if column_map:
                col_header_row_idx = header_row_idx
            else:
                return

        # Fill data rows (everything after the column header row)
        data_count = int(data.get(f"{section_prefix}_count", "0"))
        if data_count == 0:
            return

        for entry_idx in range(1, data_count + 1):
            target_row_idx = col_header_row_idx + entry_idx
            if target_row_idx >= len(table.rows):
                # Table doesn't have enough rows — skip extra entries
                break

            row = table.rows[target_row_idx]
            self._fill_row_by_column_map(row, column_map, section_prefix, entry_idx, data)

    def _process_single_table_sections(
        self, table: Table, data: dict[str, str]
    ) -> None:
        """Process a table, handling both section-header and non-section tables."""
        # Detect section headers anywhere in the table
        section_locations: list[tuple[int, str]] = []

        for r_idx, row in enumerate(table.rows):
            row_text = " ".join(cell.text.strip() for cell in row.cells).strip()
            for pattern, prefix in self._section_compiled.items():
                if pattern.search(row_text):
                    section_locations.append((r_idx, prefix))
                    break

        if not section_locations:
            return

        for header_r_idx, prefix in section_locations:
            self._fill_one_section(table, header_r_idx, prefix, data)

    def _fill_one_section(
        self,
        table: Table,
        header_r_idx: int,
        prefix: str,
        data: dict[str, str],
    ) -> None:
        """Fill one section within a table, starting at the given header row."""

        # Find the column header row
        col_header_r_idx = header_r_idx + 1
        if col_header_r_idx >= len(table.rows):
            # Maybe header row itself has column labels
            col_header_r_idx = header_r_idx
            column_map = self._detect_column_mappings(
                table.rows[col_header_r_idx], prefix
            )
        else:
            column_map = self._detect_column_mappings(
                table.rows[col_header_r_idx], prefix
            )
            if not column_map:
                # Try the header row itself
                col_header_r_idx = header_r_idx
                column_map = self._detect_column_mappings(
                    table.rows[col_header_r_idx], prefix
                )

        if not column_map:
            return

        # Determine how many entries to fill
        # Handle inconsistent count key naming across sections
        count_key = f"{prefix}_count"
        if prefix == "work":
            count_key = "work_experience_count"
        elif prefix == "family":
            count_key = "family_members_count"
        elif prefix == "social":
            # Social experience is a single string field, not a list
            data_count = 1
            count_key = ""
        data_count = int(data.get(count_key, "0")) if count_key else data_count
        if data_count == 0:
            return

        # Fill data rows after the column header row
        # If header row == column header row, data starts at header_r_idx + 1
        # Otherwise data starts at col_header_r_idx + 1
        data_start_r_idx = col_header_r_idx + 1

        # Don't fill into the next section's rows
        next_section_r_idx = len(table.rows)  # default: end of table
        for r_idx in range(data_start_r_idx, len(table.rows)):
            row_text = " ".join(
                cell.text.strip() for cell in table.rows[r_idx].cells
            ).strip()
            for pattern, other_prefix in self._section_compiled.items():
                if pattern.search(row_text) and other_prefix != prefix:
                    next_section_r_idx = r_idx
                    break

        available_rows = next_section_r_idx - data_start_r_idx

        for entry_idx in range(1, min(data_count, available_rows) + 1):
            target_r_idx = data_start_r_idx + entry_idx - 1
            row = table.rows[target_r_idx]
            self._fill_row_by_column_map(
                row, column_map, prefix, entry_idx, data
            )

    def _detect_column_mappings(
        self, row, prefix: str
    ) -> dict[int, str]:
        """Detect which column maps to which field suffix.

        Returns: {col_index: field_suffix}
        e.g. {0: "date_range", 1: "school", 2: "major", 3: "degree"}
        """
        col_patterns = self._column_compiled.get(prefix, [])
        if not col_patterns:
            return {}

        column_map: dict[int, str] = {}

        for c_idx, cell in enumerate(row.cells):
            text = cell.text.strip()
            if not text:
                continue
            for pattern, suffix in col_patterns:
                if pattern.search(text):
                    # Don't overwrite if already mapped (merged cells)
                    if c_idx not in column_map:
                        column_map[c_idx] = suffix
                    break

        return column_map

    def _fill_row_by_column_map(
        self,
        row,
        column_map: dict[int, str],
        prefix: str,
        entry_idx: int,
        data: dict[str, str],
    ) -> None:
        """Fill a single row based on the column mapping.

        For each column, look up the corresponding field value:
        - "date_range" -> combine {prefix}_{idx}_start_date and _end_date
        - "index" -> entry number
        - Other suffixes -> {prefix}_{idx}_{suffix}
        """
        for c_idx, suffix in column_map.items():
            if c_idx >= len(row.cells):
                continue

            cell = row.cells[c_idx]

            # Skip cells that already have content (don't overwrite)
            if cell.text.strip():
                continue

            if suffix == "date_range":
                start = data.get(f"{prefix}_{entry_idx}_start_date", "")
                end = data.get(f"{prefix}_{entry_idx}_end_date", "")
                value = f"{start} - {end}" if start and end else f"{start}{end}"
            elif suffix == "index":
                value = str(entry_idx)
            elif prefix == "social":
                value = data.get("social_experience", "")
            else:
                value = data.get(f"{prefix}_{entry_idx}_{suffix}", "")

            if value:
                self._set_cell_text(cell, value)

    # ──────────────────── Phase 3: Single-Value Labels ────────────────────

    def _fill_single_value_labels(self, doc: Document, data: dict[str, str]) -> None:
        """Smart-detect label cells and fill adjacent blanks (multi-direction)."""
        for table in doc.tables:
            for r_idx, row in enumerate(table.rows):
                cells = row.cells
                for c_idx, cell in enumerate(cells):
                    text = cell.text.strip()
                    if not text:
                        continue

                    for pattern, key in self._compiled:
                        if not pattern.search(text):
                            continue

                        value = data.get(key, "")
                        if not value:
                            continue

                        # Skip if already filled (section filler may have done it)
                        if value and value in cell.text:
                            break

                        # Strategy 1: Adjacent right cell
                        if c_idx + 1 < len(cells):
                            next_cell = cells[c_idx + 1]
                            if not next_cell.text.strip():
                                self._set_cell_text(next_cell, value)
                                break

                        # Strategy 2: Cell below (next row, same column)
                        if r_idx + 1 < len(table.rows):
                            below_cell = table.rows[r_idx + 1].cells[c_idx]
                            if not below_cell.text.strip():
                                self._set_cell_text(below_cell, value)
                                break

                        # Strategy 3: Same cell with "label:___" format
                        if text.endswith(":") or text.endswith("："):
                            new_text = text + " " + value
                            self._set_cell_text(cell, new_text)
                            break

                        # Strategy 4: Same cell with "label：" + empty space
                        if "：" in text or ":" in text:
                            # Replace trailing colons/underscores
                            new_text = re.sub(
                                r"[:：]\s*$", f"：{value}", text
                            )
                            if new_text != text:
                                self._set_cell_text(cell, new_text)
                                break

        # Also fill paragraph-based labels (label: value format)
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            for pattern, key in self._compiled:
                if pattern.search(text):
                    value = data.get(key, "")
                    if not value:
                        continue
                    # Check if text is just the label (no value yet)
                    # e.g., "姓名：" with nothing after
                    match = re.match(
                        r"^(.*?[:：])\s*$", text
                    )
                    if match and not text[len(match.group(1)):].strip():
                        prefix_label = match.group(1)
                        if para.runs:
                            para.runs[0].text = f"{prefix_label} {value}"
                            for run in para.runs[1:]:
                                run.text = ""
                        break

    # ──────────────────── Utility Methods ────────────────────

    def _set_cell_text(self, cell, text: str) -> None:
        """Set text in a table cell, preserving first paragraph formatting."""
        if cell.paragraphs:
            para = cell.paragraphs[0]
            if para.runs:
                para.runs[0].text = text
            else:
                run = para.add_run(text)
                # Try to copy formatting from adjacent cell's run
                run.font.size = None  # inherit
            # Clear extra paragraphs
            for extra in cell.paragraphs[1:]:
                for run in extra.runs:
                    run.text = ""
        else:
            cell.text = text

    def _find_blank_cell_multi_direction(
        self, table: Table, r_idx: int, c_idx: int
    ) -> str | None:
        """Find a blank cell near (r_idx, c_idx) in multiple directions.

        Returns the cell reference string, or None if no blank found.
        Priority: right > below > same-cell-with-colon.
        """
        cells = table.rows[r_idx].cells

        # Right
        if c_idx + 1 < len(cells):
            next_cell = cells[c_idx + 1]
            if not next_cell.text.strip():
                return f"R:{r_idx}C:{c_idx + 1}"

        # Below
        if r_idx + 1 < len(table.rows):
            below_cell = table.rows[r_idx + 1].cells[c_idx]
            if not below_cell.text.strip():
                return f"R:{r_idx + 1}C:{c_idx}"

        return None
