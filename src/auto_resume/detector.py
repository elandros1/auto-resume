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
# IMPORTANT: More specific patterns MUST come before generic ones
# to avoid false matches (e.g. "证明人电话" before "证明人", "紧急联系人" before "联系人")
FIELD_MAPPINGS: list[tuple[str, str]] = [
    # ── Basic personal info ──
    (r"姓\s*名", "name"),
    (r"性\s*别", "gender"),
    (r"出生年月|出生日期|生日|出生时间", "birth_date"),
    (r"年\s*龄", "age"),
    # phone: 联系方式/联系电话/电话 — context-aware for emergency contact
    (r"联系电话|联系方式|手机|联络电话|移动电话|手机号码|电\s*话", "phone"),
    (r"电子邮箱|邮箱|Email|E-mail|电子邮件|电子信箱|Email地址", "email"),
    (r"籍\s*贯|籍贯地|出生地|生源地|原籍", "hometown"),
    (r"政治面貌|党派", "political_status"),
    (r"身份证号|证件号码|身份证号码|身份证明|身份证", "id_number"),
    (r"民\s*族|族\s*别", "nationality"),
    (r"婚姻状况|是否已婚|婚否|婚\s*姻", "marital_status"),
    # address: 现通讯地址 before 户籍地址, both before generic 地址
    (r"现通讯地址|通讯地址|联系地址|家庭住址|现住址|居住地|住\s*址", "address"),
    (r"现户籍地址|户口所在地|户籍地|户籍所在地|户口所在|户\s*口", "hukou_location"),
    (r"邮\s*编|邮政编码|邮编代码", "postal_code"),
    (r"身\s*高|身高cm|身高CM", "height"),
    (r"视\s*力|眼睛度数", "vision"),
    (r"照片|相片", "photo_path"),
    # ── Professional info (specific before generic) ──
    (r"学术专长|专业特长|专长|特长", "specialty"),
    (r"技能爱好|技能特长|专业技能|技\s*能", "skills_summary"),
    (r"计算机熟练程度|计算机水平|计算机能力", "computer_proficiency"),
    (r"外语等级|外语熟练程度|外语水平|外语能力|外语程度", "foreign_language"),
    (r"原单位|原工作单位", "previous_employer"),
    (r"从事的岗位|现任岗位|现任职务", "current_position"),
    # professional_certificate BEFORE professional_title (more specific)
    (r"职称/职业资格证|专业职称|职业资格|职业资格证书|专业技术资格", "professional_certificate"),
    (r"职\s*称", "professional_title"),
    (r"社会经历|社会实践", "social_experience"),
    # ── Health & status ──
    (r"既往病史|病史|健康状况|健康情况|健康状\s*态", "medical_history"),
    (r"生育情况|生育状\s*态", "childbearing_status"),
    (r"是否应届|应届", "is_fresh_graduate"),
    # ── Emergency contact (before generic 联系人) ──
    (r"紧急联系人|紧急联系", "emergency_contact"),
    (r"关\s*系", "emergency_contact_relationship"),
    # ── Job seeking ──
    (r"求职意向|意向岗位", "expected_position"),
    (r"应聘二级学院|应聘院系|应聘学院|应聘部门", "applied_college"),
    (r"应聘岗位|应聘的岗位|应聘岗\s*位", "applied_position"),
    (r"岗\s*位", "applied_position"),
    (r"期望岗位|期望职位", "expected_position"),
    (r"期望月薪|期望薪资|期望薪酬|薪资要求|月薪要求", "expected_salary"),
    (r"期望城市|意向城市|期望工作地点", "expected_city"),
    (r"到岗时间|可入职时间|可到岗时间", "availability"),
    # ── Self evaluation ──
    (r"自我评价|自我介绍|个人简介", "self_evaluation"),
    # ── Spouse info (specific before generic) ──
    (r"配偶姓名|爱人姓名", "spouse_name"),
    (r"配偶出生年月|配偶生日|配偶出生日期", "spouse_birth_date"),
    (r"配偶工作单位|配偶单位", "spouse_work_unit"),
    (r"配偶电话|配偶联系", "spouse_phone"),
    (r"配偶籍贯", "spouse_hometown"),
    (r"配偶学历|配偶学位|配偶学历/学位", "spouse_education"),
    (r"配偶职称", "spouse_professional_title"),
    (r"配偶性别", "spouse_gender"),
    # ── Additional form fields ──
    (r"最高学位|最高学历", "highest_degree"),
    (r"主要科研成果|科研成果", "research_achievements"),
    (r"备\s*注", "remarks"),
    # ── Extra fields from various universities ──
    (r"出生地|出生地点", "birthplace"),
    (r"参加工作时间|参加工作年月|起始工作时间", "work_start_date"),
    (r"国\s*籍|国籍", "nationality_country"),
    (r"人事档案存放单位|档案存放地|档案所在单位", "archive_location"),
    (r"保险公积金|社保公积金|保险缴纳情况", "insurance_status"),
    (r"招聘来源|招聘信息来源|获取招聘信息来源", "recruitment_source"),
    (r"推荐人|推荐人姓名", "referrer"),
    (r"可授课程|教授课程|能讲授课程|承担课程", "teachable_courses"),
    (r"普通话等级|普通话水平|普通话", "mandarin_level"),
    (r"外语种类|外语语种|外语类型", "foreign_language_type"),
    (r"用人类型|用工类型|用工形式", "employment_type"),
    (r"所在国家|所在国|国\s*别", "country"),
    (r"所在省市|所在省|所在城市|省/市|省市", "province_city"),
    (r"承担导师情况|导师情况", "advisor_status"),
    (r"已培养硕士|已指导硕士|指导硕士数", "masters_supervised"),
    (r"已培养博士|已指导博士|指导博士数", "doctors_supervised"),
    (r"应聘岗位类型|岗位类型", "position_type"),
    (r"岗位名称|岗位号|岗位编号", "position_code"),
    (r"所属学院|所属院系", "affiliated_college"),
    (r"所属教研室|教研室", "affiliated_department"),
    (r"来校工作时间|到校时间", "arrival_date"),
    (r"学术兼职|学术职务|学术兼职情况", "academic_positions"),
    (r"人才称号|人才项目|人才称号项目", "talent_title"),
    (r"现聘岗位等级|现任岗位等级|现岗位等级", "current_position_level"),
    (r"是否退休|退休", "is_retired"),
    (r"诚信承诺|本人承诺", "integrity_commitment"),
    (r"奖惩情况|奖惩", "awards_punishments"),
    (r"处分|受过处分|处分情况", "disciplinary_record"),
    # ── Summary fields (for paragraph-style templates) ──
    (r"毕业院校|学校名称|院校", "education_summary"),
    (r"证书|资格证书|职业证书", "certificates_summary"),
    (r"语言能力|外语水平|语言", "languages_summary"),
    (r"兴趣爱好|爱\s*好", "hobbies_summary"),
]

# Override mappings for spouse section: when inside "配偶" section,
# these patterns take priority over the general FIELD_MAPPINGS
SPOUSE_FIELD_OVERRIDES: list[tuple[str, str]] = [
    (r"姓\s*名", "spouse_name"),
    (r"出生年月|出生日期|生日|出生时间", "spouse_birth_date"),
    (r"籍\s*贯", "spouse_hometown"),
    (r"学历/学位|学历|学位", "spouse_education"),
    (r"职\s*称", "spouse_professional_title"),
    (r"工作单位|单位", "spouse_work_unit"),
    (r"电话|联系电话|手机|联系方式", "spouse_phone"),
    (r"有无既往病史|病史", "medical_history"),
    (r"性\s*别", "spouse_gender"),
]

# Override mappings for emergency contact section
EMERGENCY_CONTACT_OVERRIDES: list[tuple[str, str]] = [
    (r"姓\s*名|紧急联系人|紧急联系", "emergency_contact"),
    (r"电话|联系人电话|联系电话|手机|联系方式", "emergency_contact_phone"),
    (r"关\s*系", "emergency_contact_relationship"),
]

# Section header patterns: identify multi-row table sections
# Maps section header regex -> (data_prefix, is_spouse_section)
SECTION_PATTERNS: dict[str, str] = {
    r"教育经历|学习经历|教育背景|学历背景|学习简历|教育情况": "education",
    r"工作经历|工作经验|工作背景|职业经历|工作简历|工作经历/实践活动": "work",
    r"项目经验|项目经历|科研项目|主持的主要科研项目|科研及成果": "project",
    r"发表论文|论文列表|学术成果|发表的论文|发表论文及出版著作": "publication",
    r"获奖情况|荣誉奖项|获奖经历": "award",
    r"家庭成员|家庭情况|主要社会关系|家庭关系|配偶及子女": "family",
    r"社会经历|社会实践": "social",
}

# Patterns that mark section boundaries for context tracking
SECTION_BOUNDARY_PATTERNS: list[tuple[str, str]] = [
    (r"紧急联系人|紧急联系", "emergency_contact"),
    (r"配偶及子女|配偶情况|家庭情况|家庭成员|家庭关系|主要社会关系", "spouse"),
    (r"教育经历|学习经历|教育背景|学历背景|学习简历|教育情况", "education"),
    (r"工作经历|工作经验|工作背景|职业经历|工作简历|工作经历/实践活动", "work"),
    (r"项目经验|项目经历|科研项目|主持的主要科研项目|科研及成果", "project"),
    (r"发表论文|论文列表|学术成果|发表的论文|发表论文及出版著作", "publication"),
    (r"获奖情况|荣誉奖项|获奖经历", "award"),
    (r"主要科研成果|科研成果|学术专长", "research"),
    (r"备\s*注", "remarks"),
    (r"来源", "source"),
    (r"基本资料|基本情况|个人信息|基本信息", "personal"),
]

# Column header patterns: map column header text to field suffix
COLUMN_MAPPINGS: dict[str, dict[str, str]] = {
    "education": {
        # date_range before generic 时间
        r"起止时间|起讫时间|在校时间|起止年月|"
        r"学习时间|就读时间|从何时至何时|何年何月|时\s*间|年月": "date_range",
        r"毕业院校|院校名称|学校名称|毕业学校|就读院校|"
        r"就读学校|何学校|院\s*校|院\s*校|学\s*校": "school",
        r"毕业专业|专\s*业|所学专业|专业名称|学习专业|专业方向": "major",
        r"学\s*历|学历层次": "degree",
        r"学\s*位|学位名称|学位类型": "degree",
        r"研究方向|方向|研究方\s*向|课题方向": "research_direction",
        r"学习层次|层次|培养层次|学历层次": "education_level",
        r"学习形式|办学形式|培养方式|就读形式": "education_form",
        r"证明人电话|证明电话": "reference_phone",
        r"证明人|证人": "reference_person",
        r"GPA|成绩|学业绩点|成绩绩点": "gpa",
        r"备注|说明|描述|备注说明": "description",
    },
    "work": {
        # date_range before generic 时间
        r"起止时间|起讫时间|工作时间|起止年月|"
        r"何年何月|从何时至何时|任职时间|时\s*间": "date_range",
        r"工作单位|单位名称|任何单位|任职单位|"
        r"就职单位|所在单位|何单位|单\s*位|公司": "company",
        r"岗位/职务|职\s*位|职务|担任的职务|任职|"
        r"工作职务|任何职务|岗\s*位": "position",
        r"职称|专业技术职务|专业技术职称|技术职称|任职职称": (
            "professional_title"
        ),
        r"部门|院系|科室|所在部门": "department",
        # 证明人电话 BEFORE 证明人 (specific before generic)
        r"证明人电话|证明电话": "reference_phone",
        r"证明人|证人": "reference_person",
        r"工作内容|职责|描述|备注|工作内容描述|教学专业|"
        r"任教专业|教授课程|教学内容|工作职责|主要工作": "description",
    },
    "project": {
        r"起止时间|时间|项目时间|起止年月|年月|项目起止": "date_range",
        r"项目名称|名称|课题名称|项目": "name",
        r"角\s*色|职务|承担工作|主持|参与|"
        r"承担角色|项目角色|担任角色": "role",
        r"技术|技术栈|使用技术|技术路线|研究方法": "technologies",
        r"经费来源|来源|资金来源|资助来源": "funding_source",
        r"经费|金额|经费金额|项目经费|资助金额": "funding_amount",
        r"描述|内容|备注|项目描述|项目内容|研究内容": "description",
    },
    "publication": {
        r"序号|编号|No": "index",
        r"论文题目|题目|名称|论文标题|文章题目|"
        r"文章标题|论文名称": "title",
        r"期刊|发表刊物|杂志|发表期刊|刊物名称|"
        r"期刊名称|刊载": "journal",
        r"时间|日期|发表时间|发表日期|刊登时间": "date",
        r"作者|作者顺序|署名|第几作者|作者排名": "authors",
        r"收录|收录情况|SCI|EI|索引|收录类别|收录类型": "index",
    },
    "award": {
        r"时间|日期|获奖时间|获奖日期|获奖年月": "date",
        r"奖项名称|名称|获奖|奖励名称|荣誉名称|获奖名称": "title",
        r"级别|等级|获奖级别|奖励级别": "level",
        r"颁发单位|授予单位|颁奖单位|授奖单位|颁发机构": "issuer",
    },
    "family": {
        r"姓名|称谓|姓名称谓|家\s*属\s*姓\s*名": "name",
        r"关系|与本人关系|亲属关系|家属关系": "relationship",
        r"性\s*别": "gender",
        r"出生年月|出生日期|年龄|出生时间": "birth_date",
        r"电话|联系电话|手机|联系方式": "phone",
        r"工作单位|单位|任职单位|所在单位": "work_unit",
        r"职务|岗位|职务岗位|工作职务": "position",
        r"地址|家庭住址|住址|联系地址": "address",
        r"政治面貌|党派": "political_status",
    },
    "social": {
        r"起止时间|时间|起止年月|年月": "date_range",
        r"内容|经历|描述|社会活动|实践内容": "description",
        r"地点|单位|活动地点|实践地点": "location",
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
        self._spouse_compiled = [
            (re.compile(p, re.IGNORECASE), k)
            for p, k in SPOUSE_FIELD_OVERRIDES
        ]
        self._emergency_contact_compiled = [
            (re.compile(p, re.IGNORECASE), k)
            for p, k in EMERGENCY_CONTACT_OVERRIDES
        ]
        self._boundary_compiled = [
            (re.compile(p, re.IGNORECASE), ctx)
            for p, ctx in SECTION_BOUNDARY_PATTERNS
        ]

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
        """Fill one section within a table, starting at the given header row.

        Scans forward up to 5 rows after the section header to find
        the column header row, handling templates with blank rows
        between the section header and column headers.
        """

        # Find the column header row by scanning forward
        col_header_r_idx = None
        column_map = {}

        # Check rows header_r_idx through header_r_idx + 5
        max_scan = min(header_r_idx + 6, len(table.rows))
        best_match_count = 0

        for r_idx in range(header_r_idx, max_scan):
            row = table.rows[r_idx]
            candidate_map = self._detect_column_mappings(row, prefix)
            match_count = len(candidate_map)
            if match_count > best_match_count:
                best_match_count = match_count
                col_header_r_idx = r_idx
                column_map = candidate_map
            if match_count >= 2:
                # Good enough — found the column header row
                break

        if not column_map or best_match_count == 0:
            return

        # Determine how many entries to fill
        count_key = f"{prefix}_count"
        if prefix == "work":
            count_key = "work_experience_count"
        elif prefix == "family":
            count_key = "family_members_count"
        elif prefix == "social":
            data_count = 1
            count_key = ""
        data_count = int(data.get(count_key, "0")) if count_key else data_count
        if data_count == 0:
            return

        # Fill data rows after the column header row
        data_start_r_idx = col_header_r_idx + 1

        # Don't fill into the next section's rows
        next_section_r_idx = len(table.rows)
        for r_idx in range(data_start_r_idx, len(table.rows)):
            row_text = " ".join(
                cell.text.strip() for cell in table.rows[r_idx].cells
            ).strip()
            for pattern, other_prefix in self._section_compiled.items():
                if pattern.search(row_text) and other_prefix != prefix:
                    next_section_r_idx = r_idx
                    break

        available_rows = next_section_r_idx - data_start_r_idx

        # Track filled cells to avoid duplicate writes in merged regions
        filled_cells: set[tuple[int, int]] = set()

        for entry_idx in range(1, min(data_count, available_rows) + 1):
            target_r_idx = data_start_r_idx + entry_idx - 1
            row = table.rows[target_r_idx]
            self._fill_row_by_column_map(
                row, column_map, prefix, entry_idx, data, filled_cells
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
        filled_cells: set[tuple[int, int]] | None = None,
    ) -> None:
        """Fill a single row based on the column mapping.

        For each column, look up the corresponding field value:
        - "date_range" -> combine {prefix}_{idx}_start_date and _end_date
        - "index" -> entry number
        - Other suffixes -> {prefix}_{idx}_{suffix}
        """
        if filled_cells is None:
            filled_cells = set()

        for c_idx, suffix in column_map.items():
            if c_idx >= len(row.cells):
                continue

            cell = row.cells[c_idx]

            # Skip cells that already have content (don't overwrite)
            if cell.text.strip():
                continue

            # Skip cells we've already filled (merged cell dedup)
            cell_id = (id(row), c_idx)
            if cell_id in filled_cells:
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
                filled_cells.add(cell_id)

    # ──────────────────── Phase 3: Single-Value Labels ────────────────────

    def _fill_single_value_labels(self, doc: Document, data: dict[str, str]) -> None:
        """Smart-detect label cells and fill adjacent blanks (multi-direction).

        Context-aware: when inside a "配偶" section, uses spouse-specific
        field mappings instead of general ones.
        When inside "紧急联系人" section, uses emergency contact mappings.
        """
        for table in doc.tables:
            # Track which section context we're in
            current_context: str = "personal"  # default
            filled_cells: set[tuple[int, int]] = set()

            for r_idx, row in enumerate(table.rows):
                cells = row.cells
                row_text = " ".join(c.text.strip() for c in cells).strip()

                # Check if this row is a section boundary
                for pattern, ctx in self._boundary_compiled:
                    if pattern.search(row_text):
                        current_context = ctx
                        break

                for c_idx, cell in enumerate(cells):
                    text = cell.text.strip()
                    if not text:
                        continue

                    # Skip already filled cells
                    cell_id = (id(row), c_idx)
                    if cell_id in filled_cells:
                        continue

                    # Determine which mapping set to use based on context
                    if current_context == "spouse":
                        mapping_set = self._spouse_compiled + self._compiled
                    elif current_context == "emergency_contact":
                        mapping_set = (
                            self._emergency_contact_compiled + self._compiled
                        )
                    else:
                        mapping_set = self._compiled

                    for pattern, key in mapping_set:
                        if not pattern.search(text):
                            continue

                        # For spouse/emergency context, skip if general
                        # mapping already matched and value is empty
                        value = data.get(key, "")
                        if not value:
                            continue

                        # Skip if cell already contains exactly this value
                        # (avoid substring false positives)
                        stripped_text = cell.text.strip()
                        if (
                            stripped_text.endswith(value)
                            and len(stripped_text) <= len(value) + 5
                        ):
                            break

                        # Strategy 1: Adjacent right cell
                        if c_idx + 1 < len(cells):
                            next_cell = cells[c_idx + 1]
                            next_id = (id(row), c_idx + 1)
                            if not next_cell.text.strip() and next_id not in filled_cells:
                                self._set_cell_text(next_cell, value)
                                filled_cells.add(next_id)
                                break

                        # Strategy 2: Cell below (next row, same column)
                        if r_idx + 1 < len(table.rows):
                            below_row = table.rows[r_idx + 1]
                            below_cell = below_row.cells[c_idx]
                            below_id = (id(below_row), c_idx)
                            if (not below_cell.text.strip()
                                    and below_id not in filled_cells):
                                self._set_cell_text(below_cell, value)
                                filled_cells.add(below_id)
                                break

                        # Strategy 3: Same cell with "label:___" format
                        if text.endswith(":") or text.endswith("："):
                            new_text = text + " " + value
                            self._set_cell_text(cell, new_text)
                            filled_cells.add(cell_id)
                            break

                        # Strategy 4: Same cell with "label：" + empty space
                        if "：" in text or ":" in text:
                            new_text = re.sub(
                                r"[:：]\s*$", f"：{value}", text
                            )
                            if new_text != text:
                                self._set_cell_text(cell, new_text)
                                filled_cells.add(cell_id)
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
