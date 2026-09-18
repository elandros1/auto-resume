"""Fuzzy field matching layer for auto-detecting resume form fields.

This module provides automatic field recognition without requiring exact
regex patterns for every possible field name variation. It uses three
strategies in order:

1. **Similarity scoring** — difflib SequenceMatcher to compare the label
   text against a database of known field aliases. Handles typos, minor
   word order changes, and abbreviated forms.

2. **Semantic keyword mapping** — each data field has a set of Chinese
   keywords associated with its meaning. If a label contains those keywords,
   it matches even if the exact wording is completely different.

3. **Substring containment** — if a known alias is a substring of the
   label (or vice versa), it matches. Handles cases like "应聘二级学院"
   containing "应聘" and "学院".

Usage:
    from auto_resume.fuzzy_matcher import FuzzyMatcher
    matcher = FuzzyMatcher()
    key, score = matcher.match("联系电话")  # → ("phone", 0.95)
    key, score = matcher.match("移动手机号码")  # → ("phone", 0.80)
    key, score = matcher.match("完全没见过的字段")  # → (None, 0.0)
"""

from __future__ import annotations

from difflib import SequenceMatcher
from typing import NamedTuple


class MatchResult(NamedTuple):
    """Result of a fuzzy match attempt."""
    key: str | None       # Resume data key, or None if no match
    score: float          # Confidence score 0.0–1.0
    method: str           # "regex" | "similarity" | "semantic" | "substring" | "none"


# ──────────────────── Field Alias Database ────────────────────
# Each field has a list of known aliases (Chinese label variations).
# The fuzzy matcher compares input against these aliases.
# This is the "training data" — adding aliases improves accuracy.

FIELD_ALIASES: dict[str, list[str]] = {
    # ── Basic personal info ──
    "name": ["姓名", "名字", "姓名称谓", "姓 名"],
    "gender": ["性别", "男女"],
    "birth_date": ["出生年月", "出生日期", "生日", "出生时间", "出生", "年月日"],
    "age": ["年龄", "岁数", "年纪"],
    "phone": [
        "联系电话", "联系方式", "手机", "电话", "联络电话",
        "移动电话", "手机号码", "手机号", "电话号码", "联系",
    ],
    "email": [
        "电子邮箱", "邮箱", "Email", "E-mail", "电子邮件",
        "电子信箱", "Email地址", "电子邮箱地址", "邮件",
    ],
    "hometown": ["籍贯", "籍贯地", "出生地", "生源地", "原籍", "祖籍"],
    "political_status": ["政治面貌", "党派", "政治"],
    "id_number": ["身份证号", "证件号码", "身份证号码", "身份证明", "身份证", "证件号"],
    "nationality": ["民族", "族别", "族"],
    "marital_status": ["婚姻状况", "是否已婚", "婚否", "婚姻", "婚况"],
    "address": [
        "通讯地址", "联系地址", "地址", "住址", "家庭住址",
        "现住址", "居住地", "现通讯地址",
    ],
    "hukou_location": [
        "户口所在地", "户籍地", "户口", "户籍所在地", "户口所在",
        "现户籍地址", "户籍", "户口地址",
    ],
    "postal_code": ["邮编", "邮政编码", "邮编代码", "编码"],
    "height": ["身高", "身高cm", "身高CM"],
    "vision": ["视力", "眼睛度数", "视力情况"],
    "photo_path": ["照片", "相片", "免冠照", "证件照"],

    # ── Professional info ──
    "specialty": ["专业特长", "特长", "学术专长", "专长", "专业特长爱好"],
    "skills_summary": ["技能爱好", "技能特长", "专业技能", "技能", "爱好技能"],
    "computer_proficiency": ["计算机熟练程度", "计算机水平", "计算机能力", "计算机"],
    "foreign_language": [
        "外语熟练程度", "外语水平", "外语能力", "外语程度",
        "外语等级", "外语", "外语能力等级",
    ],
    "previous_employer": ["原单位", "原工作单位", "前单位"],
    "current_position": ["从事的岗位", "现任岗位", "现任职务", "现任"],
    "professional_title": ["职称", "专业技术职务", "技术职称", "任职职称"],
    "professional_certificate": [
        "职称/职业资格证", "专业职称", "职业资格", "职业资格证书",
        "专业技术资格", "资格证", "职业资格证",
    ],
    "social_experience": ["社会经历", "社会实践", "社会活动"],

    # ── Health & status ──
    "medical_history": [
        "既往病史", "病史", "健康状况", "健康情况", "健康",
        "健康状态",
    ],
    "childbearing_status": ["生育情况", "生育状态", "生育"],
    "is_fresh_graduate": ["是否应届", "应届", "应届生"],

    # ── Emergency contact ──
    "emergency_contact": ["紧急联系人", "紧急联系", "紧急"],
    "emergency_contact_phone": ["联系人电话", "紧急联系电话", "紧急电话", "紧急联系手机"],
    "emergency_contact_relationship": ["关系", "与本人关系", "亲属关系"],

    # ── Job seeking ──
    "expected_position": ["求职意向", "意向岗位", "期望岗位", "期望职位"],
    "applied_position": ["应聘岗位", "应聘的岗位", "岗位", "应聘", "应聘职位"],
    "applied_college": [
        "应聘二级学院", "应聘院系", "应聘学院", "应聘部门",
        "应聘单位", "二级学院",
    ],
    "expected_salary": ["期望月薪", "期望薪资", "期望薪酬", "薪资要求", "月薪要求", "月薪"],
    "expected_city": ["期望城市", "意向城市", "期望工作地点", "意向地点"],
    "availability": ["到岗", "入职", "上班", "可上班", "到岗时间"],

    # ── Self evaluation ──
    "self_evaluation": ["自我评价", "自我介绍", "个人简介", "个人评价", "自述", "自我"],

    # ── Spouse info ──
    "spouse_name": ["配偶姓名", "爱人姓名", "配偶名", "爱人"],
    "spouse_birth_date": ["配偶出生年月", "配偶生日", "配偶出生日期"],
    "spouse_work_unit": ["配偶工作单位", "配偶单位", "配偶工作"],
    "spouse_phone": ["配偶电话", "配偶联系", "配偶手机"],
    "spouse_hometown": ["配偶籍贯", "配偶出生地"],
    "spouse_education": ["配偶学历", "配偶学位", "配偶学历/学位"],
    "spouse_professional_title": ["配偶职称", "配偶职务"],
    "spouse_gender": ["配偶性别"],

    # ── Additional form fields ──
    "highest_degree": ["最高学位", "最高学历"],
    "research_achievements": ["主要科研成果", "科研成果", "科研成果及"],
    "remarks": ["备注", "说明", "附注"],

    # ── Extra university fields ──
    "birthplace": ["出生地", "出生地点", "出生地方"],
    "work_start_date": ["参加工作时间", "参加工作年月", "起始工作时间", "参加工作"],
    "nationality_country": ["国籍", "国别", "国家"],
    "archive_location": ["人事档案存放单位", "档案存放地", "档案所在单位", "档案"],
    "insurance_status": ["保险公积金", "社保公积金", "保险缴纳情况", "社保"],
    "recruitment_source": ["招聘来源", "招聘信息来源", "获取招聘信息来源", "信息来源"],
    "referrer": ["推荐人", "推荐人姓名", "推荐"],
    "teachable_courses": ["可授课程", "教授课程", "能讲授课程", "承担课程", "授课"],
    "mandarin_level": ["普通话等级", "普通话水平", "普通话"],
    "foreign_language_type": ["外语种类", "外语语种", "外语类型", "语种"],
    "employment_type": ["用人类型", "用工类型", "用工形式", "用工"],
    "country": ["所在国家", "所在国", "国别"],
    "province_city": ["所在省市", "所在省", "所在城市", "省/市", "省市"],
    "advisor_status": ["承担导师情况", "导师情况", "导师"],
    "masters_supervised": ["已培养硕士", "已指导硕士", "指导硕士数"],
    "doctors_supervised": ["已培养博士", "已指导博士", "指导博士数"],
    "position_type": ["应聘岗位类型", "岗位类型"],
    "position_code": ["岗位名称", "岗位号", "岗位编号"],
    "affiliated_college": ["所属学院", "所属院系", "所在学院"],
    "affiliated_department": ["所属教研室", "教研室", "所在教研室"],
    "arrival_date": ["来校工作时间", "到校时间", "到校"],
    "academic_positions": ["学术兼职", "学术职务", "学术兼职情况", "学术"],
    "talent_title": ["人才称号", "人才项目", "人才称号项目", "人才"],
    "current_position_level": ["现聘岗位等级", "现任岗位等级", "现岗位等级", "岗位等级"],
    "is_retired": ["是否退休", "退休"],
    "integrity_commitment": ["诚信承诺", "本人承诺", "承诺"],
    "awards_punishments": ["奖惩情况", "奖惩"],
    "disciplinary_record": ["处分", "受过处分", "处分情况"],
    "signature": ["签名", "签字", "手写签名"],
    "education_duration": ["学制", "学制年限"],
    "graduation_date": ["毕业时间", "毕业日期", "毕业年月"],
    "awards_summary": ["个人主要荣誉及获奖", "主要荣誉", "荣誉获奖"],

    # ── Summary fields ──
    "education_summary": ["毕业院校", "学校名称", "院校", "毕业学校"],
    "certificates_summary": ["证书", "资格证书", "职业证书", "资格"],
    "languages_summary": ["语言能力", "外语水平", "语言"],
    "hobbies_summary": ["兴趣爱好", "爱好", "特长爱好"],
}

# ──────────────────── Semantic Keyword Mapping ────────────────────
# Maps Chinese semantic keywords to field keys.
# If a label contains ANY of these keywords, it's a candidate match.
# This catches completely new wordings that similarity scoring would miss.

SEMANTIC_KEYWORDS: dict[str, list[str]] = {
    "name": ["姓名", "名字"],
    "gender": ["性别", "男", "女"],
    "birth_date": ["出生", "生日", "出生年月"],
    "age": ["年龄", "岁"],
    "phone": ["电话", "手机", "联系", "电话号码"],
    "email": ["邮箱", "邮件", "email", "E-mail", "电子"],
    "hometown": ["籍贯", "祖籍", "出生地", "生源地"],
    "political_status": ["政治", "党", "党员", "团员", "群众"],
    "id_number": ["身份证", "证件", "身份"],
    "nationality": ["民族", "族"],
    "marital_status": ["婚", "婚姻", "已婚", "未婚"],
    "address": ["地址", "住址", "通讯", "居住", "住"],
    "hukou_location": ["户口", "户籍"],
    "postal_code": ["邮编", "编码", "邮政"],
    "height": ["身高", "高"],
    "vision": ["视力", "眼", "度数"],
    "specialty": ["特长", "专长", "特"],
    "computer_proficiency": ["计算机", "电脑"],
    "foreign_language": ["外语", "英语", "语言能力"],
    "professional_title": ["职称", "职务"],
    "professional_certificate": ["资格证", "资格", "证书"],
    "medical_history": ["健康", "病史", "身体"],
    "childbearing_status": ["生育", "子女", "育有"],
    "emergency_contact": ["紧急", "应急"],
    "expected_salary": ["薪资", "月薪", "薪酬", "工资", "待遇"],
    "expected_position": ["意向", "期望岗位", "求职"],
    "applied_position": ["应聘", "岗位", "职位"],
    "applied_college": ["学院", "院系", "部门"],
    "self_evaluation": ["自我评价", "自我", "个人简介", "评价", "自述"],
    "remarks": ["备注", "说明", "附注"],
    "availability": ["到岗", "上班", "入职"],
    "research_achievements": ["科研成果", "科研"],
    "teachable_courses": ["课程", "授课", "教学"],
    "education_summary": ["毕业院校", "院校", "学校"],
    "is_fresh_graduate": ["应届", "毕业生"],
    "signature": ["签名", "签字"],
    "graduation_date": ["毕业时间", "毕业"],
    "awards_summary": ["荣誉", "获奖", "奖励"],
    "education_duration": ["学制"],
}

# ──────────────────── Column Header Aliases ────────────────────
# Aliases for multi-row table column headers, organized by section.

COLUMN_ALIASES: dict[str, dict[str, list[str]]] = {
    "education": {
        "date_range": [
            "起止时间", "时间", "起讫时间", "在校时间", "起止年月",
            "年月", "学习时间", "就读时间", "从何时至何时",
            "何年何月", "起止", "时间段",
        ],
        "school": [
            "院校名称", "学校名称", "院校", "毕业院校", "学校",
            "就读院校", "毕业学校", "就读学校", "何学校",
            "院 校", "学 校",
        ],
        "major": [
            "专业", "所学专业", "专业名称", "学习专业",
            "专业方向", "毕业专业",
        ],
        "degree": ["学历", "学位", "学历层次", "学位名称", "学位类型"],
        "education_duration": ["学制", "学制年限"],
        "research_direction": ["研究方向", "方向", "课题方向", "专业（方向）", "专业方向"],
        "education_level": ["学习层次", "层次", "培养层次", "学历层次"],
        "education_form": ["办学形式", "培养方式", "学习形式", "就读形式"],
        "reference_person": ["证明人", "证人"],
        "reference_phone": ["证明人电话", "证明电话", "证明人手机"],
        "gpa": ["GPA", "成绩", "绩点", "学业绩点"],
        "description": ["备注", "说明", "描述"],
    },
    "work": {
        "date_range": [
            "起止时间", "时间", "起讫时间", "工作时间", "起止年月",
            "年月", "何年何月", "从何时至何时", "任职时间", "起止",
        ],
        "company": [
            "工作单位", "单位名称", "单位", "公司", "任何单位",
            "任职单位", "就职单位", "所在单位", "何单位",
        ],
        "position": [
            "职位", "职务", "岗位", "担任的职务", "任职",
            "工作职务", "任何职务", "岗位/职务",
        ],
        "professional_title": ["职称", "专业技术职务", "技术职称", "任职职称"],
        "department": ["部门", "院系", "科室", "所在部门"],
        "reference_person": ["证明人", "证人"],
        "reference_phone": ["证明人电话", "证明电话", "证明人手机"],
        "description": [
            "工作内容", "职责", "描述", "备注", "工作内容描述",
            "教学专业", "任教专业", "教授课程", "教学内容",
            "工作职责", "主要工作",
        ],
    },
    "project": {
        "date_range": ["起止时间", "时间", "项目时间", "起止年月", "年月"],
        "name": ["项目名称", "名称", "课题名称", "项目"],
        "role": ["角色", "职务", "承担工作", "主持", "参与", "承担角色"],
        "technologies": ["技术", "技术栈", "使用技术", "技术路线", "研究方法"],
        "funding_source": ["经费来源", "来源", "资金来源", "资助来源"],
        "funding_amount": ["经费", "金额", "经费金额", "项目经费"],
        "description": ["描述", "内容", "备注", "项目描述", "研究内容"],
    },
    "publication": {
        "index": ["序号", "编号", "No"],
        "title": ["论文题目", "题目", "名称", "论文标题", "文章题目", "论文名称"],
        "journal": ["期刊", "发表刊物", "杂志", "发表期刊", "刊物名称", "刊载"],
        "date": ["时间", "日期", "发表时间", "发表日期", "刊登时间"],
        "authors": ["作者", "作者顺序", "署名", "第几作者", "作者排名"],
    },
    "award": {
        "date": ["时间", "日期", "获奖时间", "获奖日期"],
        "title": ["奖项名称", "名称", "获奖", "奖励名称", "荣誉名称"],
        "level": ["级别", "等级", "获奖级别"],
        "issuer": ["颁发单位", "授予单位", "颁奖单位", "授奖单位", "颁发机构"],
    },
    "family": {
        "name": ["姓名", "称谓", "姓名称谓", "家属姓名"],
        "relationship": ["关系", "与本人关系", "亲属关系", "家属关系"],
        "gender": ["性别"],
        "birth_date": ["出生年月", "出生日期", "年龄", "出生时间"],
        "phone": ["电话", "联系电话", "手机", "联系方式"],
        "work_unit": ["工作单位", "单位", "任职单位", "所在单位"],
        "position": ["职务", "岗位", "职务岗位", "工作职务"],
        "address": ["地址", "家庭住址", "住址", "联系地址"],
        "political_status": ["政治面貌", "党派"],
    },
    "social": {
        "date_range": ["起止时间", "时间", "起止年月", "年月"],
        "description": ["内容", "经历", "描述", "社会活动", "实践内容"],
        "location": ["地点", "单位", "活动地点", "实践地点"],
    },
}


class FuzzyMatcher:
    """Auto-detect resume field labels using fuzzy matching.

    Three-layer matching strategy:
    1. Substring containment (fast, catches partial matches)
    2. Similarity scoring (difflib SequenceMatcher, catches typos/variants)
    3. Semantic keywords (catches completely new wordings)

    Attributes:
        similarity_threshold: Minimum similarity ratio (0-1) for a match.
        semantic_threshold: Minimum number of keyword hits for a match.
    """

    def __init__(
        self,
        similarity_threshold: float = 0.6,
        semantic_threshold: int = 1,
    ):
        self.similarity_threshold = similarity_threshold
        self.semantic_threshold = semantic_threshold
        self._alias_cache: dict[str, list[str]] = FIELD_ALIASES
        self._semantic_cache: dict[str, list[str]] = SEMANTIC_KEYWORDS
        self._column_cache: dict[str, dict[str, list[str]]] = COLUMN_ALIASES

    def match(self, label: str) -> MatchResult:
        """Match a single field label to a resume data key.

        Args:
            label: The field label text from the Word document (e.g. "联系电话").

        Returns:
            MatchResult with the best matching key, score, and method.
        """
        label = label.strip()
        if not label:
            return MatchResult(None, 0.0, "none")

        # Layer 1: Substring containment (fastest)
        result = self._match_substring(label)
        if result.key:
            return result

        # Layer 2: Similarity scoring
        result = self._match_similarity(label)
        if result.key and result.score >= self.similarity_threshold:
            return result

        # Layer 3: Semantic keywords
        result = self._match_semantic(label)
        if result.key:
            return result

        return MatchResult(None, 0.0, "none")

    def match_column(
        self, header: str, section: str
    ) -> tuple[str | None, float]:
        """Match a table column header to a field suffix.

        Args:
            header: Column header text (e.g. "毕业院校").
            section: Section prefix (e.g. "education", "work").

        Returns:
            (field_suffix, score) or (None, 0.0).
        """
        header = header.strip()
        if not header:
            return None, 0.0

        col_aliases = self._column_cache.get(section, {})
        if not col_aliases:
            return None, 0.0

        best_suffix: str | None = None
        best_score = 0.0

        for suffix, aliases in col_aliases.items():
            for alias in aliases:
                # Substring check
                if alias in header or header in alias:
                    score = max(
                        len(alias) / len(header) if len(header) > 0 else 0,
                        len(header) / len(alias) if len(alias) > 0 else 0,
                    )
                    score = max(score, 0.85)  # boost substring matches
                else:
                    # Similarity
                    score = SequenceMatcher(None, alias, header).ratio()

                if score > best_score:
                    best_score = score
                    best_suffix = suffix

        if best_score >= self.similarity_threshold:
            return best_suffix, best_score
        return None, 0.0

    # ──────────────────── Private Methods ────────────────────

    def _match_substring(self, label: str) -> MatchResult:
        """Check if any known alias is a substring of the label (or vice versa)."""
        best_key = None
        best_score = 0.0

        for key, aliases in self._alias_cache.items():
            for alias in aliases:
                if alias in label:
                    # Alias is part of label — score based on coverage
                    score = len(alias) / len(label) if len(label) > 0 else 1.0
                    score = max(score, 0.85)  # boost: substring is strong signal
                    if score > best_score:
                        best_score = score
                        best_key = key
                elif label in alias:
                    # Label is part of alias (e.g. "电话" in "联系电话")
                    score = len(label) / len(alias) if len(alias) > 0 else 0.5
                    if score > best_score:
                        best_score = score
                        best_key = key

        if best_key:
            return MatchResult(best_key, best_score, "substring")
        return MatchResult(None, 0.0, "none")

    def _match_similarity(self, label: str) -> MatchResult:
        """Use difflib SequenceMatcher to find the closest alias."""
        best_key = None
        best_score = 0.0

        for key, aliases in self._alias_cache.items():
            for alias in aliases:
                score = SequenceMatcher(None, alias, label).ratio()
                if score > best_score:
                    best_score = score
                    best_key = key

        if best_key and best_score > 0:
            return MatchResult(best_key, best_score, "similarity")
        return MatchResult(None, 0.0, "none")

    def _match_semantic(self, label: str) -> MatchResult:
        """Check if the label contains semantic keywords for any field."""
        best_key = None
        best_hits = 0

        for key, keywords in self._semantic_cache.items():
            hits = sum(1 for kw in keywords if kw in label)
            if hits > best_hits:
                best_hits = hits
                best_key = key

        if best_key and best_hits >= self.semantic_threshold:
            # Score based on keyword coverage
            score = min(best_hits * 0.3, 0.9)
            return MatchResult(best_key, score, "semantic")
        return MatchResult(None, 0.0, "none")
