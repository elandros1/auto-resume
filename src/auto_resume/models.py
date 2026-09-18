"""Data models for resume information.

Supports comprehensive fields required by Chinese university recruitment forms:
- Basic info (name, gender, birth, ID, ethnicity, political status, etc.)
- Physical info (height, vision, marital status)
- Contact info (phone, email, address, postal code, hukou)
- Education (school, major, degree, level, form, reference person)
- Work experience (company, position, title, reference person)
- Research projects (name, role, period)
- Publications (title, journal, date, authors)
- Family members (name, relationship, phone, work unit)
- Skills, certificates, languages, hobbies
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class Education:
    """A single education entry."""
    school: str = ""
    major: str = ""
    degree: str = ""
    start_date: str = ""
    end_date: str = ""
    gpa: str = ""
    description: str = ""
    research_direction: str = ""   # 研究方向
    education_level: str = ""      # 学习层次 (本科/硕士/博士)
    education_form: str = ""       # 办学形式 (全日制/非全日制/成人/网络教育)
    reference_person: str = ""     # 证明人
    reference_phone: str = ""      # 证明人电话
    duration: str = ""             # 学制

    def to_text(self) -> str:
        parts = [f"{self.school} | {self.major} | {self.degree}"]
        if self.start_date or self.end_date:
            parts.append(f"{self.start_date} - {self.end_date}")
        if self.research_direction:
            parts.append(f"方向: {self.research_direction}")
        if self.education_level:
            parts.append(self.education_level)
        if self.education_form:
            parts.append(self.education_form)
        if self.gpa:
            parts.append(f"GPA: {self.gpa}")
        if self.description:
            parts.append(self.description)
        return "  ".join(parts)


@dataclass
class WorkExperience:
    """A single work experience entry."""
    company: str = ""
    position: str = ""
    start_date: str = ""
    end_date: str = ""
    department: str = ""
    description: str = ""
    reference_person: str = ""     # 证明人
    reference_phone: str = ""      # 证明人电话
    professional_title: str = ""  # 职称 (教授/副教授/讲师等)

    def to_text(self) -> str:
        parts = [f"{self.company} | {self.position}"]
        if self.department:
            parts[0] += f" | {self.department}"
        if self.professional_title:
            parts.append(f"职称: {self.professional_title}")
        if self.start_date or self.end_date:
            parts.append(f"{self.start_date} - {self.end_date}")
        if self.description:
            parts.append(self.description)
        return "  ".join(parts)


@dataclass
class Project:
    """A single project entry."""
    name: str = ""
    role: str = ""
    start_date: str = ""
    end_date: str = ""
    description: str = ""
    technologies: str = ""
    funding_source: str = ""      # 经费来源
    funding_amount: str = ""      # 经费金额

    def to_text(self) -> str:
        parts = [f"{self.name}"]
        if self.role:
            parts.append(f"| {self.role}")
        if self.start_date or self.end_date:
            parts.append(f"| {self.start_date} - {self.end_date}")
        if self.technologies:
            parts.append(f"| Tech: {self.technologies}")
        if self.funding_source:
            parts.append(f"| {self.funding_source}")
        if self.description:
            parts.append(f"| {self.description}")
        return " ".join(parts)


@dataclass
class Publication:
    """A single publication entry."""
    title: str = ""               # 论文题目
    journal: str = ""             # 期刊/发表刊物
    date: str = ""                # 发表时间
    authors: str = ""             # 作者及顺序
    index: str = ""               # 收录情况 (SCI/EI/核心等)

    def to_text(self) -> str:
        parts = [self.title]
        if self.journal:
            parts.append(f"| {self.journal}")
        if self.date:
            parts.append(f"| {self.date}")
        if self.authors:
            parts.append(f"| {self.authors}")
        if self.index:
            parts.append(f"| {self.index}")
        return " ".join(parts)


@dataclass
class Award:
    """A single award entry."""
    date: str = ""                # 获奖时间
    title: str = ""               # 奖项名称
    level: str = ""               # 级别 (国家级/省级/市级)
    issuer: str = ""              # 颁发单位

    def to_text(self) -> str:
        parts = []
        if self.date:
            parts.append(self.date)
        if self.title:
            parts.append(self.title)
        if self.level:
            parts.append(self.level)
        if self.issuer:
            parts.append(self.issuer)
        return " | ".join(parts)


@dataclass
class FamilyMember:
    """A single family member entry."""
    name: str = ""                # 姓名
    relationship: str = ""        # 与本人关系
    gender: str = ""              # 性别
    birth_date: str = ""          # 出生年月
    phone: str = ""               # 电话
    work_unit: str = ""           # 工作单位
    position: str = ""            # 职务
    address: str = ""             # 地址
    political_status: str = ""    # 政治面貌

    def to_text(self) -> str:
        parts = [self.name]
        if self.relationship:
            parts.append(f"({self.relationship})")
        if self.gender:
            parts.append(self.gender)
        if self.phone:
            parts.append(self.phone)
        if self.work_unit:
            parts.append(self.work_unit)
        return " ".join(parts)


@dataclass
class Skill:
    """A skill category."""
    category: str = ""
    items: str = ""


@dataclass
class ResumeData:
    """Complete resume data structure.

    Supports all fields commonly required by Chinese university
    recruitment application forms.
    """
    # ── Basic personal info ──
    name: str = ""
    gender: str = ""
    birth_date: str = ""
    age: str = ""                         # 年龄
    phone: str = ""
    email: str = ""
    address: str = ""
    postal_code: str = ""                 # 邮编
    hometown: str = ""                   # 籍贯
    hukou_location: str = ""             # 户口所在地
    political_status: str = ""
    id_number: str = ""
    nationality: str = ""                 # 民族
    marital_status: str = ""             # 婚姻状况
    height: str = ""                     # 身高
    vision: str = ""                      # 视力
    photo_path: str = ""

    # ── Professional info ──
    specialty: str = ""                  # 专业特长
    computer_proficiency: str = ""       # 计算机熟练程度
    foreign_language: str = ""            # 外语熟练程度
    previous_employer: str = ""          # 原单位
    current_position: str = ""           # 从事的岗位
    professional_title: str = ""         # 职称
    social_experience: str = ""          # 社会经历

    # ── Job seeking info ──
    job_intent: str = ""
    applied_position: str = ""           # 应聘岗位
    expected_position: str = ""
    expected_salary: str = ""
    expected_city: str = ""
    availability: str = ""

    # ── Self evaluation ──
    self_evaluation: str = ""

    # ── Spouse info (for forms that ask separately) ──
    spouse_name: str = ""
    spouse_birth_date: str = ""
    spouse_work_unit: str = ""
    spouse_phone: str = ""
    spouse_hometown: str = ""             # 配偶籍贯
    spouse_education: str = ""           # 配偶学历/学位
    spouse_professional_title: str = ""  # 配偶职称
    spouse_gender: str = ""              # 配偶性别

    # ── Additional form fields ──
    highest_degree: str = ""            # 最高学位
    medical_history: str = ""           # 既往病史/健康状况
    research_achievements: str = ""     # 主要科研成果
    remarks: str = ""                   # 备注

    # ── Emergency contact ──
    emergency_contact: str = ""         # 紧急联系人
    emergency_contact_phone: str = ""   # 紧急联系人电话
    emergency_contact_relationship: str = ""  # 紧急联系人关系

    # ── Extra fields from various universities ──
    birthplace: str = ""                # 出生地
    work_start_date: str = ""          # 参加工作时间
    nationality_country: str = ""      # 国籍
    archive_location: str = ""         # 人事档案存放单位
    insurance_status: str = ""         # 保险公积金
    recruitment_source: str = ""       # 招聘来源
    referrer: str = ""                 # 推荐人
    teachable_courses: str = ""        # 可授课程
    professional_certificate: str = "" # 职称/职业资格证
    mandarin_level: str = ""           # 普通话等级
    foreign_language_type: str = ""    # 外语种类
    employment_type: str = ""          # 用工类型
    country: str = ""                  # 所在国家
    province_city: str = ""            # 所在省市
    advisor_status: str = ""           # 承担导师情况
    masters_supervised: str = ""       # 已培养硕士数
    doctors_supervised: str = ""       # 已培养博士数
    position_type: str = ""            # 应聘岗位类型
    position_code: str = ""            # 岗位名称/岗位号
    affiliated_college: str = ""       # 所属学院
    affiliated_department: str = ""    # 所属教研室
    arrival_date: str = ""             # 来校工作时间
    academic_positions: str = ""      # 学术兼职
    talent_title: str = ""            # 人才称号
    current_position_level: str = ""  # 现聘岗位等级
    is_retired: str = ""               # 是否退休
    integrity_commitment: str = ""     # 诚信承诺
    awards_punishments: str = ""       # 奖惩情况
    disciplinary_record: str = ""     # 处分情况
    childbearing_status: str = ""      # 生育情况
    is_fresh_graduate: str = ""        # 是否应届
    applied_college: str = ""          # 应聘二级学院
    signature: str = ""                # 签名
    education_duration: str = ""      # 学制
    graduation_date: str = ""         # 毕业时间
    awards_summary: str = ""          # 个人主要荣誉及获奖
    education_summary: str = ""       # 教育经历概述
    work_experience_summary: str = ""  # 工作经历概述
    skills_summary: str = ""          # 专业技能概述
    certificates_summary: str = ""    # 证书概述
    work_years: str = ""              # 工作年限
    military_service: str = ""        # 参军情况
    info_channel: str = ""            # 招聘信息获取渠道

    # ── Lists (multi-entry) ──
    education: list[Education] = field(default_factory=list)
    work_experience: list[WorkExperience] = field(default_factory=list)
    projects: list[Project] = field(default_factory=list)
    publications: list[Publication] = field(default_factory=list)
    awards: list[Award] = field(default_factory=list)
    family_members: list[FamilyMember] = field(default_factory=list)
    skills: list[Skill] = field(default_factory=list)
    certificates: list[str] = field(default_factory=list)
    languages: list[str] = field(default_factory=list)
    hobbies: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ResumeData:
        """Create ResumeData from a dictionary."""
        resume = cls()

        # Simple string fields
        simple_fields = [
            # Basic info
            "name", "gender", "birth_date", "age", "phone", "email",
            "address", "postal_code", "hometown", "hukou_location",
            "political_status", "id_number", "nationality",
            "marital_status", "height", "vision", "photo_path",
            # Professional info
            "specialty", "computer_proficiency", "foreign_language",
            "previous_employer", "current_position", "professional_title",
            "social_experience",
            # Job seeking
            "job_intent", "applied_position", "applied_college",
            "expected_position", "expected_salary", "expected_city",
            "availability",
            # Self evaluation
            "self_evaluation",
            # Spouse
            "spouse_name", "spouse_birth_date", "spouse_work_unit",
            "spouse_phone", "spouse_hometown", "spouse_education",
            "spouse_professional_title", "spouse_gender",
            # Additional form fields
            "highest_degree", "medical_history", "research_achievements",
            "remarks",
            # Emergency contact
            "emergency_contact", "emergency_contact_phone",
            "emergency_contact_relationship",
            # Extra fields from various universities
            "birthplace", "work_start_date", "nationality_country",
            "archive_location", "insurance_status", "recruitment_source",
            "referrer", "teachable_courses", "professional_certificate",
            "mandarin_level", "foreign_language_type", "employment_type",
            "country", "province_city", "advisor_status",
            "masters_supervised", "doctors_supervised", "position_type",
            "position_code", "affiliated_college", "affiliated_department",
            "arrival_date", "academic_positions", "talent_title",
            "current_position_level", "is_retired", "integrity_commitment",
            "awards_punishments", "disciplinary_record",
            "childbearing_status", "is_fresh_graduate",
            "signature", "education_duration", "graduation_date",
            "awards_summary", "education_summary", "work_experience_summary",
            "skills_summary", "certificates_summary",
            "work_years", "military_service", "info_channel",
        ]
        for f in simple_fields:
            if f in data:
                setattr(resume, f, str(data[f]))

        # List/object fields — gracefully handle strings and lists
        # Helper: filter dict keys to only valid dataclass fields
        def _filter_fields(d, cls):
            if not isinstance(d, dict):
                return {}
            valid = {f.name for f in __import__("dataclasses").fields(cls)}
            return {k: v for k, v in d.items() if k in valid}

        if "education" in data:
            val = data["education"]
            if isinstance(val, list):
                resume.education = [
                    Education(**_filter_fields(e, Education)) if isinstance(e, dict) else Education()
                    for e in val
                ]
        if "work_experience" in data:
            val = data["work_experience"]
            if isinstance(val, list):
                resume.work_experience = [
                    WorkExperience(**_filter_fields(w, WorkExperience)) if isinstance(w, dict) else WorkExperience()
                    for w in val
                ]
        if "projects" in data:
            val = data["projects"]
            if isinstance(val, list):
                resume.projects = [
                    Project(**_filter_fields(p, Project)) if isinstance(p, dict) else Project()
                    for p in val
                ]
        if "publications" in data:
            val = data["publications"]
            if isinstance(val, list):
                resume.publications = [
                    Publication(**_filter_fields(p, Publication)) if isinstance(p, dict) else Publication()
                    for p in val
                ]
        if "awards" in data:
            val = data["awards"]
            if isinstance(val, list):
                resume.awards = [
                    Award(**_filter_fields(a, Award)) if isinstance(a, dict) else Award(title=str(a))
                    for a in val
                ]
            elif isinstance(val, str):
                resume.awards_summary = val
        if "family_members" in data:
            val = data["family_members"]
            if isinstance(val, list):
                resume.family_members = [
                    FamilyMember(**_filter_fields(f, FamilyMember)) if isinstance(f, dict) else FamilyMember()
                    for f in val
                ]
        if "skills" in data:
            val = data["skills"]
            if isinstance(val, list):
                resume.skills = [
                    Skill(**s) if isinstance(s, dict) else Skill(name=str(s))
                    for s in val
                ]
        if "certificates" in data:
            val = data["certificates"]
            if isinstance(val, list):
                resume.certificates = [str(c) for c in val]
            elif isinstance(val, str):
                resume.certificates = [val]
        if "languages" in data:
            val = data["languages"]
            if isinstance(val, list):
                resume.languages = [str(lang) for lang in val]
            elif isinstance(val, str):
                resume.languages = [val]
        if "hobbies" in data:
            val = data["hobbies"]
            if isinstance(val, list):
                resume.hobbies = [str(h) for h in val]
            elif isinstance(val, str):
                resume.hobbies = [val]

        # ── Flat education keys: education_1_school, education_1_start, etc. ──
        # Supports both list format and flat key format in the same JSON
        for i in range(1, 6):
            prefix = f"education_{i}_"
            flat_data = {
                k.replace(prefix, "").replace("start", "start_date").replace(
                    "end", "end_date"
                ).replace("level", "education_level").replace(
                    "form", "education_form"
                ).replace("witness_phone", "reference_phone").replace(
                    "witness", "reference_person"
                ): v
                for k, v in data.items()
                if k.startswith(prefix)
            }
            if flat_data and any(flat_data.values()):
                # Only add if we don't already have this index from list format
                if len(resume.education) < i:
                    while len(resume.education) < i - 1:
                        resume.education.append(Education())
                    resume.education.append(Education(**flat_data))
                else:
                    # Update existing entry with flat keys
                    edu = resume.education[i - 1]
                    for k, v in flat_data.items():
                        if v and hasattr(edu, k):
                            setattr(edu, k, v)

        # ── Flat work keys: work_1_unit, work_1_position, etc. ──
        for i in range(1, 6):
            prefix = f"work_{i}_"
            flat_data: dict[str, str] = {}
            for k, v in data.items():
                if k.startswith(prefix):
                    field = k.replace(prefix, "")
                    # Map flat key names to WorkExperience attribute names
                    # Order matters: longer patterns first
                    field = field.replace("unit", "company").replace(
                        "start", "start_date"
                    ).replace("end", "end_date").replace(
                        "witness_phone", "reference_phone"
                    ).replace("witness", "reference_person")
                    flat_data[field] = v
            if flat_data and any(flat_data.values()):
                if len(resume.work_experience) < i:
                    while len(resume.work_experience) < i - 1:
                        resume.work_experience.append(WorkExperience())
                    resume.work_experience.append(WorkExperience(**flat_data))
                else:
                    work = resume.work_experience[i - 1]
                    for k, v in flat_data.items():
                        if v and hasattr(work, k):
                            setattr(work, k, v)

        # ── Flat family member keys: family_1_name, family_1_phone, etc. ──
        for i in range(1, 6):
            prefix = f"family_{i}_"
            flat_data = {}
            for k, v in data.items():
                if k.startswith(prefix):
                    field = k.replace(prefix, "")
                    flat_data[field] = v
            if flat_data and any(flat_data.values()):
                if len(resume.family_members) < i:
                    while len(resume.family_members) < i - 1:
                        resume.family_members.append(FamilyMember())
                    resume.family_members.append(FamilyMember(**flat_data))
                else:
                    fam = resume.family_members[i - 1]
                    for k, v in flat_data.items():
                        if v and hasattr(fam, k):
                            setattr(fam, k, v)

        return resume

    @classmethod
    def from_json(cls, path: str | Path) -> ResumeData:
        """Load resume data from a JSON file."""
        path = Path(path)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)

    def to_flat_dict(self) -> dict[str, str]:
        """Convert to a flat dictionary for template substitution.

        Creates keys like {{name}}, {{phone}}, {{education_1_school}}, etc.
        Lists are expanded into indexed entries and summary strings.
        """
        d: dict[str, str] = {}

        # ── Simple fields ──
        simple_fields = [
            "name", "gender", "birth_date", "age", "phone", "email",
            "address", "postal_code", "hometown", "hukou_location",
            "political_status", "id_number", "nationality",
            "marital_status", "height", "vision", "photo_path",
            "specialty", "computer_proficiency", "foreign_language",
            "previous_employer", "current_position", "professional_title",
            "social_experience",
            "job_intent", "applied_position", "applied_college",
            "expected_position", "expected_salary", "expected_city",
            "availability",
            "self_evaluation",
            "spouse_name", "spouse_birth_date", "spouse_work_unit",
            "spouse_phone", "spouse_hometown", "spouse_education",
            "spouse_professional_title", "spouse_gender",
            "highest_degree", "medical_history", "research_achievements",
            "remarks",
            "emergency_contact", "emergency_contact_phone",
            "emergency_contact_relationship",
            "birthplace", "work_start_date", "nationality_country",
            "archive_location", "insurance_status", "recruitment_source",
            "referrer", "teachable_courses", "professional_certificate",
            "mandarin_level", "foreign_language_type", "employment_type",
            "country", "province_city", "advisor_status",
            "masters_supervised", "doctors_supervised", "position_type",
            "position_code", "affiliated_college", "affiliated_department",
            "arrival_date", "academic_positions", "talent_title",
            "current_position_level", "is_retired", "integrity_commitment",
            "awards_punishments", "disciplinary_record",
            "childbearing_status", "is_fresh_graduate",
            "signature", "education_duration", "graduation_date",
            "awards_summary", "education_summary", "work_experience_summary",
            "skills_summary", "certificates_summary",
            "work_years", "military_service", "info_channel",
        ]

        for f in simple_fields:
            d[f] = getattr(self, f) or ""

        # ── Education ──
        d["education_count"] = str(len(self.education))
        edu_summary = "\n".join(e.to_text() for e in self.education)
        d["education_summary"] = edu_summary or self.education_summary
        for i, edu in enumerate(self.education, 1):
            d[f"education_{i}_school"] = edu.school
            d[f"education_{i}_major"] = edu.major
            d[f"education_{i}_degree"] = edu.degree
            d[f"education_{i}_start_date"] = edu.start_date
            d[f"education_{i}_end_date"] = edu.end_date
            d[f"education_{i}_gpa"] = edu.gpa
            d[f"education_{i}_description"] = edu.description
            d[f"education_{i}_research_direction"] = edu.research_direction
            d[f"education_{i}_education_level"] = edu.education_level
            d[f"education_{i}_education_form"] = edu.education_form
            d[f"education_{i}_reference_person"] = edu.reference_person
            d[f"education_{i}_reference_phone"] = edu.reference_phone
            d[f"education_{i}"] = edu.to_text()

        # ── Work experience ──
        d["work_experience_count"] = str(len(self.work_experience))
        work_summary = "\n".join(w.to_text() for w in self.work_experience)
        d["work_experience_summary"] = work_summary or self.work_experience_summary
        for i, work in enumerate(self.work_experience, 1):
            d[f"work_{i}_company"] = work.company
            d[f"work_{i}_position"] = work.position
            d[f"work_{i}_start_date"] = work.start_date
            d[f"work_{i}_end_date"] = work.end_date
            d[f"work_{i}_department"] = work.department
            d[f"work_{i}_description"] = work.description
            d[f"work_{i}_reference_person"] = work.reference_person
            d[f"work_{i}_reference_phone"] = work.reference_phone
            d[f"work_{i}_professional_title"] = work.professional_title
            d[f"work_{i}"] = work.to_text()

        # ── Projects ──
        d["projects_count"] = str(len(self.projects))
        d["projects_summary"] = "\n".join(p.to_text() for p in self.projects)
        for i, proj in enumerate(self.projects, 1):
            d[f"project_{i}_name"] = proj.name
            d[f"project_{i}_role"] = proj.role
            d[f"project_{i}_start_date"] = proj.start_date
            d[f"project_{i}_end_date"] = proj.end_date
            d[f"project_{i}_description"] = proj.description
            d[f"project_{i}_technologies"] = proj.technologies
            d[f"project_{i}_funding_source"] = proj.funding_source
            d[f"project_{i}_funding_amount"] = proj.funding_amount
            d[f"project_{i}"] = proj.to_text()

        # ── Publications ──
        d["publications_count"] = str(len(self.publications))
        d["publications_summary"] = "\n".join(p.to_text() for p in self.publications)
        for i, pub in enumerate(self.publications, 1):
            d[f"publication_{i}_title"] = pub.title
            d[f"publication_{i}_journal"] = pub.journal
            d[f"publication_{i}_date"] = pub.date
            d[f"publication_{i}_authors"] = pub.authors
            d[f"publication_{i}_index"] = pub.index
            d[f"publication_{i}"] = pub.to_text()

        # ── Awards ──
        d["awards_count"] = str(len(self.awards))
        d["awards_summary"] = "\n".join(a.to_text() for a in self.awards)
        for i, award in enumerate(self.awards, 1):
            d[f"award_{i}_date"] = award.date
            d[f"award_{i}_title"] = award.title
            d[f"award_{i}_level"] = award.level
            d[f"award_{i}_issuer"] = award.issuer
            d[f"award_{i}"] = award.to_text()

        # ── Family members ──
        d["family_members_count"] = str(len(self.family_members))
        d["family_members_summary"] = "\n".join(f.to_text() for f in self.family_members)
        for i, member in enumerate(self.family_members, 1):
            d[f"family_{i}_name"] = member.name
            d[f"family_{i}_relationship"] = member.relationship
            d[f"family_{i}_gender"] = member.gender
            d[f"family_{i}_birth_date"] = member.birth_date
            d[f"family_{i}_phone"] = member.phone
            d[f"family_{i}_work_unit"] = member.work_unit
            d[f"family_{i}_position"] = member.position
            d[f"family_{i}_address"] = member.address
            d[f"family_{i}_political_status"] = member.political_status
            d[f"family_{i}"] = member.to_text()

        # ── Skills ──
        d["skills_count"] = str(len(self.skills))
        skills_sum = "\n".join(
            f"{s.category}: {s.items}" for s in self.skills
        )
        d["skills_summary"] = skills_sum or self.skills_summary
        for i, skill in enumerate(self.skills, 1):
            d[f"skill_{i}_category"] = skill.category
            d[f"skill_{i}_items"] = skill.items

        # ── Certificates ──
        d["certificates_count"] = str(len(self.certificates))
        certs_sum = "、".join(self.certificates)
        d["certificates_summary"] = certs_sum or self.certificates_summary
        for i, cert in enumerate(self.certificates, 1):
            d[f"certificate_{i}"] = cert

        # ── Languages ──
        d["languages_count"] = str(len(self.languages))
        d["languages_summary"] = "、".join(self.languages)
        for i, lang in enumerate(self.languages, 1):
            d[f"language_{i}"] = lang

        # ── Hobbies ──
        d["hobbies_count"] = str(len(self.hobbies))
        d["hobbies_summary"] = "、".join(self.hobbies)
        for i, h in enumerate(self.hobbies, 1):
            d[f"hobby_{i}"] = h

        return d

    def to_dict(self) -> dict[str, Any]:
        """Convert to a serializable dictionary."""
        return asdict(self)

    def to_json(self, path: str | Path) -> None:
        """Save resume data to a JSON file."""
        path = Path(path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def sample(cls) -> ResumeData:
        """Create a sample ResumeData with example data for template generation."""
        data = {
            "name": "张三",
            "gender": "男",
            "birth_date": "1995-06",
            "age": "29",
            "phone": "13800138000",
            "email": "zhangsan@example.com",
            "address": "北京市海淀区中关村大街1号",
            "postal_code": "100080",
            "hometown": "浙江杭州",
            "hukou_location": "浙江省杭州市",
            "political_status": "中共党员",
            "id_number": "3301**********1234",
            "nationality": "汉",
            "marital_status": "未婚",
            "height": "175cm",
            "vision": "5.0/5.0",
            "specialty": "人工智能与机器学习，擅长自然语言处理和大模型训练",
            "computer_proficiency": "熟练使用Python、Java、C++，精通PyTorch、TensorFlow框架",
            "foreign_language": "英语（CET-6，流利，可作为工作语言）",
            "previous_employer": "清华大学计算机系",
            "current_position": "助理研究员",
            "professional_title": "助理研究员（中级职称）",
            "social_experience": "2020-2021年担任研究生会学术部部长",
            "job_intent": "高校教师",
            "applied_position": "计算机科学与技术 讲师/副教授",
            "expected_position": "计算机科学与技术 讲师/副教授",
            "expected_salary": "面议",
            "expected_city": "北京",
            "availability": "随时到岗",
            "self_evaluation": "博士毕业于985高校，发表SCI论文5篇，主持省级科研项目1项。",
            "emergency_contact": "李紧急",
            "emergency_contact_phone": "13900001111",
            "emergency_contact_relationship": "妻子",
            "highest_degree": "博士研究生",
            "graduation_date": "2024.06",
            "education_duration": "五年",
            "birthplace": "浙江杭州",
            "applied_college": "计算机学院",
            "professional_certificate": "高校教师资格证",
            "is_fresh_graduate": "否",
            "childbearing_status": "无",
            "education": [
                {
                    "school": "浙江大学",
                    "major": "计算机科学与技术",
                    "degree": "博士",
                    "start_date": "2019.09",
                    "end_date": "2024.06",
                    "gpa": "3.8/4.0",
                    "description": "研究方向：人工智能与机器学习",
                    "education_level": "博士研究生",
                    "education_form": "全日制",
                    "reference_person": "李教授",
                    "reference_phone": "0571-12345678",
                },
                {
                    "school": "浙江大学",
                    "major": "计算机科学与技术",
                    "degree": "学士",
                    "start_date": "2015.09",
                    "end_date": "2019.06",
                    "gpa": "3.9/4.0",
                    "education_level": "本科",
                    "education_form": "全日制",
                    "reference_person": "王教授",
                    "reference_phone": "0571-87654321",
                },
            ],
            "work_experience": [
                {
                    "company": "清华大学计算机系",
                    "position": "助理研究员",
                    "start_date": "2024.07",
                    "end_date": "至今",
                    "department": "人工智能研究所",
                    "description": "从事大模型训练优化研究",
                    "reference_person": "张院士",
                    "reference_phone": "010-12345678",
                    "professional_title": "助理研究员",
                },
            ],
            "projects": [
                {
                    "name": "基于大模型的知识图谱构建",
                    "role": "项目负责人",
                    "start_date": "2023.01",
                    "end_date": "2023.12",
                    "description": "利用LLM自动提取实体关系",
                    "funding_source": "省级科研基金",
                    "funding_amount": "30万元",
                },
            ],
            "publications": [
                {
                    "title": "Efficient Training of Large Language Models",
                    "journal": "ACL 2023",
                    "date": "2023.07",
                    "authors": "第一作者",
                    "index": "CCF-A",
                },
            ],
            "awards": [
                {"date": "2021.05", "title": "国家奖学金", "level": "国家级", "issuer": "教育部"},
            ],
            "family_members": [
                {
                    "name": "张父",
                    "relationship": "父亲",
                    "gender": "男",
                    "birth_date": "1965.03",
                    "phone": "13900001111",
                    "work_unit": "杭州某中学",
                    "position": "教师",
                    "political_status": "中共党员",
                },
            ],
            "skills": [
                {"category": "编程语言", "items": "Python, Java, C++, Go"},
                {"category": "框架工具", "items": "PyTorch, TensorFlow, Docker, Git"},
            ],
            "certificates": ["大学英语六级 (CET-6)", "高校教师资格证"],
            "languages": ["英语（流利）"],
            "hobbies": ["阅读", "跑步", "围棋"],
        }
        return cls.from_dict(data)
