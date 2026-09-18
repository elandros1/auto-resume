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
    education_level: str = ""      # 学习层次 (本科/硕士/博士)
    education_form: str = ""       # 办学形式 (全日制/非全日制/成人/网络教育)
    reference_person: str = ""     # 证明人
    reference_phone: str = ""      # 证明人电话

    def to_text(self) -> str:
        parts = [f"{self.school} | {self.major} | {self.degree}"]
        if self.start_date or self.end_date:
            parts.append(f"{self.start_date} - {self.end_date}")
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
            "job_intent", "applied_position", "expected_position",
            "expected_salary", "expected_city", "availability",
            # Self evaluation
            "self_evaluation",
            # Spouse
            "spouse_name", "spouse_birth_date", "spouse_work_unit",
            "spouse_phone",
        ]
        for f in simple_fields:
            if f in data:
                setattr(resume, f, str(data[f]))

        # List/object fields
        if "education" in data:
            resume.education = [Education(**e) for e in data["education"]]
        if "work_experience" in data:
            resume.work_experience = [WorkExperience(**w) for w in data["work_experience"]]
        if "projects" in data:
            resume.projects = [Project(**p) for p in data["projects"]]
        if "publications" in data:
            resume.publications = [Publication(**p) for p in data["publications"]]
        if "awards" in data:
            resume.awards = [Award(**a) for a in data["awards"]]
        if "family_members" in data:
            resume.family_members = [FamilyMember(**f) for f in data["family_members"]]
        if "skills" in data:
            resume.skills = [Skill(**s) for s in data["skills"]]
        if "certificates" in data:
            resume.certificates = list(data["certificates"])
        if "languages" in data:
            resume.languages = list(data["languages"])
        if "hobbies" in data:
            resume.hobbies = list(data["hobbies"])

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
            "job_intent", "applied_position", "expected_position",
            "expected_salary", "expected_city", "availability",
            "self_evaluation",
            "spouse_name", "spouse_birth_date", "spouse_work_unit",
            "spouse_phone",
        ]
        for f in simple_fields:
            d[f] = getattr(self, f) or ""

        # ── Education ──
        d["education_count"] = str(len(self.education))
        d["education_summary"] = "\n".join(e.to_text() for e in self.education)
        for i, edu in enumerate(self.education, 1):
            d[f"education_{i}_school"] = edu.school
            d[f"education_{i}_major"] = edu.major
            d[f"education_{i}_degree"] = edu.degree
            d[f"education_{i}_start_date"] = edu.start_date
            d[f"education_{i}_end_date"] = edu.end_date
            d[f"education_{i}_gpa"] = edu.gpa
            d[f"education_{i}_description"] = edu.description
            d[f"education_{i}_education_level"] = edu.education_level
            d[f"education_{i}_education_form"] = edu.education_form
            d[f"education_{i}_reference_person"] = edu.reference_person
            d[f"education_{i}_reference_phone"] = edu.reference_phone
            d[f"education_{i}"] = edu.to_text()

        # ── Work experience ──
        d["work_experience_count"] = str(len(self.work_experience))
        d["work_experience_summary"] = "\n".join(w.to_text() for w in self.work_experience)
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
        d["skills_summary"] = "\n".join(
            f"{s.category}: {s.items}" for s in self.skills
        )
        for i, skill in enumerate(self.skills, 1):
            d[f"skill_{i}_category"] = skill.category
            d[f"skill_{i}_items"] = skill.items

        # ── Certificates ──
        d["certificates_count"] = str(len(self.certificates))
        d["certificates_summary"] = "、".join(self.certificates)
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
