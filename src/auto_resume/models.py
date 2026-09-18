"""Data models for resume information."""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
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

    def to_text(self) -> str:
        parts = [f"{self.school} | {self.major} | {self.degree}"]
        if self.start_date or self.end_date:
            parts.append(f"{self.start_date} - {self.end_date}")
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

    def to_text(self) -> str:
        parts = [f"{self.company} | {self.position}"]
        if self.department:
            parts[0] += f" | {self.department}"
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

    def to_text(self) -> str:
        parts = [f"{self.name}"]
        if self.role:
            parts.append(f"| {self.role}")
        if self.start_date or self.end_date:
            parts.append(f"| {self.start_date} - {self.end_date}")
        if self.technologies:
            parts.append(f"| Tech: {self.technologies}")
        if self.description:
            parts.append(f"| {self.description}")
        return " ".join(parts)


@dataclass
class Skill:
    """A skill category."""
    category: str = ""
    items: str = ""


@dataclass
class ResumeData:
    """Complete resume data structure."""
    # Basic info
    name: str = ""
    gender: str = ""
    birth_date: str = ""
    phone: str = ""
    email: str = ""
    address: str = ""
    hometown: str = ""
    political_status: str = ""
    id_number: str = ""
    nationality: str = ""
    marital_status: str = ""
    photo_path: str = ""

    # Job seeking
    job_intent: str = ""
    expected_position: str = ""
    expected_salary: str = ""
    expected_city: str = ""
    availability: str = ""

    # Self evaluation
    self_evaluation: str = ""

    # Lists
    education: list[Education] = field(default_factory=list)
    work_experience: list[WorkExperience] = field(default_factory=list)
    projects: list[Project] = field(default_factory=list)
    skills: list[Skill] = field(default_factory=list)
    certificates: list[str] = field(default_factory=list)
    languages: list[str] = field(default_factory=list)
    hobbies: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ResumeData:
        """Create ResumeData from a dictionary."""
        resume = cls()

        # Simple fields
        simple_fields = [
            "name", "gender", "birth_date", "phone", "email", "address",
            "hometown", "political_status", "id_number", "nationality",
            "marital_status", "photo_path", "job_intent", "expected_position",
            "expected_salary", "expected_city", "availability", "self_evaluation",
        ]
        for f in simple_fields:
            if f in data:
                setattr(resume, f, str(data[f]))

        # List fields
        if "education" in data:
            resume.education = [Education(**e) for e in data["education"]]
        if "work_experience" in data:
            resume.work_experience = [WorkExperience(**w) for w in data["work_experience"]]
        if "projects" in data:
            resume.projects = [Project(**p) for p in data["projects"]]
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

        This creates keys like {{name}}, {{phone}}, {{education_1}}, etc.
        Lists are expanded into indexed entries and summary strings.
        """
        d: dict[str, str] = {}

        # Simple fields
        for f in [
            "name", "gender", "birth_date", "phone", "email", "address",
            "hometown", "political_status", "id_number", "nationality",
            "marital_status", "photo_path", "job_intent", "expected_position",
            "expected_salary", "expected_city", "availability", "self_evaluation",
        ]:
            d[f] = getattr(self, f) or ""

        # Education
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
            d[f"education_{i}"] = edu.to_text()

        # Work experience
        d["work_experience_count"] = str(len(self.work_experience))
        d["work_experience_summary"] = "\n".join(w.to_text() for w in self.work_experience)
        for i, work in enumerate(self.work_experience, 1):
            d[f"work_{i}_company"] = work.company
            d[f"work_{i}_position"] = work.position
            d[f"work_{i}_start_date"] = work.start_date
            d[f"work_{i}_end_date"] = work.end_date
            d[f"work_{i}_department"] = work.department
            d[f"work_{i}_description"] = work.description
            d[f"work_{i}"] = work.to_text()

        # Projects
        d["projects_count"] = str(len(self.projects))
        d["projects_summary"] = "\n".join(p.to_text() for p in self.projects)
        for i, proj in enumerate(self.projects, 1):
            d[f"project_{i}_name"] = proj.name
            d[f"project_{i}_role"] = proj.role
            d[f"project_{i}_start_date"] = proj.start_date
            d[f"project_{i}_end_date"] = proj.end_date
            d[f"project_{i}_description"] = proj.description
            d[f"project_{i}_technologies"] = proj.technologies
            d[f"project_{i}"] = proj.to_text()

        # Skills
        d["skills_count"] = str(len(self.skills))
        d["skills_summary"] = "\n".join(
            f"{s.category}: {s.items}" for s in self.skills
        )
        for i, skill in enumerate(self.skills, 1):
            d[f"skill_{i}_category"] = skill.category
            d[f"skill_{i}_items"] = skill.items

        # Certificates
        d["certificates_count"] = str(len(self.certificates))
        d["certificates_summary"] = "、".join(self.certificates)
        for i, cert in enumerate(self.certificates, 1):
            d[f"certificate_{i}"] = cert

        # Languages
        d["languages_count"] = str(len(self.languages))
        d["languages_summary"] = "、".join(self.languages)
        for i, lang in enumerate(self.languages, 1):
            d[f"language_{i}"] = lang

        # Hobbies
        d["hobbies_count"] = str(len(self.hobbies))
        d["hobbies_summary"] = "、".join(self.hobbies)
        for i, h in enumerate(self.hobbies, 1):
            d[f"hobby_{i}"] = h

        return d

    def to_dict(self) -> dict[str, Any]:
        """Convert to a serializable dictionary."""
        return asdict(self)
