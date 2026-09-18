"""Template engine for filling Word documents with resume data."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.shared import Inches
from docx.text.paragraph import Paragraph

from .models import ResumeData


class TemplateEngine:
    """Fill Word document templates with resume data.

    Supports two placeholder syntaxes:
    1. Double-brace: {{name}}, {{phone}}, {{education_1_school}}
    2. Dollar-brace: ${name}, ${phone}
    """

    # Pattern matches {{key}} or ${key}
    PLACEHOLDER_RE = re.compile(r"\{\{(\w+)\}\}|\$\{(\w+)\}")

    def __init__(self, resume: ResumeData):
        self.resume = resume
        self.data = resume.to_flat_dict()

    def fill_template(self, template_path: str | Path, output_path: str | Path) -> Path:
        """Fill a Word template and save to output path.

        Args:
            template_path: Path to the .docx template file.
            output_path: Path to save the filled document.

        Returns:
            Path to the saved file.
        """
        template_path = Path(template_path)
        output_path = Path(output_path)

        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        doc = Document(str(template_path))

        # Process all paragraphs
        for paragraph in doc.paragraphs:
            self._fill_paragraph(paragraph)

        # Process all tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._fill_paragraph(paragraph)

        # Process headers and footers
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                self._fill_paragraph(paragraph)
            for paragraph in section.footer.paragraphs:
                self._fill_paragraph(paragraph)

        # Insert photo if specified
        if self.resume.photo_path:
            photo_path = Path(self.resume.photo_path)
            if photo_path.exists():
                self._insert_photo(doc, photo_path)

        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
        return output_path

    def _fill_paragraph(self, paragraph: Paragraph) -> None:
        """Replace placeholders in a paragraph, preserving formatting.

        Works at the run level to preserve font, size, color, bold, etc.
        """
        full_text = paragraph.text
        if not self.PLACEHOLDER_RE.search(full_text):
            return

        # Strategy: if all runs have same formatting, replace in first run
        # Otherwise, merge text, do replacement, put in first run, clear rest
        if len(paragraph.runs) == 0:
            return

        # Check if single placeholder occupies entire paragraph
        stripped = full_text.strip()
        match = self.PLACEHOLDER_RE.fullmatch(stripped)
        if match:
            key = match.group(1) or match.group(2)
            value = self.data.get(key, "")
            # Put value in first run, preserve its formatting
            paragraph.runs[0].text = value
            for run in paragraph.runs[1:]:
                run.text = ""
            return

        # General case: replace placeholders within text
        new_text = self._replace_placeholders(full_text)

        # Put all text in first run, clear others
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""

    def _replace_placeholders(self, text: str) -> str:
        """Replace all placeholders in a text string."""

        def replacer(m: re.Match) -> str:
            key = m.group(1) or m.group(2)
            return self.data.get(key, m.group(0))

        return self.PLACEHOLDER_RE.sub(replacer, text)

    def _insert_photo(self, doc: Document, photo_path: Path) -> None:
        """Insert a photo into the document.

        Looks for a paragraph containing {{photo}} or {{photo_path}}
        and replaces it with the image.
        """
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text in ("{{photo}}", "{{photo_path}}", "${photo}", "${photo_path}"):
                # Clear the paragraph
                for run in paragraph.runs:
                    run.text = ""
                # Add picture
                run = paragraph.add_run()
                run.add_picture(str(photo_path), width=Inches(1.5))
                return

    def fill_all(
        self,
        templates_dir: str | Path,
        output_dir: str | Path,
        name_prefix: str = "",
    ) -> list[Path]:
        """Fill all templates in a directory.

        Args:
            templates_dir: Directory containing .docx templates.
            output_dir: Directory to save filled documents.
            name_prefix: Optional prefix for output filenames.

        Returns:
            List of paths to saved files.
        """
        templates_dir = Path(templates_dir)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        results: list[Path] = []
        for template in sorted(templates_dir.glob("*.docx")):
            if template.name.startswith("~$"):
                continue
            output_name = f"{name_prefix}{template.stem}_filled.docx"
            output_path = output_dir / output_name
            self.fill_template(template, output_path)
            results.append(output_path)

        return results

    def list_available_keys(self) -> list[str]:
        """Return all available placeholder keys."""
        return sorted(self.data.keys())


def fill_resume_to_template(
    resume_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
) -> Path:
    """Convenience function: fill a single template with resume data.

    Args:
        resume_path: Path to resume JSON file.
        template_path: Path to Word template .docx file.
        output_path: Path to save the filled document.

    Returns:
        Path to the saved file.
    """
    resume = ResumeData.from_json(resume_path)
    engine = TemplateEngine(resume)
    return engine.fill_template(template_path, output_path)


def fill_resume_to_all_templates(
    resume_path: str | Path,
    templates_dir: str | Path,
    output_dir: str | Path,
) -> list[Path]:
    """Convenience function: fill all templates with resume data.

    Args:
        resume_path: Path to resume JSON file.
        templates_dir: Directory with .docx templates.
        output_dir: Directory to save filled documents.

    Returns:
        List of saved file paths.
    """
    resume = ResumeData.from_json(resume_path)
    engine = TemplateEngine(resume)
    name = resume.name or "resume"
    return engine.fill_all(templates_dir, output_dir, name_prefix=f"{name}_")
