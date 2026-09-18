"""Generate sample Word templates for testing and demonstration."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_cell_font(cell, name="宋体", size=10.5, bold=False):
    """Set font for all text in a cell."""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)
            run.font.name = name
            run.font.bold = bold
            # Set East Asian font
            run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def add_chinese_font(paragraph, name="宋体", size=10.5, bold=False):
    """Add a run with Chinese font support."""
    run = paragraph.add_run()
    run.font.size = Pt(size)
    run.font.name = name
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        from lxml import etree
        rfonts = etree.SubElement(rpr, qn("w:rFonts"))
    rfonts.set(qn("w:eastAsia"), name)
    return run


def create_template_1(out_path: str | Path) -> Path:
    """Create Template 1: 求职登记表 (Job Application Form)

    A standard table-based form with labels on left, blanks on right.
    """
    doc = Document()

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("求职登记表")
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.name = "宋体"
    title._element.rPr = run._element.rPr if run._element.rPr is not None else None

    # Main table: 8 rows, 4 columns (label-value-label-value pattern)
    table = doc.add_table(rows=9, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: Name | [blank] | Gender | [blank]
    cells = table.rows[0].cells
    cells[0].text = "姓名"
    cells[2].text = "性别"

    # Row 1: Birth | [blank] | Phone | [blank]
    cells = table.rows[1].cells
    cells[0].text = "出生年月"
    cells[2].text = "联系电话"

    # Row 2: Email | [blank] | Hometown | [blank]
    cells = table.rows[2].cells
    cells[0].text = "电子邮箱"
    cells[2].text = "籍贯"

    # Row 3: Political | [blank] | Nationality | [blank]
    cells = table.rows[3].cells
    cells[0].text = "政治面貌"
    cells[2].text = "民族"

    # Row 4: ID | [blank] | Marital | [blank]
    cells = table.rows[4].cells
    cells[0].text = "身份证号"
    cells[2].text = "婚姻状况"

    # Row 5: Address (merged, full width)
    cells = table.rows[5].cells
    cells[0].text = "通讯地址"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 6: Expected position (merged)
    cells = table.rows[6].cells
    cells[0].text = "期望岗位"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 7: Expected salary (merged)
    cells = table.rows[7].cells
    cells[0].text = "期望薪资"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 8: Education (merged)
    cells = table.rows[8].cells
    cells[0].text = "毕业院校"
    cells[1].merge(cells[2]).merge(cells[3])

    # Set fonts for all cells
    for row in table.rows:
        for cell in row.cells:
            set_cell_font(cell, "宋体", 10.5)

    # Self evaluation section
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("自我评价：")
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.name = "宋体"

    p2 = doc.add_paragraph()
    run2 = p2.add_run("{{self_evaluation}}")
    run2.font.size = Pt(10.5)
    run2.font.name = "宋体"

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def create_template_2(out_path: str | Path) -> Path:
    """Create Template 2: 个人简历 (Personal Resume)

    A paragraph-style resume with {{placeholder}} markers.
    """
    doc = Document()

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("个人简历")
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.name = "宋体"

    # Section: Basic Info
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("一、基本信息")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    info_lines = [
        "姓    名：{{name}}          性    别：{{gender}}",
        "出生年月：{{birth_date}}      联系电话：{{phone}}",
        "电子邮箱：{{email}}          籍    贯：{{hometown}}",
        "政治面貌：{{political_status}}      民    族：{{nationality}}",
        "通讯地址：{{address}}",
    ]
    for line in info_lines:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(10.5)
        run.font.name = "宋体"

    # Section: Job Intent
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("二、求职意向")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    intent_lines = [
        "期望岗位：{{expected_position}}",
        "期望薪资：{{expected_salary}}",
        "期望城市：{{expected_city}}",
        "到岗时间：{{availability}}",
    ]
    for line in intent_lines:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(10.5)
        run.font.name = "宋体"

    # Section: Education
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("三、教育背景")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("{{education_summary}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    # Section: Work Experience
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("四、工作经历")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("{{work_experience_summary}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    # Section: Projects
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("五、项目经历")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("{{projects_summary}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    # Section: Skills
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("六、专业技能")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("{{skills_summary}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    # Section: Certificates
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("七、证书与语言")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("证书：{{certificates_summary}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("语言：{{languages_summary}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    # Section: Self Evaluation
    doc.add_paragraph()
    h = doc.add_paragraph()
    run = h.add_run("八、自我评价")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "宋体"

    p = doc.add_paragraph()
    run = p.add_run("{{self_evaluation}}")
    run.font.size = Pt(10.5)
    run.font.name = "宋体"

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def create_template_3(out_path: str | Path) -> Path:
    """Create Template 3: 高校招聘应聘表 (University Application Form)

    A complex table form simulating a real university recruitment form.
    """
    doc = Document()

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("高校教师应聘登记表")
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.name = "宋体"

    # Main table - simpler structure with 4 columns
    table = doc.add_table(rows=11, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: Name | blank | Gender | blank
    cells = table.rows[0].cells
    cells[0].text = "姓名"
    cells[2].text = "性别"

    # Row 1: Birth | blank | Hometown | blank
    cells = table.rows[1].cells
    cells[0].text = "出生年月"
    cells[2].text = "籍贯"

    # Row 2: Political | blank | Nationality | blank
    cells = table.rows[2].cells
    cells[0].text = "政治面貌"
    cells[2].text = "民族"

    # Row 3: Phone | blank | Email | blank
    cells = table.rows[3].cells
    cells[0].text = "联系电话"
    cells[2].text = "电子邮箱"

    # Row 4: ID | blank | Marital | blank
    cells = table.rows[4].cells
    cells[0].text = "身份证号"
    cells[2].text = "婚姻状况"

    # Row 5: Address (label + merged blank)
    cells = table.rows[5].cells
    cells[0].text = "通讯地址"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 6: Expected position (label + merged blank)
    cells = table.rows[6].cells
    cells[0].text = "应聘岗位"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 7: Expected salary + city
    cells = table.rows[7].cells
    cells[0].text = "期望薪资"
    cells[2].text = "期望城市"

    # Row 8: Education (label + merged blank)
    cells = table.rows[8].cells
    cells[0].text = "教育背景"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 9: Work experience
    cells = table.rows[9].cells
    cells[0].text = "工作经历"
    cells[1].merge(cells[2]).merge(cells[3])

    # Row 10: Skills/Certificates
    cells = table.rows[10].cells
    cells[0].text = "专业技能"
    cells[1].merge(cells[2]).merge(cells[3])

    # Set fonts
    for row in table.rows:
        for cell in row.cells:
            set_cell_font(cell, "宋体", 10.5)

    # Self evaluation section (paragraphs)
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("自我评价：")
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.name = "宋体"

    p2 = doc.add_paragraph()
    run2 = p2.add_run("{{self_evaluation}}")
    run2.font.size = Pt(10.5)
    run2.font.name = "宋体"

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def generate_all_templates(output_dir: str | Path) -> list[Path]:
    """Generate all sample templates."""
    output_dir = Path(output_dir)
    results = [
        create_template_1(output_dir / "求职登记表.docx"),
        create_template_2(output_dir / "个人简历.docx"),
        create_template_3(output_dir / "高校教师应聘表.docx"),
    ]
    return results


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "templates"
    paths = generate_all_templates(out)
    for p in paths:
        print(f"Generated: {p}")
