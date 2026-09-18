"""CLI entry point for auto-resume."""

from __future__ import annotations

import sys
from pathlib import Path

import click
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.table import Table as RichTable

from . import get_resource_path
from .detector import FieldDetector
from .engine import TemplateEngine
from .models import ResumeData
from .template_generator import generate_all_templates

console = Console()


@click.group()
@click.version_option("1.0.0", prog_name="auto-resume")
def cli():
    """Auto-Resume: 自动填写简历/表格到 Word 文档

    一份数据，自动填充到任意学校的招聘表格。
    """
    pass


@cli.command()
@click.option(
    "-r", "--resume",
    "resume_path",
    required=True,
    type=click.Path(exists=True),
    help="简历数据 JSON 文件路径",
)
@click.option(
    "-t", "--template",
    "template_path",
    type=click.Path(exists=True),
    help="Word 模板文件路径 (.docx)",
)
@click.option(
    "-d", "--templates-dir",
    "templates_dir",
    type=click.Path(exists=True),
    help="模板目录（批量填充所有 .docx 模板）",
)
@click.option(
    "-o", "--output",
    "output_dir",
    default="./output",
    type=click.Path(),
    help="输出目录（默认: ./output）",
)
@click.option(
    "--smart/--no-smart",
    default=True,
    help="启用智能识别模式（自动检测中文标签填充）",
)
def fill(resume_path, template_path, templates_dir, output_dir, smart):
    """填充模板：用简历数据自动填写 Word 文档"""
    resume = ResumeData.from_json(resume_path)

    if not templates_dir and not template_path:
        # Use built-in templates
        templates_dir = str(get_resource_path("templates"))
        if not Path(templates_dir).exists():
            console.print("[yellow]未指定模板，正在生成内置模板...[/yellow]")
            generate_all_templates(templates_dir)

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if template_path:
        # Single template
        template_path = Path(template_path)
        out_name = f"{resume.name}_{template_path.stem}_filled.docx"
        out_path = output_dir / out_name

        if smart:
            detector = FieldDetector()
            data = resume.to_flat_dict()
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                console=console,
            ) as progress:
                task = progress.add_task(f"智能填充 {template_path.name}...", total=None)
                result = detector.auto_fill(template_path, data, out_path)
                progress.update(task, completed=True)

            console.print(f"\n[green]✓ 已生成: {result}[/green]")
            console.print(f"  [dim]模板: {template_path}[/dim]")
            console.print(f"  [dim]数据: {resume_path}[/dim]")
        else:
            engine = TemplateEngine(resume)
            result = engine.fill_template(template_path, out_path)
            console.print(f"[green]✓ 已生成: {result}[/green]")

    elif templates_dir:
        # Batch fill all templates
        templates_dir = Path(templates_dir)
        templates = list(templates_dir.glob("*.docx"))
        templates = [t for t in templates if not t.name.startswith("~$")]

        if not templates:
            console.print("[red]错误: 目录中没有找到 .docx 模板文件[/red]")
            sys.exit(1)

        console.print(f"[cyan]找到 {len(templates)} 个模板，开始批量填充...[/cyan]\n")

        results = []
        detector = FieldDetector() if smart else None
        data = resume.to_flat_dict() if smart else None

        for template in templates:
            out_name = f"{resume.name}_{template.stem}_filled.docx"
            out_path = output_dir / out_name

            try:
                if smart:
                    detector.auto_fill(template, data, out_path)
                else:
                    engine = TemplateEngine(resume)
                    engine.fill_template(template, out_path)
                results.append((template.name, out_path, "成功"))
                console.print(f"  [green]✓[/green] {template.name} → {out_path.name}")
            except Exception as e:
                results.append((template.name, None, str(e)))
                console.print(f"  [red]✗[/red] {template.name} → {e}")

        # Summary
        console.print()
        summary = RichTable(title="填充结果汇总")
        summary.add_column("模板", style="cyan")
        summary.add_column("输出文件", style="green")
        summary.add_column("状态", style="yellow")

        for name, path, status in results:
            summary.add_row(name, path.name if path else "-", status)

        console.print(summary)
        count = len(results)
        console.print(
            f"\n[bold green]完成！共 {count} 个文件已生成到 {output_dir}/[/bold green]"
        )


@cli.command()
@click.option(
    "-r", "--resume",
    "resume_path",
    required=True,
    type=click.Path(exists=True),
    help="简历数据 JSON 文件路径",
)
@click.option(
    "-t", "--template",
    "template_path",
    required=True,
    type=click.Path(exists=True),
    help="要预览的 Word 模板路径",
)
def preview(resume_path, template_path):
    """预览：查看模板中哪些字段会被自动识别和填充"""

    resume = ResumeData.from_json(resume_path)
    data = resume.to_flat_dict()
    detector = FieldDetector()
    previews = detector.preview_mapping(template_path)

    if not previews:
        console.print("[yellow]未检测到可自动识别的字段[/yellow]")
        console.print('[dim]提示: 确保模板中使用中文标签（如姓名、电话等）[/dim]')
        return

    table = RichTable(title=f"智能识别结果 — {Path(template_path).name}")
    table.add_column("序号", style="dim", width=4)
    table.add_column("模板中的标签", style="cyan")
    table.add_column("映射到字段", style="yellow")
    table.add_column("当前值", style="green")
    table.add_column("位置", style="dim")

    for i, p in enumerate(previews, 1):
        value = data.get(p["resume_key"], "")
        if len(value) > 30:
            value = value[:30] + "..."
        table.add_row(str(i), p["label"], p["resume_key"], value or "[空]", p["location"])

    console.print(table)
    console.print(f"\n[bold]共识别 {len(previews)} 个字段[/bold]")
    matched = sum(1 for p in previews if data.get(p["resume_key"]))
    console.print(f"[green]已匹配: {matched}[/green]  [red]未匹配: {len(previews) - matched}[/red]")


@cli.command()
@click.option(
    "-o", "--output",
    "output_dir",
    default="./templates",
    type=click.Path(),
    help="输出目录（默认: ./templates）",
)
def generate(output_dir):
    """生成内置示例模板"""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    console.print("[cyan]正在生成内置模板...[/cyan]")
    paths = generate_all_templates(output_dir)

    for p in paths:
        console.print(f"  [green]✓[/green] {p.name}")

    console.print(f"\n[bold green]完成！{len(paths)} 个模板已生成到 {output_dir}/[/bold green]")
    console.print("[dim]你可以编辑这些模板，或直接上传学校的表格模板来使用[/dim]")


@cli.command()
@click.option(
    "-r", "--resume",
    "resume_path",
    required=True,
    type=click.Path(exists=True),
    help="简历数据 JSON 文件路径",
)
def keys(resume_path):
    """查看简历数据中所有可用的占位符字段"""
    resume = ResumeData.from_json(resume_path)
    data = resume.to_flat_dict()

    table = RichTable(title="可用占位符字段", show_lines=True)
    table.add_column("字段名", style="cyan", width=30)
    table.add_column("值", style="green")

    for key in sorted(data.keys()):
        value = data[key]
        if len(value) > 60:
            value = value[:60] + "..."
        table.add_row(key, value or "[空]")

    console.print(table)
    console.print(f"\n[bold]共 {len(data)} 个字段[/bold]")
    console.print("[dim]在 Word 模板中使用 {{字段名}} 或 ${字段名} 作为占位符[/dim]")


@cli.command(name="new")
@click.argument("output_path", type=click.Path())
def new_resume(output_path):
    """创建一个新的简历数据文件（模板）"""
    import json

    template = {
        # Basic info
        "name": "",
        "gender": "",
        "birth_date": "",
        "age": "",
        "phone": "",
        "email": "",
        "address": "",
        "postal_code": "",
        "hometown": "",
        "hukou_location": "",
        "political_status": "",
        "id_number": "",
        "nationality": "",
        "marital_status": "",
        "height": "",
        "vision": "",
        "photo_path": "",
        # Professional info
        "specialty": "",
        "computer_proficiency": "",
        "foreign_language": "",
        "previous_employer": "",
        "current_position": "",
        "professional_title": "",
        "social_experience": "",
        # Job seeking
        "job_intent": "",
        "applied_position": "",
        "expected_position": "",
        "expected_salary": "",
        "expected_city": "",
        "availability": "",
        # Self evaluation
        "self_evaluation": "",
        # Spouse
        "spouse_name": "",
        "spouse_birth_date": "",
        "spouse_work_unit": "",
        "spouse_phone": "",
        # Education
        "education": [
            {
                "school": "",
                "major": "",
                "degree": "",
                "start_date": "",
                "end_date": "",
                "gpa": "",
                "description": "",
                "education_level": "",
                "education_form": "",
                "reference_person": "",
                "reference_phone": ""
            }
        ],
        # Work experience
        "work_experience": [
            {
                "company": "",
                "position": "",
                "start_date": "",
                "end_date": "",
                "department": "",
                "description": "",
                "reference_person": "",
                "reference_phone": "",
                "professional_title": ""
            }
        ],
        # Projects
        "projects": [
            {
                "name": "",
                "role": "",
                "start_date": "",
                "end_date": "",
                "description": "",
                "technologies": "",
                "funding_source": "",
                "funding_amount": ""
            }
        ],
        # Publications
        "publications": [
            {
                "title": "",
                "journal": "",
                "date": "",
                "authors": "",
                "index": ""
            }
        ],
        # Awards
        "awards": [
            {
                "date": "",
                "title": "",
                "level": "",
                "issuer": ""
            }
        ],
        # Family members
        "family_members": [
            {
                "name": "",
                "relationship": "",
                "gender": "",
                "birth_date": "",
                "phone": "",
                "work_unit": "",
                "position": "",
                "address": "",
                "political_status": ""
            }
        ],
        # Skills
        "skills": [
            {"category": "", "items": ""}
        ],
        "certificates": [],
        "languages": [],
        "hobbies": []
    }

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(template, f, ensure_ascii=False, indent=2)

    console.print(f"[green]✓ 已创建简历模板: {output_path}[/green]")
    console.print("[dim]编辑这个 JSON 文件填入你的信息，然后运行 fill 命令[/dim]")


if __name__ == "__main__":
    cli()
