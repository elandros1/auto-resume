"""Standalone entry point for PyInstaller packaging.

This file avoids relative imports (from . import ...) which fail when
PyInstaller bundles the app as a single executable.
"""
from __future__ import annotations

import sys
from pathlib import Path

# When PyInstaller bundles with --onefile, __file__ points to a
# temp dir; we need the package's parent on the path.
if getattr(sys, "frozen", False):
    base_dir = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).parent
else:
    base_dir = Path(__file__).parent

package_parent = str(base_dir)
if package_parent not in sys.path:
    sys.path.insert(0, package_parent)

src_parent = str(base_dir.parent)
if src_parent not in sys.path:
    sys.path.insert(0, src_parent)

import click
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.table import Table as RichTable

from auto_resume.ai_mapper import AIFieldMapper
from auto_resume.detector import FieldDetector
from auto_resume.engine import TemplateEngine
from auto_resume.models import ResumeData
from auto_resume.template_generator import generate_all_templates

console = Console()


@click.group()
def cli() -> None:
    """Auto-Resume: 一份简历数据，自动填充到任意学校的招聘表格"""


@cli.command()
@click.argument("output_path", type=click.Path())
def new(output_path: str) -> None:
    """创建简历数据模板文件"""
    sample = ResumeData.sample()
    sample.to_json(output_path)
    console.print(f"[green]✓ 已创建简历模板: {output_path}[/green]")
    console.print("[dim]请编辑此文件，填入你的真实信息[/dim]")


@cli.command()
@click.option("-r", "--resume", "resume_path", required=True, type=click.Path(exists=True), help="简历数据文件 (JSON)")
@click.option("-t", "--template", "template_path", type=click.Path(exists=True), help="Word模板文件")
@click.option("-d", "--templates-dir", "templates_dir", type=click.Path(exists=True), help="模板目录（批量）")
@click.option("-o", "--output", "output_dir", default="output", type=click.Path(), help="输出目录")
@click.option("--smart/--no-smart", default=True, help="智能模式（默认开启）")
@click.option(
    "--ai/--no-ai",
    default=False,
    help="启用AI模式：用大模型自动映射字段（需配置AI_API_KEY）",
)
@click.option(
    "--ai-key",
    "ai_api_key",
    default=None,
    help="AI API密钥（也可通过环境变量AI_API_KEY设置）",
)
@click.option(
    "--ai-model",
    "ai_model",
    default="deepseek-chat",
    help="AI模型名称（默认: deepseek-chat）",
)
@click.option(
    "--ai-base-url",
    "ai_base_url",
    default="https://api.deepseek.com/v1",
    help="AI API地址（默认: DeepSeek）",
)
def fill(
    resume_path: str,
    template_path: str | None,
    templates_dir: str | None,
    output_dir: str,
    smart: bool,
    ai: bool,
    ai_api_key: str | None,
    ai_model: str,
    ai_base_url: str,
) -> None:
    """填充模板：用简历数据自动填写 Word 文档

    三种模式：
    1. --ai     AI模式：用大模型自动映射，适配任何高校模板（终极方案）
    2. --smart  智能模式：正则+模糊匹配+语义关键词（默认）
    3. --no-smart 基础模式：仅替换占位符
    """
    if not template_path and not templates_dir:
        console.print("[red]错误: 必须指定 --template 或 --templates-dir[/red]")
        sys.exit(1)

    resume = ResumeData.from_json(resume_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if template_path:
        # Single template mode
        tpl_path = Path(template_path)
        out_name = f"{resume.name}_{tpl_path.stem}_filled.docx"
        out_path = output_dir / out_name

        if ai:
            ai_mapper = AIFieldMapper(
                api_key=ai_api_key,
                model=ai_model,
                base_url=ai_base_url,
            )
            if not ai_mapper.is_available():
                console.print("[red]错误: 未配置AI API密钥[/red]")
                console.print("[dim]请通过 --ai-key 参数或 AI_API_KEY 环境变量设置[/dim]")
                sys.exit(1)

            flat_data = resume.to_flat_dict()
            resume_keys = list(flat_data.keys())

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                console=console,
            ) as progress:
                task = progress.add_task("AI识别字段映射...", total=None)
                ai_mapping = ai_mapper.map_fields(tpl_path, resume_keys)
                progress.update(task, completed=True)

            console.print(f"[green]✓ AI识别到 {len(ai_mapping)} 个字段映射[/green]")

            detector = FieldDetector()
            detector.enable_ai(ai_mapping)

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                console=console,
            ) as progress:
                task = progress.add_task(f"AI填充 {tpl_path.name}...", total=None)
                result = detector.auto_fill(tpl_path, flat_data, out_path)
                progress.update(task, completed=True)

            console.print(f"\n[green]✓ 已生成: {result}[/green]")
            console.print(f"  [dim]模板: {tpl_path}[/dim]")
            console.print(f"  [dim]数据: {resume_path}[/dim]")
            console.print(f"  [dim]模式: AI ({ai_model})[/dim]")

        elif smart:
            detector = FieldDetector()
            data = resume.to_flat_dict()
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                console=console,
            ) as progress:
                task = progress.add_task(f"智能填充 {tpl_path.name}...", total=None)
                result = detector.auto_fill(tpl_path, data, out_path)
                progress.update(task, completed=True)

            console.print(f"\n[green]✓ 已生成: {result}[/green]")
            console.print(f"  [dim]模板: {tpl_path}[/dim]")
            console.print(f"  [dim]数据: {resume_path}[/dim]")

        else:
            engine = TemplateEngine(resume)
            result = engine.fill_template(tpl_path, out_path)
            console.print(f"[green]✓ 已生成: {result}[/green]")

    elif templates_dir:
        # Batch fill all templates
        tpl_dir = Path(templates_dir)
        templates = [t for t in tpl_dir.glob("*.docx") if not t.name.startswith("~$")]

        if not templates:
            console.print("[red]错误: 目录中没有找到 .docx 模板文件[/red]")
            sys.exit(1)

        console.print(f"[cyan]找到 {len(templates)} 个模板，开始批量填充...[/cyan]\n")

        results = []
        detector = None
        ai_mapper = None
        data = None

        if ai:
            ai_mapper = AIFieldMapper(
                api_key=ai_api_key,
                model=ai_model,
                base_url=ai_base_url,
            )
            if not ai_mapper.is_available():
                console.print("[red]错误: 未配置AI API密钥[/red]")
                sys.exit(1)
            detector = FieldDetector()
            data = resume.to_flat_dict()
            console.print(f"[cyan]AI模式: 使用 {ai_model}[/cyan]")
        elif smart:
            detector = FieldDetector()
            data = resume.to_flat_dict()

        for template in templates:
            out_name = f"{resume.name}_{template.stem}_filled.docx"
            out_path = output_dir / out_name

            try:
                if ai:
                    flat_data = resume.to_flat_dict()
                    resume_keys = list(flat_data.keys())
                    ai_mapping = ai_mapper.map_fields(template, resume_keys)
                    detector.enable_ai(ai_mapping)
                    result = detector.auto_fill(template, flat_data, out_path)
                    detector.disable_ai()
                    results.append((template.name, out_path, "成功"))
                    console.print(f"  [green]✓[/green] {template.name} → {out_path.name}")
                elif smart:
                    result = detector.auto_fill(template, data, out_path)
                    results.append((template.name, out_path, "成功"))
                    console.print(f"  [green]✓[/green] {template.name} → {out_path.name}")
                else:
                    engine = TemplateEngine(resume)
                    engine.fill_template(template, out_path)
                    results.append((template.name, out_path, "成功"))
                    console.print(f"  [green]✓[/green] {template.name} → {out_path.name}")
            except Exception as e:
                results.append((template.name, None, str(e)))
                console.print(f"  [red]✗[/red] {template.name} → {e}")

        console.print(f"\n[cyan]批量完成: {sum(1 for r in results if r[2]=='成功')}/{len(results)} 成功[/cyan]")


@cli.command()
@click.option("-r", "--resume", "resume_path", required=True, type=click.Path(exists=True))
@click.option("-t", "--template", "template_path", required=True, type=click.Path(exists=True))
@click.option("--ai/--no-ai", default=False, help="启用AI模式预览")
@click.option("--ai-key", "ai_api_key", default=None)
@click.option("--ai-model", "ai_model", default="deepseek-chat")
@click.option("--ai-base-url", "ai_base_url", default="https://api.deepseek.com/v1")
def preview(
    resume_path: str,
    template_path: str,
    ai: bool,
    ai_api_key: str | None,
    ai_model: str,
    ai_base_url: str,
) -> None:
    """预览：查看模板中哪些字段会被自动识别"""
    resume = ResumeData.from_json(resume_path)
    data = resume.to_flat_dict()
    detector = FieldDetector()

    if ai:
        ai_mapper = AIFieldMapper(
            api_key=ai_api_key,
            model=ai_model,
            base_url=ai_base_url,
        )
        if ai_mapper.is_available():
            mapping = ai_mapper.map_fields(template_path, list(data.keys()))
            detector.enable_ai(mapping)
            console.print("[cyan]AI映射结果:[/cyan]")
            for label, key in mapping.items():
                if key and not key.startswith("section:"):
                    console.print(f"  [green]✓[/green] {label} → {key}")
                elif key:
                    console.print(f"  [yellow]§[/yellow] {label} → {key}")
                else:
                    console.print(f"  [red]✗[/red] {label}")

    tpl_path = Path(template_path)
    matches = detector.detect_fields(tpl_path, data)

    console.print(f"\n[cyan]模板: {tpl_path.name}[/cyan]")
    console.print(f"共 {len(matches)} 个字段:\n")

    table = RichTable(show_header=True, header_style="bold blue")
    table.add_column("标签", style="cyan")
    table.add_column("匹配字段", style="green")
    table.add_column("分数")
    table.add_column("方法")

    for m in matches:
        score_str = f"{m['score']:.0%}" if m["score"] > 0 else "—"
        table.add_row(m["label"], m["key"] or "[red]未匹配[/red]", score_str, m["method"])

    console.print(table)

    matched_count = sum(1 for m in matches if m["key"])
    console.print(f"\n[green]已匹配: {matched_count}/{len(matches)}[/green]")


@cli.command()
@click.option("-o", "--output", "output_dir", default="templates", type=click.Path(), help="输出目录")
def generate(output_dir: str) -> None:
    """生成内置示例模板"""
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    generate_all_templates(out)
    console.print(f"[green]✓ 已生成模板到: {out}[/green]")


@cli.command()
@click.option("-r", "--resume", "resume_path", required=True, type=click.Path(exists=True))
def keys(resume_path: str) -> None:
    """查看简历数据中的所有可用字段"""
    resume = ResumeData.from_json(resume_path)
    data = resume.to_flat_dict()
    console.print(f"[cyan]共 {len(data)} 个字段:[/cyan]\n")
    for k, v in sorted(data.items()):
        val_display = v[:40] + "..." if len(v) > 40 else v
        console.print(f"  [green]{k}[/green]: {val_display}")


if __name__ == "__main__":
    cli()
