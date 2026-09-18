# Auto-Resume

> 一份简历数据，自动填充到任意学校的招聘表格

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/Tests-26%20passed-brightgreen.svg)](https://docs.pytest.org/)

## 解决什么问题

每个高校的招聘简历/登记表格式都不一样。手动逐个填写费时费力。

Auto-Resume 让你：
1. **维护一份**简历数据（JSON 格式）
2. 下载任意学校的 Word 表格
3. 工具**自动识别**表中的"姓名""电话"等中文标签并填充
4. 一键生成填好的文件

## 快速开始

```bash
# 安装
pip install -e .

# 创建你的简历数据文件
auto-resume new my_resume.json

# 编辑 my_resume.json 填入你的信息

# 生成内置模板
auto-resume generate -o templates/

# 预览：查看模板中哪些字段会被自动识别
auto-resume preview -r my_resume.json -t templates/高校教师应聘表.docx

# 智能填充单个模板
auto-resume fill -r my_resume.json -t templates/高校教师应聘表.docx -o output/

# 批量填充所有模板
auto-resume fill -r my_resume.json -d templates/ -o output/

# 查看所有可用字段
auto-resume keys -r my_resume.json
```

## 工作原理

### 智能识别模式

工具扫描 Word 文档中的中文标签（如"姓名""联系电话""电子邮箱"），自动找到旁边的空白位置，填入对应数据。

支持识别的字段：

| 中文标签 | 对应字段 |
|---------|---------|
| 姓名 | name |
| 性别 | gender |
| 出生年月 | birth_date |
| 联系电话/手机 | phone |
| 电子邮箱/Email | email |
| 籍贯 | hometown |
| 政治面貌 | political_status |
| 民族 | nationality |
| 身份证号 | id_number |
| 婚姻状况 | marital_status |
| 通讯地址 | address |
| 应聘岗位/期望岗位 | expected_position |
| 期望薪资 | expected_salary |
| 期望城市 | expected_city |
| 毕业院校 | education_summary |
| 工作经历 | work_experience_summary |
| 专业技能 | skills_summary |
| 证书 | certificates_summary |
| 语言能力 | languages |
| 自我评价 | self_evaluation |

### 占位符模式

也可以在 Word 模板中直接使用占位符标记：
- `{{name}}` — 双花括号语法
- `${name}` — 美元花括号语法

## 使用任意学校的表格

```bash
# 1. 下载学校的 Word 表格模板
# 2. 预览哪些字段会被自动识别
auto-resume preview -r my_resume.json -t 学校A应聘表.docx

# 3. 智能填充
auto-resume fill -r my_resume.json -t 学校A应聘表.docx -o output/

# 4. 对学校B也做同样的事
auto-resume fill -r my_resume.json -t 学校B登记表.docx -o output/
```

## 内置模板

| 模板 | 说明 |
|------|------|
| `求职登记表.docx` | 标准表格格式，左标签右填空 |
| `个人简历.docx` | 段落式简历，使用占位符 |
| `高校教师应聘表.docx` | 复杂表格，模拟高校招聘场景 |

## 项目结构

```
src/auto_resume/
├── __init__.py          # 包入口
├── __main__.py          # CLI 命令（fill/preview/generate/keys/new）
├── models.py            # 简历数据模型
├── engine.py            # 模板填充引擎（占位符替换）
├── detector.py          # 智能字段识别（自动检测中文标签）
├── template_generator.py # 内置模板生成器
├── samples/
│   └── resume.json      # 示例简历数据
└── templates/
    ├── 求职登记表.docx
    ├── 个人简历.docx
    └── 高校教师应聘表.docx
```

## License

MIT
