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

## 安装

三种方式，任选一种：

### 方式一：一键安装（推荐）

**Linux / macOS:**
```bash
curl -fsSL https://raw.githubusercontent.com/elandros1/auto-resume/main/install.sh | bash
```

**Windows:**
```cmd
curl -fsSL https://raw.githubusercontent.com/elandros1/auto-resume/main/install.bat | cmd
```

### 方式二：pip 从 GitHub 安装

```bash
pip install git+https://github.com/elandros1/auto-resume.git
```

### 方式三：下载可执行文件

到 [Releases 页面](https://github.com/elandros1/auto-resume/releases) 下载对应平台的可执行文件，无需安装 Python。

## 快速开始

```bash
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

## 三种填充模式

| 模式 | 命令参数 | 原理 | 适配性 |
|------|---------|------|--------|
| **AI 模式** | `--ai` | 大模型自动理解每个标签含义 | 适配任何高校，终极方案 |
| **智能模式** | `--smart`（默认） | 正则+模糊+语义三层匹配 | 覆盖 90%+ 常见字段 |
| **基础模式** | `--no-smart` | 替换 {{占位符}} | 需模板有占位符 |

### AI 模式（终极方案）

使用大语言模型（DeepSeek/OpenAI 等）自动理解模板中每个标签的含义，映射到你的简历数据字段。**不管高校模板怎么变，AI 都能自动适配，无需改代码。**

```bash
# 设置 API Key（推荐 DeepSeek，便宜好用）
export AI_API_KEY="sk-your-key"

# AI 模式填充任意高校模板
auto-resume fill -r my_resume.json -t 福建水利电力.docx --ai

# AI 模式预览
auto-resume preview -r my_resume.json -t 某高校模板.docx --ai

# 批量 AI 填充
auto-resume fill -r my_resume.json -d 模板目录/ --ai

# 也可以用其他模型
auto-resume fill -r my_resume.json -t template.docx \
  --ai --ai-model gpt-4o \
  --ai-base-url https://api.openai.com/v1 \
  --ai-key sk-xxx
```

#### 隐私保护

AI 模式**只发送模板中的字段标签**（如"姓名""联系电话"）给大模型，**绝不发送你的个人数据**。大模型只做字段映射，填充在本地完成。

### 智能模式（默认）

工具扫描 Word 文档中的中文标签，通过三层匹配引擎自动识别：

1. **正则精确匹配** — 80+ 个预置模式，快速匹配已知字段
2. **模糊相似度匹配** — 自动识别拼写变体和近义词（"电邮地址"→email）
3. **语义关键词匹配** — 按关键词推断完全没见过的写法（"可上班日期"→availability）

支持 50+ 个字段，包括：姓名、性别、出生年月、籍贯、民族、政治面貌、学历、学位、邮箱、电话、身份证、地址、婚姻状况、应聘岗位、学习经历（多行）、工作经历（多行）、家庭成员（多行）等。

### 复杂表格处理

- **水平合并单元格**：自动跳过合并区域，找到值单元格
- **垂直合并单元格**：检测 vMerge，防止跨行串值
- **多行表格区域**：自动识别学习/工作/家庭经历表头，按行批量填充
- **上下文感知**：在"配偶"区域内，"姓名"自动映射到配偶姓名
- **字段回退**：当主要字段为空时，自动尝试备选字段

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
├── models.py            # 简历数据模型（50+ 字段）
├── engine.py            # 模板填充引擎（占位符替换）
├── detector.py          # 智能字段识别（四层匹配+合并单元格处理）
├── ai_mapper.py         # AI 大模型字段映射（终极方案）
├── fuzzy_matcher.py     # 模糊相似度+语义关键词匹配
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
