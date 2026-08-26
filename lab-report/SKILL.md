---
name: lab-report
description: 自动撰写/填写实验报告 Word 文档（.docx + .pdf）。把实验数据（手写照片/Excel/文本）按实验原理计算、绘图，结合 Word 模板和教材，生成排版规范的实验报告。遇到"写实验报告"、"填写实验报告"、"处理实验数据并出报告" 或类似任务时使用。
---

# 实验报告自动填写（可移植）

把「手写/电子实验数据 + Word 模板 + 教材/思考题」整理成一份**填好、绘图、排版规范**的实验报告（.docx + .pdf）。本技能**可移植**：不依赖任何固定目录，core 工具库与流水线脚本都随技能捆绑在本目录 `scripts/` 下，可在任意工作目录运行。

## 资源布局

本技能自带一套可复用的 Python 工具库（resourceBase 即本技能目录）：

- `scripts/run.py` —— 一键流水线入口（可移植，`--workdir` 指定工程目录）
- `scripts/core/` —— 通用工具库，跨实验复用：
  - `docx_helpers.py` —— 报告排版（段落/标题/公式/表格/图片）
  - `formula.py` —— mathtext 公式渲染（独立公式 + 正文内联）
  - `mathsym.py` —— 数学符号 → 真上下标富文本
  - `template_parser.py` —— 解析 Word 模板（章节/表格/要求）
  - `plotting.py` —— 通用绘图（散点+拟合/折线/柱状）

需要时，把 `scripts/` 整目录拷贝（或由 run.py 自动铺设）到当前工程目录使用。

## 通用 6 环节流程（适用于任意实验）

```
① 素材收集 → ② 数据识别 → ③ 模板解析 → ④ 数据处理与绘图 → ⑤ 内容撰写 → ⑥ 排版生成与交付
```

| 环节 | 做什么 | 用到的工具 | 产出 |
|------|--------|-----------|------|
| ① 素材收集 | 收集 Word 模板、数据（手写照片/Excel/文本）、教材页 | 用户提供，归入工程 `inputs/` | 素材文件 |
| ② 数据识别 | 读手写照片、转结构化数据 | `read_image` 读图 + OCR 核对；**看不清必须与用户确认，禁止猜测** | `config.json` |
| ③ 模板解析 | 提取模板的章节/数据表格/填写要求/思考题 | `scripts/core/template_parser.py` | 模板结构化信息 |
| ④ 数据处理与绘图 | 按原理计算（拟合/求值/误差），按规范绘图 | 实验模块 `analyze()` + `plot()` | results + 图表 |
| ⑤ 内容撰写 | 写目的/原理/装置/过程/讨论/结论/思考题，结合教材 | 实验模块 `build()` 正文 | 正文文本 |
| ⑥ 排版生成 | 符号统一排版 → docx → PDF | `scripts/run.py` | `outputs/*.docx` + `.pdf` |

## 工程目录约定

建议每个实验一个工程目录，形如：

```
<工作目录>/
  config.json            ← 学生信息/数据/常数（第 ② 步产出）
  inputs/                ← 素材（模板、数据照片、教材）
  outputs/               ← 产出的 .docx / .pdf
  charts/                ← 图表 png、公式 png
  core/                  ← 由 run.py 自动铺入（或拷贝）
  experiments/           ← 每个实验一个适配模块 *.py
```

## 第 ① 步 素材收集

把 Word 模板、数据照片、教材页（思考题/参考）放入工程 `inputs/`。模板若为加颜色的"填写要求"说明文字，要单独留意——正式报告生成前需删除这些说明。

## 第 ② 步 数据识别（关键，易错）

用 `read_image` 读手写照片，把数据整理进 `config.json`。**铁律：看得不清晰的关键数值（尤其温度表、周期等），必须在动手前与用户核对确认，禁止猜测。** 手写数据看错会全盘皆错。

`config.json` 通常含：`学生信息`（姓名/学号/分组/日期）、`初始条件`（仪器参数/常数/参考值）、各数据表（列表）、`输出文件名`、`实验标题`。

## 第 ③ 步 模板解析

用 `template_parser` 提取章节/表格/填写要求：

```python
from core.template_parser import parse_template
info = parse_template(r'<工程>/inputs/<模板>.doc')
# info['标题'] / info['章节'] / info['表格'] / info['说明']
```

**注意**：模板解析依赖本机 Microsoft Word（Word COM）。模板为 `.doc` 旧格式时，直接用 python-docx **重建**报告更可靠，避免 Word COM 填充损坏源文件。

## 第 ④ 步 数据配置 + 适配模块

### 建配置 config.json
按实验数据整理（见第 ② 步规范）。

### 建实验适配模块 experiments/<实验>.py
每个实验一个模块，实现三个函数：

```python
import os, sys, math
sys.path.insert(0, <工程>)               # 使 core 可导入
from core.docx_helpers import (new_document, add_heading, add_para, add_plain_para,
                               add_formula, add_picture_center, set_cell_text)
from core.formula import render_formula
from core import plotting

CHARTS = os.path.join(<工程>, 'charts')
FORMULA_DIR = os.path.join(CHARTS, 'formulas')

def analyze(config):           # 数据处理：从 config 计算，返回 results 字典
    ...                        # 例：线性拟合、某物理量、误差、不确定度
    return results

def plot(config, results):     # 绘图：生成图表 png，返回路径
    img = os.path.join(CHARTS, 'xxx.png')
    plotting.plot_fit_scatter(x, [y], ['实验数据'], img, xlabel='$t$/min', ylabel='$y$')
    return img

def build(config, results, out_path):   # 生成 docx：用 core 拼报告
    doc = new_document()
    add_plain_para(doc, config['实验标题'], size=16, bold=True, align=1, indent_first=False)
    add_heading(doc, '一、实验目的'); add_para(doc, '...')
    add_formula(doc, '公式key', FORMULA_DIR)     # 先 render_formula 出 png
    add_picture_center(doc, 图路径)
    doc.save(out_path)
```

## 第 ⑤ 步 内容撰写规范

正文结合教材，覆盖：**实验目的 → 原理 → 装置 → 过程 → 数据记录 → 数据分析 → 讨论 → 结论 → 思考题**。思考题要以教材为准、作答完整。正文写作采用"像书一样"的排版语言：

- 数学变量用 `$...$` 包裹，自动转**斜体 + 真上下标**（基线对齐），如 `$\theta_0$`→θ₀、`$m_1'$`→m₁′
- 单位/函数用正体：`$\mathrm{J/(g\cdot K)}$`
- 复杂公式（分式/大括号）用 `render_formula` 渲染成图，`add_formula` 居中插入
- 含实验数值的公式图用 `render_formula` 动态渲染（数值从 `results` 读取），换数据后必须重render
- 常量/说明文字用 `add_plain_para`（不解析 `$...$`）

## 第 ⑥ 步 排版生成与交付

运行流水线（core 与 run.py 用本技能捆绑的 `scripts/`，或已拷入工程）：

```powershell
python scripts\run.py --workdir <工程目录>
# 或指定配置与实验模块：
python scripts\run.py --workdir <工程目录> --config xxx.json --experiment experiments/yyy.py
```

输出到 `<工程>/outputs/<输出文件名>.docx/.pdf`。

## 排版规范（"像书一样"的标准）

- 汉字**宋体**、英文/数字 **Times New Roman**、正文小四 12pt、1.5 倍行距、首行缩进 2 字符
- 数学变量**斜体**、单位/函数**正体**；上下标用 Word **真上下标**
- 独立公式居中、Computer Modern 数学字体图片
- 图表：坐标轴斜体、单位规范、图题居中、网格虚线

## 交付前检查

1. **关键数值已与用户核对**（温度/周期等，防抄写错误）
2. 模板的带颜色"说明/要求"文字已在正式报告里移除（依据是模板，不是最终内容）
3. 文件名命名为 `学号_姓名_实验标题`
4. 换数据后已**重跑** `run.py`（公式图、图表都需重新生成）
5. PDF 已成功导出（依赖本机安装 Microsoft Word / Word COM）

## 已知环境依赖

- Python 包：`python-docx`、`matplotlib`、`Pillow`、`pywin32`（win32com，PDF 导出/模板解析用）
- 本机需安装 Microsoft Word（用于 PDF 导出与 .doc 模板解析）