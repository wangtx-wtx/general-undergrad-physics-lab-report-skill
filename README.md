# general-undergrad-physics-lab-report-skill

**EN:** Portable agent skill for auto-writing typeset undergraduate physics lab reports (.docx/.pdf) — compute, plot, and draft the full report from handed data and a Word template. Self-contained, one-command install, reusable across experiments.

**中文：** 自动撰写排版规范的本科生物理实验报告（.docx/.pdf）的可移植 agent 技能——从手写数据与 Word 模板出发，完成计算、绘图与整篇撰写。自包含、一键安装、跨实验可复用。

---

## 特性 / Features

- **通用 6 环节流水线**：素材收集 → 数据识别 → 模板解析 → 数据处理与绘图 → 内容撰写 → 排版生成
- **可移植**：`SKILL.md` + `scripts/` 全套工具库捆绑一体；`run.py` 用 `--workdir` 指定任意工程目录，跨实验复用
- **排版规范**：宋体汉字 / Times New Roman 西文、正文小四 1.5 倍行距、首行缩进 2 字符、`$...$` 数学片段转真上下标、居中公式图、网格虚线图表
- **安装即用**：Windows / macOS / Linux 一键安装脚本，自动铺技能、装依赖、语法自检
- **中英双语 / Bilingual**：README 与说明文档中英双语，便于国内与国际用户使用

## 目录结构 / Structure

```
lab-report/
  SKILL.md           技能指令正文（通用 SKILL.md 格式）
  scripts/run.py     一键流水线入口（--workdir 指定工程目录）
  scripts/core/      通用工具库（跨实验复用）
    docx_helpers.py   报告排版（段落/标题/公式/表格/图片）
    formula.py        mathtext 公式渲染（独立公式 + 内联）
    mathsym.py        数学符号 → 真上下标富文本
    template_parser.py 解析 Word 模板（章节/表格/要求）
    plotting.py       通用绘图（散点+拟合/折线/柱状）
  requirements.txt   Python 依赖
  install.ps1        Windows 一键安装
  install.sh         macOS / Linux 一键安装
  INSTALL.md         分发与安装说明
```

## 快速开始 / Quick Start

```powershell
# 1. 安装依赖 / Install dependencies
pip install -r requirements.txt

# 2. 把 lab-report 放进你 agent 的 skills 目录，或直接运行 install.ps1 / install.sh 铺设
#    Put the lab-report folder into your agent's skills directory, or run the install script.

# 3. 在任意实验工程目录运行 / Run in any experiment workspace
python scripts\run.py --workdir <实验工程目录>
```

输出 `<工程>/outputs/<学号_姓名_实验标题>.docx/.pdf`。
完整流程与排版规范见 `lab-report/INSTALL.md` 与 `lab-report/SKILL.md`。

## 安装与使用（给使用者）/ Getting Started for Users

> 拿到本仓库后，按下面 4 步即可把这个技能装进你自己的 agent 环境，并用它生成实验报告。

**第 1 步 — 获取代码 / Get the code**

```bash
git clone https://github.com/wangtx-wtx/general-undergrad-physics-lab-report-skill
```
或在仓库主页点 **Code → Download ZIP** 下载解压。

**第 2 步 — 安装技能 / Install the skill**

把 `lab-report` 目录放进你 agent 的 **skills 目录**（不同 agent 位置不同，详见 `lab-report/INSTALL.md`）。也可运行自带安装脚本自动铺设：

- Windows：在 `lab-report` 目录执行 `.\install.ps1`
- macOS / Linux：执行 `bash install.sh`

脚本会自动把技能铺到当前目录的 `skills/lab-report`，并安装 Python 依赖、做语法自检。

**第 3 步 — 装 Python 依赖 / Install Python dependencies**

```bash
pip install -r requirements.txt
```
依赖：`python-docx`、`matplotlib`、`Pillow`、`pywin32`。

**第 4 步 — 准备一个实验工程 / Prepare an experiment workspace**

按 `lab-report/SKILL.md` 第 ①~④ 步，为你的实验准备一个工程目录：`config.json`（学生信息/数据/常数）、`inputs/`（Word 模板与数据照片）、以及一个实现 `analyze/plot/build` 三个函数的实验适配模块 `experiments/xxx.py`（可参考本仓库示例写法）。然后：

```bash
python scripts\run.py --workdir <你的实验工程目录>
```

输出 `<工程>/outputs/<学号_姓名_实验标题>.docx/.pdf`。

> **注意**：PDF 导出与 `.doc` 模板解析依赖本机可把 docx 转 pdf 的工具（Microsoft Word / LibreOffice）。

## 开箱即用 / Out of the Box

> 适配模块（`experiments/xxx.py`）是为让 agent 自动生成而设计的。把你的实验数据与 Word 模板交给装了本技能的 code 型 agent，agent 会按 `SKILL.md` 骨架自动编写该实验的计算/绘图/排版模块，**无需人工手写代码**；仅对较特殊或冷门的实验，建议人工确认公式与拟合方式。

## 使用提示 / Notes

- **数据识别必须人工核对**：手写数值看不清晰请先确认，避免抄写错误 / Verify handwritten data with the user before computing.
- 换数据后须重跑 `run.py`（公式图与图表需重新生成）/ Re-run `run.py` after changing data.
- **PDF 导出**依赖本机可把 docx 转 pdf 的工具（Microsoft Word / LibreOffice）/ PDF export requires a docx-to-pdf tool on the machine.

## 许可证 / License

[MIT](LICENSE)

Copyright (c) 2026 wangtx-wtx
