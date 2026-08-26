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

## 使用提示 / Notes

- **数据识别必须人工核对**：手写数值看不清晰请先确认，避免抄写错误 / Verify handwritten data with the user before computing.
- 换数据后须重跑 `run.py`（公式图与图表需重新生成）/ Re-run `run.py` after changing data.
- **PDF 导出**依赖本机可把 docx 转 pdf 的工具（Microsoft Word / LibreOffice）/ PDF export requires a docx-to-pdf tool on the machine.

## 许可证 / License

[MIT](LICENSE)

Copyright (c) 2026 wangtx-wtx