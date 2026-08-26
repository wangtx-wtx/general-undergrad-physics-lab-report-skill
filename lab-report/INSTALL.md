# lab-report 分发与安装

本技能是**自包含可移植**目录包（`SKILL.md` 指令 + `scripts/` 工具库 + 安装脚本），不依赖任何固定工程目录。解压后放到 agent 的 skills 目录即可使用，任何支持 `SKILL.md` 技能的 agent 环境都通用。

## 给接收者的一句话

> 解压 `lab-report.zip`，进入 `lab-report` 文件夹，运行安装脚本。Windows 用 PowerShell 执行 `install.ps1`，macOS/Linux 执行 `bash install.sh`。安装脚本会自动：装 Python 依赖 → 语法自检，并把技能放到当前目录供你接入。

## 两种安装方式

### 方式一：用自带脚本（快）

解压后进入 `lab-report`，运行：

```powershell
# Windows
.\install.ps1
```

```bash
# macOS / Linux
bash install.sh
```

脚本会安装依赖，并把 `lab-report` 技能铺到**当前目录**的 `skills/lab-report`。如需让所有会话全局可用，把该目录复制到你 agent 的 skills 目录（方式二）。

### 方式二：手动放技能目录（推荐，最干净）

1. 把整个 `lab-report` 目录**复制到你 agent 的 skills 目录**（即 `<agent数据目录>/skills/`，不同 agent 位置不同；新增技能通常自动发现、无需重启）。
2. 安装依赖：

```powershell
pip install -r .\lab-report\requirements.txt
```

## 运行依赖（装依赖后即可）

- Python 3 + `python-docx`、`matplotlib`、`Pillow`、`pywin32`（`requirements.txt` 已列出）
- **Microsoft Word**：PDF 导出与 `.doc` 模板解析使用 Word COM

## 快速自检

```powershell
python scripts\run.py --workdir <任意实验工程目录>
```

输出 `<工程>/outputs/<学号_姓名_实验标题>.docx/.pdf`。

## 目录总览

```
lab-report/
  SKILL.md           指令正文（通用 SKILL.md 技能格式）
  scripts/run.py     可移植流水线入口（--workdir 指定工程目录）
  scripts/core/      通用工具库（docx_helpers / formula / mathsym / template_parser / plotting）
  requirements.txt   Python 依赖
  install.ps1        一键安装（Windows）
  install.sh         一键安装（macOS / Linux）
```