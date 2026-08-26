#!/usr/bin/env bash
# lab-report 技能安装脚本 (macOS/Linux)
# 用法:  bash install.sh
# 作用: 装 Python 依赖 -> 语法自检 -> 把技能铺到当前目录的 skills/lab-report
set -euo pipefail
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # lab-report 目录

echo "== 安装 Python 依赖 =="
if ! command -v python3 >/dev/null; then
  echo "未找到 python3，请先安装 Python 3 后重跑"
else
  python3 -m pip install -r "$SRC/requirements.txt" || echo "依赖安装失败，请手动: pip install -r requirements.txt"
fi

echo ""
echo "== 语法自检 =="
if python3 -m py_compile "$SRC"/scripts/core/*.py "$SRC/scripts/run.py"; then
  echo "脚本语法 OK。"
else
  echo "py_compile 有报错，请检查。"
fi

echo ""
echo "== 铺设技能 =="
mkdir -p "$(pwd)/skills"
rm -rf "$(pwd)/skills/lab-report"
cp -R "$SRC" "$(pwd)/skills/lab-report"
find "$(pwd)/skills/lab-report" -type d -name __pycache__ -prune -exec rm -rf {} + 2>/dev/null || true
echo "已铺到: $(pwd)/skills/lab-report"
echo "提示: 如需所有会话全局可用，请把 lab-report 目录复制到你 agent 的 skills 目录。"
echo "提示: 首次生成 PDF 需本机可把 docx 转为 pdf（如 Microsoft Word/LibreOffice）。"