# -*- coding: utf-8 -*-
# lab-report 技能安装脚本 (Windows / PowerShell)
# 用法: 在 PowerShell 里执行  .\install.ps1
# 作用: 装 Python 依赖 -> 语法自检 -> 把技能铺到当前目录的 skills/lab-report
$ErrorActionPreference = 'Stop'
$src = (Split-Path -Parent $PSScriptRoot)   # lab-report 目录

Write-Host '== 安装 Python 依赖 =='
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
  Write-Warning '未找到 python，请先安装 Python 3 后重跑'
} else {
  python -m pip install -r (Join-Path $src 'requirements.txt')
  if ($LASTEXITCODE -ne 0) { Write-Warning '依赖安装失败，请手动执行: pip install -r requirements.txt' }
}

Write-Host ''
Write-Host '== 语法自检 =='
$pyFiles = @(Get-ChildItem -Path (Join-Path $src 'scripts') -Recurse -Filter '*.py' -File)
$err = $false
foreach ($f in $pyFiles) {
  python -m py_compile $f.FullName 2>$null
  if ($LASTEXITCODE -ne 0) { Write-Warning ("编译失败: " + $f.Name); $err = $true }
}
if (-not $err) { Write-Host '脚本语法 OK。' } else { Write-Warning 'py_compile 有报错，请检查 Python/脚本。' }

Write-Host ''
Write-Host '== 铺设技能 =='
$dst = Join-Path (Get-Location) 'skills\lab-report'
New-Item -ItemType Directory -Path (Join-Path (Get-Location) 'skills') -Force | Out-Null
if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
Copy-Item -Recurse -Force $src $dst
Get-ChildItem $dst -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
  ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host ("已铺到: " + $dst)
Write-Host '提示: 如需所有会话全局可用，请把 lab-report 目录复制到你 agent 的 skills 目录。'
Write-Host '提示: 首次生成 PDF 需在本机安装 Microsoft Word (Word COM)。'