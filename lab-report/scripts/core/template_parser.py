# -*- coding: utf-8 -*-
"""
通用模板解析：读取实验报告 Word 模板（.doc/.docx），提取
标题、章节结构、数据表格（含表头/待填格）、说明文字（带颜色要求）。

用法:
    from core.template_parser import parse_template
    info = parse_template(r'...模板.doc')
    info['标题']          # 报告标题
    info['章节']          # [{'标题': '一、实验目的', '段落': [...], '要求': [...]}, ...]
    info['表格']          # [{'标题': '2. 饱和食盐水温度随时间的变化', '行': [['t(min)','0','1',...], ...], '结构': (6,8)}, ...]
    info['说明']          # 带颜色的说明文字（模板要求，正式报告需删除）
"""
import os, re, json

def _word_com():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    return word

def _read_doc(path):
    """用 Word COM 读取 .doc/.docx 的段落与表格。返回 (段落列表, 表格列表)。"""
    ext = os.path.splitext(path)[1].lower()
    word = _word_com()
    doc = word.Documents.Open(path, False, True)   # read-only
    paras = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        txt = p.Range.Text.replace('\r', '').replace('\n', '').replace('\x07', '')
        color = None
        try:
            if p.Range.Font.Color != -16777216:  # 非自动色（带颜色=说明文字）
                color = p.Range.Font.Color
        except Exception:
            pass
        paras.append({'text': txt, 'color': color})
    tables = []
    for t in range(1, doc.Tables.Count + 1):
        tb = doc.Tables(t)
        rows = []
        for r in range(1, tb.Rows.Count + 1):
            row = []
            for c in range(1, tb.Columns.Count + 1):
                try:
                    cell = tb.Cell(r, c).Range.Text.replace('\r', '').replace('\n', '').replace('\x07', '')
                except Exception:
                    cell = ''
                row.append(cell)
            rows.append(row)
        tables.append({'行': rows, '结构': (len(rows), len(rows[0]) if rows else 0)})
    doc.Close(False)
    word.Quit()
    return paras, tables

def _is_heading(text):
    """判断是否为章节标题（如 一、实验目的 / 1. 作图）。"""
    t = text.strip()
    if re.match(r'^[一二三四五六七八九十]+、', t):
        return True
    if re.match(r'^\d+[\.、]', t) and len(t) < 30:
        return True
    return False

def parse_template(path):
    """解析模板文件，返回结构化信息。"""
    paras, tables = _read_doc(path)

    # 标题 = 第一个非空非说明段落
    title = ''
    for p in paras:
        t = p['text'].strip()
        if t and '实验' in t:
            title = t
            break

    # 章节划分
    sections = []
    cur = None
    for p in paras:
        t = p['text'].strip()
        if not t:
            continue
        if _is_heading(t) and ('、' in t or re.match(r'^\d+[\.、]', t)):
            cur = {'标题': t, '段落': [], '要求': []}
            sections.append(cur)
        else:
            if cur is None:
                cur = {'标题': '(前言)', '段落': [], '要求': []}
                sections.append(cur)
            if p['color'] is not None:
                cur['要求'].append(t)
            else:
                cur['段落'].append(t)

    # 说明文字（模板末尾的注意事项等）
    notes = [p['text'] for p in paras if p['color'] is not None and p['text'].strip()]

    return {'标题': title, '章节': sections, '表格': tables, '说明': notes}

def parse_requirements(template_path):
    """快速提取模板中的填写要求（带颜色说明文字）。"""
    info = parse_template(template_path)
    reqs = []
    for sec in info['章节']:
        for r in sec['要求']:
            reqs.append(f"[{sec['标题']}] {r}")
    reqs.extend(info['说明'])
    return reqs


if __name__ == '__main__':
    import sys, os
    # 自测：需传入一个模板文件的路径作为第一个命令行参数
    p = sys.argv[1] if len(sys.argv) > 1 else None
    if not p or not os.path.exists(p):
        print('用法: python -m core.template_parser <模板文件路径>')
        sys.exit(0)
    info = parse_template(os.path.abspath(p))
    print('标题:', info['标题'])
    print('章节:')
    for s in info['章节']:
        print(' ', s['标题'], '| 段落数:', len(s['段落']), '| 要求:', s['要求'][:2])
    print('表格数:', len(info['表格']))
    for tb in info['表格']:
        print('  结构', tb['结构'], '首行:', tb['行'][0][:5])
    print('说明:', info['说明'][:2])
