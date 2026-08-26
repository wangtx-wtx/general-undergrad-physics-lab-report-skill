# -*- coding: utf-8 -*-
"""渲染数学公式为透明 PNG（matplotlib mathtext），支持独立公式与正文内联。
统一使用 Computer Modern 数学字体，达到"像书一样"的排版效果。
v2: 修正内联公式垂直对齐（基线与文字一致）。"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os, re
from PIL import Image
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

plt.rcParams['mathtext.fontset'] = 'cm'   # Computer Modern 数学字体

def render_formula(tex, out_path, fontsize=20, dpi=500, color='black'):
    """渲染一个独立公式（mathtext）为透明 PNG。tex 不含 $ 包裹符。"""
    if not tex.startswith('$'):
        tex = '$' + tex + '$'
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.patch.set_alpha(0)
    t = fig.text(0, 0, tex, fontsize=fontsize, color=color)
    fig.canvas.draw()
    bbox = t.get_window_extent(renderer=fig.canvas.get_renderer())
    w_in = bbox.width / fig.dpi; h_in = bbox.height / fig.dpi
    fig.set_size_inches(w_in, h_in)
    t.set_position((0, 0))
    fig.savefig(out_path, dpi=dpi, transparent=True, bbox_inches='tight', pad_inches=0.05)
    plt.close(fig)
    return out_path

def render_inline(tex, out_path, fontsize=18, dpi=500):
    """渲染内联公式。返回 (路径, 像素宽, 像素高)。"""
    full = '$' + tex + '$'
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.patch.set_alpha(0)
    t = fig.text(0, 0, full, fontsize=fontsize)
    fig.canvas.draw()
    bbox = t.get_window_extent(renderer=fig.canvas.get_renderer())
    w_in = bbox.width / fig.dpi; h_in = bbox.height / fig.dpi
    fig.set_size_inches(w_in, h_in)
    t.set_position((0, 0))
    fig.savefig(out_path, dpi=dpi, transparent=True, bbox_inches='tight', pad_inches=0.02)
    plt.close(fig)
    with Image.open(out_path) as im:
        pw, ph = im.size
    return out_path, pw, ph

def set_run_font(run, name='宋体', size=12, bold=False, italic=False):
    run.font.name = name; run.font.size = Pt(size); run.font.bold = bold; run.font.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:hAnsi'), name)

def _adjust_inline_img_baseline(run, pic_height_pt):
    """调整图片 run 内联图片的垂直对齐，使公式基线与文字基线对齐。
    通过设置 wp:inline 的 distB 让图片下沉。
    """
    for child in run._element.iter():
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'inline':
            descent_pt = pic_height_pt * 0.30
            descent_emu = int(descent_pt * 12700)
            child.set(qn('w:distB'), str(descent_emu))
            child.set(qn('w:distT'), '0')
            break

def add_rich_para(doc, text, size=12, name='宋体', align=None, indent_first=True,
                  inline_dir=None, math_h=0.235):
    """富文本段落：text 中形如 $...$ 的片段用文本 run（真上下标+斜体）渲染，
    与正文天然基线对齐；其余为普通文字。"""
    from .mathsym import math_to_runs
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    if indent_first:
        pf.first_line_indent = Pt(size * 2)

    parts = re.split(r'(\$[^$]*\$)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('$') and part.endswith('$') and len(part) > 2:
            tex = part[1:-1]
            math_to_runs(p, tex, size)
        else:
            run = p.add_run(part)
            set_run_font(run, name=name, size=size)
    return p

if __name__ == '__main__':
    import tempfile
    d = os.path.join(tempfile.gettempdir(), 'lab_report_formula_test')
    os.makedirs(d, exist_ok=True)
    render_formula(r"\ln(\theta-\theta_0)=\frac{k}{C}\,t+b",
                   os.path.join(d, 'test.png'))
    print('rendered to', d)
