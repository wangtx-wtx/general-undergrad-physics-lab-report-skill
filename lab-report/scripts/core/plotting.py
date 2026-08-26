# -*- coding: utf-8 -*-
"""
通用绘图模块（matplotlib）
提供实验报告常用的绘图：散点+线性拟合图、折线图、柱状图。
中文字体自动配置；图片保存为高清 PNG。
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams

# 中文字体
rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun']
rcParams['axes.unicode_minus'] = False
rcParams['mathtext.fontset'] = 'cm'

def _linfit(x, y):
    import math
    n = len(x)
    sx = sum(x); sy = sum(y)
    sxx = sum(v*v for v in x); sxy = sum(x[i]*y[i] for i in range(n))
    syy = sum(v*v for v in y)
    a = (n*sxy - sx*sy) / (n*sxx - sx*sx)
    b = (sy - a*sx) / n
    r = (n*sxy - sx*sy) / math.sqrt((n*sxx - sx*sx) * (n*syy - sy*sy))
    return a, b, r

def plot_fit_scatter(x, y_list, labels, out_path, xlabel='', ylabel='',
                     title='', fit=True, figsize=(7.5, 5.0), dpi=150,
                     annotate_slope=True):
    """散点 + 线性拟合图（多条线）。y_list: 多个数据序列; labels: 各序列名。
    返回 (斜率列表, 相关系数列表)。"""
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    colors = ['red', 'blue', 'green', 'orange', 'purple']
    markers = ['o', 's', '^', 'D', 'v']
    linestyles = ['-', '--', '-.', ':', '-']
    slopes, rs = [], []
    for i, (y, lab) in enumerate(zip(y_list, labels)):
        a, b, r = _linfit(x, y)
        slopes.append(a); rs.append(r)
        ax.scatter(x, y, color=colors[i % len(colors)], marker=markers[i % len(markers)],
                   s=45, label=lab + ' (实验点)', zorder=3)
        if fit:
            xs = [min(x), max(x)]
            ax.plot(xs, [a*xs[0]+b, a*xs[1]+b], color=colors[i % len(colors)],
                    linestyle=linestyles[i % len(linestyles)], linewidth=1.6,
                    label=lab + f' 拟合: y=({a:.4f})x+({b:.4f})', zorder=2)
    ax.set_xlabel(xlabel, fontsize=12)
    ax.set_ylabel(ylabel, fontsize=12)
    if title:
        ax.set_title(title, fontsize=13)
    ax.grid(True, linestyle=':', alpha=0.6)
    ax.legend(fontsize=9, loc='upper right', framealpha=0.9)
    if annotate_slope and slopes:
        txt = '\n'.join(f'{lab}: 斜率={abs(slopes[i]):.4f}, r={rs[i]:.4f}'
                        for i, lab in enumerate(labels))
        ax.text(0.02, 0.04, txt, transform=ax.transAxes, fontsize=9,
                bbox=dict(boxstyle='round', fc='white', alpha=0.85, ec='gray'))
    fig.tight_layout()
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    fig.savefig(out_path, dpi=dpi)
    plt.close(fig)
    return slopes, rs

def plot_lines(x, y_list, labels, out_path, xlabel='', ylabel='', title='',
               figsize=(7.5, 5.0), dpi=150):
    """多条折线图（无拟合）。"""
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    colors = ['red', 'blue', 'green', 'orange', 'purple']
    markers = ['o', 's', '^', 'D', 'v']
    for i, (y, lab) in enumerate(zip(y_list, labels)):
        ax.plot(x, y, color=colors[i % len(colors)], marker=markers[i % len(markers)],
                linewidth=1.6, label=lab)
    ax.set_xlabel(xlabel, fontsize=12)
    ax.set_ylabel(ylabel, fontsize=12)
    if title:
        ax.set_title(title, fontsize=13)
    ax.grid(True, linestyle=':', alpha=0.6)
    ax.legend(fontsize=9)
    fig.tight_layout()
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    fig.savefig(out_path, dpi=dpi)
    plt.close(fig)

def plot_bar(categories, values, out_path, xlabel='', ylabel='', title='',
             figsize=(7.5, 5.0), dpi=150):
    """柱状图。"""
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    ax.bar(categories, values, color='steelblue', edgecolor='black')
    ax.set_xlabel(xlabel, fontsize=12)
    ax.set_ylabel(ylabel, fontsize=12)
    if title:
        ax.set_title(title, fontsize=13)
    ax.grid(True, axis='y', linestyle=':', alpha=0.6)
    for i, v in enumerate(values):
        ax.text(i, v + max(values)*0.02, f'{v:.3f}', ha='center', fontsize=9)
    fig.tight_layout()
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    fig.savefig(out_path, dpi=dpi)
    plt.close(fig)

if __name__ == '__main__':
    # 自测
    import math, os, tempfile
    x = list(range(0, 19))
    y1 = [math.log(40.3-20.8 - 1.0*i) for i in range(19)]
    y2 = [math.log(45.4-20.8 - 1.0*i) for i in range(19)]
    slopes, rs = plot_fit_scatter(x, [y1, y2], ['纯净水', '饱和食盐水'],
                                  os.path.join(tempfile.gettempdir(), '_plot_test.png'),
                                  xlabel='$t$ / min', ylabel='$\\ln(\\theta-\\theta_0)$',
                                  title='测试图')
    print('自测完成', slopes, rs)
