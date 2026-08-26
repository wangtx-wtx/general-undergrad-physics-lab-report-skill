# -*- coding: utf-8 -*-
"""
可移植实验报告一键流水线
=========================
用于任何工作目录，不硬编码路径。core/ 工具库优先取工作区，缺失时自动回退到
本技能捆绑的副本（本文件同目录的 core/）。这样同一份脚本既能在项目里运行，
也能作为 skill 资源被随时调用。

用法:
    python run.py                                    # 在当前目录找 config.json
    python run.py --workdir 某/实验目录                # 指定报告工程目录
    python run.py --workdir 某目录 --config 我的实验.json
    python run.py --workdir 某目录 --experiment experiments/xxx.py

流程: 读取 config → 实验模块 analyze() 计算 → plot() 绘图
     → build() 生成 docx → Word 导出 PDF
输出: <workdir>/outputs/<输出文件名>.docx / .pdf
"""
import os, sys, json, importlib.util, argparse, shutil

# 本文件所在目录 = 脚本包根（可能来自项目，也可能来自 skill 捆绑）
BUNDLE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BUNDLE)
sys.path.insert(0, os.path.join(BUNDLE, 'experiments'))


def ensure_workdir(workdir):
    """确认 / 创建报告工程目录，并确保其中有可用的 core 工具库。"""
    os.makedirs(workdir, exist_ok=True)
    core_dst = os.path.join(workdir, 'core')
    # 优先使用工作区已有的 core/
    if not os.path.isdir(core_dst):
        shutil.copytree(os.path.join(BUNDLE, 'core'), core_dst)
    sys.path.insert(0, workdir)
    return core_dst


def load_module(path):
    """按文件路径加载 Python 模块。"""
    name = os.path.splitext(os.path.basename(path))[0]
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def export_pdf(docx_path, pdf_path):
    """用 Word COM 将 docx 导出为 PDF。"""
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(docx_path)
        doc.ExportAsFixedFormat(pdf_path, 17)   # 17 = PDF
        doc.Close(False)
    finally:
        word.Quit()


def main():
    ap = argparse.ArgumentParser(description='可移植实验报告一键流水线')
    ap.add_argument('--workdir', default=os.getcwd(), help='报告工程目录（默认当前目录）')
    ap.add_argument('--config', default=None, help='实验配置文件，默认 <workdir>/config.json')
    ap.add_argument('--experiment', default=None, help='实验适配模块 .py，默认 <workdir>/experiments/*.py')
    args = ap.parse_args()

    workdir = os.path.abspath(args.workdir)
    core_dst = ensure_workdir(workdir)

    config_path = args.config or os.path.join(workdir, 'config.json')
    exp_path = args.experiment or _default_experiment(workdir)

    out_dir = os.path.join(workdir, 'outputs')
    os.makedirs(out_dir, exist_ok=True)

    # 1. 读配置
    with open(config_path, encoding='utf-8') as f:
        config = json.load(f)
    print('[1/5] 读取配置:', config_path)

    # 2. 加载实验模块
    exp = load_module(exp_path)
    print('[2/5] 加载实验适配:', exp_path)

    # 3. 分析 + 绘图
    results = exp.analyze(config)
    print('[3/5] 数据分析完成')
    if hasattr(exp, 'plot'):
        exp.plot(config, results)
        print('      绘图完成 (core at', core_dst, ')')

    # 4. 生成 docx
    fname = config.get('输出文件名', '实验报告')
    docx_path = os.path.join(out_dir, fname + '.docx')
    exp.build(config, results, docx_path)
    print('[4/5] 报告已生成:', docx_path)

    # 5. 导出 PDF
    pdf_path = os.path.join(out_dir, fname + '.pdf')
    if os.path.exists(pdf_path):
        os.remove(pdf_path)   # 避免被占用报错
    export_pdf(docx_path, pdf_path)
    print('[5/5] PDF 已导出:', pdf_path)
    print('完成！输出目录:', out_dir)


def _default_experiment(workdir):
    """默认实验适配模块：项目里松散放着的 .py，或 experiments/ 目录第一个。"""
    exp_dir = os.path.join(workdir, 'experiments')
    if os.path.isdir(exp_dir):
        for f in sorted(os.listdir(exp_dir)):
            if f.endswith('.py'):
                return os.path.join(exp_dir, f)
        raise SystemExit('experiments/ 下没有 .py 实验模块')
    for f in sorted(os.listdir(workdir)):
        if f.endswith('.py') and not f.startswith('run.py'):
            return os.path.join(workdir, f)
    raise SystemExit('找不到实验适配模块，请用 --experiment 指定')


if __name__ == '__main__':
    main()