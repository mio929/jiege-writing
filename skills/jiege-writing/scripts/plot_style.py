# -*- coding: utf-8 -*-
"""中文论文插图统一样式。

三条硬规则：

1. **字体**：英文与数字用 Times New Roman，中文用宋体（SimSun）。靠 matplotlib 的
   font fallback 实现——按字符逐个查字体，缺字形就落到下一个。
2. **字号**：轴标签、刻度标签、图例同一套，不随画布缩放漂移。图题不在这里，
   它属于正文段落，由 build_docx.py 按五号宋体加粗排。
3. **轴对齐**：spine 固定在轴范围起点（`('axes', 0)`），不做 outward 外扩；
   刻度朝内、只在左与下出现。移动 spine 而不同步移动刻度，是「轴线与刻度线错位」
   最常见的原因。

用法：

    from plot_style import apply_style, new_axes, save, audit

    apply_style()
    fig, ax = new_axes()
    ax.plot(x, y, color="#333333", lw=1.2)
    ax.set_xlabel("灌溉定额 / (m³·hm⁻²)")
    ax.set_ylabel("土壤含盐量 / (g·kg⁻¹)")
    ax.legend(frameon=False)
    save(fig, "fig1_salt_profile.png")
    print(audit(ax))
"""

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib import rcParams

EN_FONT = "Times New Roman"
CN_FONT = "SimSun"


def apply_style(base_pt=9.0):
    """设定全局字体与线宽。base_pt 按最终图宽 8–9 cm 估，9 pt 在 Word 里约等于小五。"""
    try:
        rcParams["font.family"] = [EN_FONT, CN_FONT]
    except Exception:
        rcParams["font.family"] = "serif"
        rcParams["font.serif"] = [EN_FONT, CN_FONT]

    rcParams["mathtext.fontset"] = "stix"
    rcParams["axes.unicode_minus"] = False  # 负号用 ASCII，避免方框

    rcParams["font.size"] = base_pt
    rcParams["axes.labelsize"] = base_pt
    rcParams["axes.titlesize"] = base_pt
    rcParams["xtick.labelsize"] = base_pt - 0.5
    rcParams["ytick.labelsize"] = base_pt - 0.5
    rcParams["legend.fontsize"] = base_pt - 0.5

    rcParams["axes.linewidth"] = 0.8
    rcParams["xtick.major.width"] = 0.8
    rcParams["ytick.major.width"] = 0.8
    rcParams["xtick.major.size"] = 3.2
    rcParams["ytick.major.size"] = 3.2
    rcParams["xtick.minor.size"] = 1.8
    rcParams["ytick.minor.size"] = 1.8

    rcParams["figure.dpi"] = 300
    rcParams["savefig.dpi"] = 600
    rcParams["savefig.bbox"] = "tight"
    rcParams["savefig.pad_inches"] = 0.02


def new_axes(width_cm=8.5, height_cm=6.0, include_zero=True):
    """建一张单轴图，spine 与刻度已对齐，坐标轴默认包含原点。

    include_zero=False 仅用于数据整体远离 0 的情形（pH、温度、水深），
    且须在图注里说明轴已截断。柱状图不能关——柱长比例会失真。
    """
    fig, ax = plt.subplots(figsize=(width_cm / 2.54, height_cm / 2.54))

    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        # 固定在轴范围起点，不随数据平移，也不 outward 外扩
        ax.spines[side].set_position(("axes", 0))

    ax.tick_params(axis="both", which="both", direction="in",
                   top=False, right=False, pad=2.5)
    ax.xaxis.set_ticks_position("bottom")
    ax.yaxis.set_ticks_position("left")

    ax._include_zero = include_zero
    return fig, ax


def finalize(ax):
    """出图前定界：把轴拉到包含 0，再留一点余量。对数轴跳过。"""
    if getattr(ax, "_include_zero", True):
        if ax.get_xscale() != "log":
            lo, hi = ax.get_xlim()
            if lo > 0:
                ax.set_xlim(left=0)
        if ax.get_yscale() != "log":
            lo, hi = ax.get_ylim()
            if lo > 0:
                ax.set_ylim(bottom=0)
    return ax


def save(fig, path):
    """出图前定界（把轴拉到含 0），再按 600 dpi 白底输出。"""
    for ax in fig.axes:
        finalize(ax)
    fig.savefig(path, dpi=600, facecolor="white", edgecolor="none")
    plt.close(fig)
    return path


def audit(ax):
    """检查轴与刻度是否可能错位、轴是否缺少原点，返回问题列表；无问题返回空列表。

    会先调 finalize 定界，所以「轴未含 0」这类问题能直接报出来。
    """
    finalize(ax)
    problems = []

    for name in ("x", "y"):
        if name == "x":
            lo, _ = ax.get_xlim()
            scale = ax.get_xscale()
        else:
            lo, _ = ax.get_ylim()
            scale = ax.get_yscale()
        if scale != "log" and lo > 0 and getattr(ax, "_include_zero", True):
            problems.append(
                "%s 轴从 %.3g 起、未含 0；柱状图柱长会失真，折线图易被误读" % (name, lo))

    for side in ("left", "bottom"):
        pos = ax.spines[side].get_position()
        # 默认对齐状态只有这两种写法是安全的
        if pos not in [("outward", 0), ("axes", 0), ("data", 0)]:
            problems.append("%s spine 位置为 %r，可能偏离刻度线" % (side, pos))

    if ax.xaxis.get_ticks_position() != "bottom":
        problems.append("x 轴刻度不在底部，与 spine 可能不重合")
    if ax.yaxis.get_ticks_position() != "left":
        problems.append("y 轴刻度不在左侧，与 spine 可能不重合")

    for name, axis in (("x", ax.xaxis), ("y", ax.yaxis)):
        for tick in axis.get_major_ticks():
            if tick.tick1line.get_markersize() <= 0:
                problems.append("%s 轴刻度线长度为 0，看不见刻度" % name)
                break

    if len(ax.spines) > 4 or getattr(ax, "right_ax", None) is not None:
        problems.append("检测到双轴，右轴刻度需单独与右 spine 对齐")

    if ax.get_title():
        problems.append("图内有标题，学术论文应把标题写进图题（如「图 1  盐分随灌溉定额的变化」），图内留空")

    return problems
