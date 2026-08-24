#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""基于三问 Excel 求解结果重绘人员排班图。

默认目录结构：
    运行结果_20260819/
    ├── 问题一优化结果_人数均衡.xlsx
    ├── 问题二优化结果.xlsx
    ├── 问题三联合优化结果.xlsx
    └── 论文重绘图/redraw_staffing_figures.py

运行示例：
    python redraw_staffing_figures.py
    python redraw_staffing_figures.py --data-dir "..." --output-dir "..."
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from pathlib import Path


def _bootstrap_workspace_dependencies() -> None:
    """在 Codex 工作区中优先加载本地依赖；普通 Python 环境可直接忽略。"""
    script_dir = Path(__file__).resolve().parent
    workspace_root = script_dir.parent.parent
    dependency_dir = workspace_root / ".runtime_deps"
    if dependency_dir.is_dir():
        sys.path.insert(0, str(dependency_dir))


_bootstrap_workspace_dependencies()

# 使服务器、命令行和无 GUI 环境都能稳定渲染，并避免用户配置目录不可写。
os.environ.setdefault("MPLBACKEND", "Agg")
os.environ.setdefault(
    "MPLCONFIGDIR",
    str(Path(tempfile.gettempdir()) / "staffing_figure_mplconfig"),
)

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager


WIDTH_MM = 183.0
HEIGHT_MM = 95.0
MM_PER_INCH = 25.4
FIGSIZE = (WIDTH_MM / MM_PER_INCH, HEIGHT_MM / MM_PER_INCH)

COLORS = {
    "q1": "#3A6EA5",       # 蓝色：问题一
    "q2": "#D97706",       # 橙色：问题二
    "q3": "#128277",       # 青绿色：问题三
    "surplus": "#F2D98B",  # 低饱和黄：制度性冗余
    "grid": "#D8DEE6",
    "text": "#20252B",
}


def choose_chinese_font() -> str:
    """选择支持中文的字体，保证 PNG/SVG/PDF 不出现乱码。"""
    available = {item.name for item in font_manager.fontManager.ttflist}
    candidates = [
        "Microsoft YaHei",
        "Microsoft YaHei UI",
        "SimHei",
        "Noto Sans CJK SC",
        "Source Han Sans SC",
        "Arial Unicode MS",
    ]
    for family in candidates:
        if family in available:
            return family
    # 保留日志警告；后续输出仍可用 DejaVu Sans 生成。
    print("警告：未检测到常用中文字体，请安装微软雅黑或思源黑体。")
    return "DejaVu Sans"


def configure_style() -> str:
    """应用面向论文印刷的统一样式。"""
    chinese_font = choose_chinese_font()
    mpl.rcParams.update(
        {
            "font.family": "sans-serif",
            "font.sans-serif": [chinese_font, "Arial", "DejaVu Sans"],
            "font.size": 8.0,
            "axes.titlesize": 11.0,
            "axes.titleweight": "semibold",
            "axes.labelsize": 9.0,
            "axes.labelweight": "normal",
            "xtick.labelsize": 7.5,
            "ytick.labelsize": 7.5,
            "legend.fontsize": 7.5,
            "axes.unicode_minus": False,
            "axes.edgecolor": "#707780",
            "axes.labelcolor": COLORS["text"],
            "xtick.color": COLORS["text"],
            "ytick.color": COLORS["text"],
            "text.color": COLORS["text"],
            "figure.facecolor": "white",
            "axes.facecolor": "white",
            "svg.fonttype": "none",   # SVG 保留可编辑文字
            "pdf.fonttype": 42,        # PDF 嵌入 TrueType 文字
            "savefig.dpi": 600,
        }
    )
    return chinese_font


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="重绘三问人员排班对比图")
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=script_dir.parent,
        help="包含三个 Excel 结果文件的目录",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=script_dir,
        help="图片输出目录",
    )
    return parser.parse_args()


def _assert_daily_index(frame: pd.DataFrame, label: str) -> None:
    days = frame["天"].to_numpy()
    expected = np.arange(1, 31)
    if len(frame) != 30:
        raise ValueError(f"{label}应有30行每日结果，实际为{len(frame)}行。")
    if frame["天"].duplicated().any():
        raise ValueError(f"{label}存在重复日期。")
    if not np.array_equal(days, expected):
        raise ValueError(f"{label}的日期必须为按顺序排列的1—30天。")


def _assert_integer_complete(frame: pd.DataFrame, columns: list[str], label: str) -> None:
    values = frame[columns]
    if values.isna().any().any():
        raise ValueError(f"{label}存在缺失值。")
    array = values.to_numpy(dtype=float)
    if not np.isfinite(array).all():
        raise ValueError(f"{label}存在非有限数值。")
    if not np.allclose(array, np.rint(array)):
        raise ValueError(f"{label}中的人数或日期必须为整数。")


def load_and_validate(data_dir: Path) -> pd.DataFrame:
    """读取三问真实结果，并在绘图前执行完整性和跨文件一致性校验。"""
    file_q1 = data_dir / "问题一优化结果_人数均衡.xlsx"
    file_q2 = data_dir / "问题二优化结果.xlsx"
    file_q3 = data_dir / "问题三联合优化结果.xlsx"
    for path in (file_q1, file_q2, file_q3):
        if not path.is_file():
            raise FileNotFoundError(f"未找到数据文件：{path}")

    q1 = pd.read_excel(file_q1, sheet_name="每日汇总", usecols=["天", "最优总人数"])
    q2 = pd.read_excel(file_q2, sheet_name="每日汇总", usecols=["天", "最优总人数"])
    q3 = pd.read_excel(
        file_q3,
        sheet_name="每日人员与检验",
        usecols=["天", "问题二条件下最低人数", "第三问实际人数", "额外人数"],
    )

    for frame, label in ((q1, "问题一"), (q2, "问题二"), (q3, "问题三")):
        frame.sort_values("天", inplace=True)
        frame.reset_index(drop=True, inplace=True)
        _assert_daily_index(frame, label)
        _assert_integer_complete(frame, list(frame.columns), label)

    q1 = q1.rename(columns={"最优总人数": "问题一最低日用工"})
    q2 = q2.rename(columns={"最优总人数": "问题二最低需求"})
    merged = q1.merge(q2, on="天", validate="one_to_one").merge(q3, on="天", validate="one_to_one")

    if not np.array_equal(
        merged["问题二最低需求"].to_numpy(),
        merged["问题二条件下最低人数"].to_numpy(),
    ):
        raise ValueError("问题二与问题三工作簿中的‘问题二最低人数’不一致。")

    calculated_surplus = merged["第三问实际人数"] - merged["问题二条件下最低人数"]
    if not np.array_equal(calculated_surplus.to_numpy(), merged["额外人数"].to_numpy()):
        raise ValueError("问题三的‘实际人数−最低人数’与‘额外人数’不一致。")
    if (calculated_surplus < 0).any():
        raise ValueError("问题三实际出勤人数不应低于问题二最低需求。")

    return merged


def style_axis(ax: plt.Axes) -> None:
    ax.set_xlim(1, 30)
    ax.set_xticks([1, 5, 10, 15, 20, 25, 30])
    ax.set_ylim(280, 620)
    ax.set_yticks(np.arange(300, 601, 50))
    ax.set_xlabel("日期/天")
    ax.set_ylabel("人数/人")
    ax.grid(axis="y", color=COLORS["grid"], linewidth=0.55, alpha=0.72)
    ax.grid(axis="x", visible=False)
    ax.tick_params(axis="both", direction="out", length=3.0, width=0.7)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_linewidth(0.8)
    ax.spines["bottom"].set_linewidth(0.8)
    ax.set_axisbelow(True)


def plot_daily_demand_vs_attendance(data: pd.DataFrame) -> plt.Figure:
    days = data["天"].to_numpy()
    minimum = data["问题二条件下最低人数"].to_numpy()
    actual = data["第三问实际人数"].to_numpy()

    fig, ax = plt.subplots(figsize=FIGSIZE, layout="constrained")
    fill = ax.fill_between(
        days,
        minimum,
        actual,
        color=COLORS["surplus"],
        alpha=0.48,
        linewidth=0,
        label="制度性冗余",
        zorder=1,
    )
    line_q2, = ax.plot(
        days,
        minimum,
        color=COLORS["q2"],
        linewidth=1.65,
        linestyle="--",
        marker="s",
        markersize=3.0,
        markerfacecolor="white",
        markeredgewidth=0.8,
        markevery=1,
        label="问题二最低需求",
        zorder=3,
    )
    line_q3, = ax.plot(
        days,
        actual,
        color=COLORS["q3"],
        linewidth=1.8,
        linestyle="-",
        marker="o",
        markersize=3.0,
        markerfacecolor="white",
        markeredgewidth=0.8,
        markevery=1,
        label="问题三实际出勤",
        zorder=4,
    )

    style_axis(ax)
    ax.set_title("连续工日约束下每日需求与实际出勤", pad=8)
    ax.legend(
        handles=[line_q2, line_q3, fill],
        loc="upper right",
        ncol=3,
        frameon=False,
        handlelength=2.4,
        columnspacing=1.25,
        handletextpad=0.55,
        borderaxespad=0.5,
    )
    return fig


def plot_three_question_comparison(data: pd.DataFrame) -> plt.Figure:
    days = data["天"].to_numpy()

    fig, ax = plt.subplots(figsize=FIGSIZE, layout="constrained")
    ax.plot(
        days,
        data["问题一最低日用工"],
        color=COLORS["q1"],
        linewidth=1.55,
        linestyle="-.",
        marker="o",
        markersize=2.9,
        markerfacecolor="white",
        markeredgewidth=0.75,
        label="问题一：最低日用工",
        zorder=2,
    )
    ax.plot(
        days,
        data["问题二最低需求"],
        color=COLORS["q2"],
        linewidth=1.65,
        linestyle="--",
        marker="s",
        markersize=2.9,
        markerfacecolor="white",
        markeredgewidth=0.75,
        label="问题二：时限下最低需求",
        zorder=3,
    )
    ax.plot(
        days,
        data["第三问实际人数"],
        color=COLORS["q3"],
        linewidth=1.8,
        linestyle="-",
        marker="^",
        markersize=3.2,
        markerfacecolor="white",
        markeredgewidth=0.8,
        label="问题三：实际出勤",
        zorder=4,
    )

    style_axis(ax)
    ax.set_title("三问每日人员规模对比", pad=8)
    ax.legend(
        loc="upper right",
        ncol=3,
        frameon=False,
        handlelength=2.4,
        columnspacing=1.2,
        handletextpad=0.55,
        borderaxespad=0.5,
    )
    return fig


def save_figure(fig: plt.Figure, output_stem: Path) -> None:
    """同时保存高分辨率预览图和可编辑矢量图。"""
    fig.savefig(output_stem.with_suffix(".png"), dpi=600, facecolor="white")
    fig.savefig(output_stem.with_suffix(".tiff"), dpi=600, facecolor="white")
    fig.savefig(output_stem.with_suffix(".svg"), facecolor="white")
    fig.savefig(output_stem.with_suffix(".pdf"), facecolor="white")
    plt.close(fig)


def main() -> None:
    args = parse_args()
    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    font_name = configure_style()
    data = load_and_validate(data_dir)

    fig1 = plot_daily_demand_vs_attendance(data)
    save_figure(fig1, output_dir / "图1_连续工日需求与实际出勤")

    fig2 = plot_three_question_comparison(data)
    save_figure(fig2, output_dir / "图2_三问每日人员规模对比")

    surplus = data["额外人数"]
    print(f"中文字体：{font_name}")
    print(f"数据校验通过：{len(data)}天，日期范围{data['天'].min()}—{data['天'].max()}。")
    print(f"制度性冗余范围：{surplus.min()}—{surplus.max()}人。")
    print(f"图片已输出至：{output_dir}")


if __name__ == "__main__":
    main()
