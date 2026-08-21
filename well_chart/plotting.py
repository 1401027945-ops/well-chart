# -*- coding: utf-8 -*-
"""绘图模块。

使用 Matplotlib 生成与模板风格一致的曲线图：
- 12×6 英寸、白底、双纵轴（左：油压/套压 MPa；右：瞬时气量 万方/天）；
- 油压 #FF0000、套压 #0000FF（线宽 1.5），瞬时气量 #008000（线宽 1.0）；
- 灰色虚线网格（透明度 0.3），图例右上角，标题为井号；
- 时间轴固定 6 个刻度并自动取整（月初/整日/整点/整分）。
"""

from __future__ import annotations

import io
import math
import os
import tempfile
from pathlib import Path

# 将 Matplotlib 缓存目录指向可写的临时目录，避免权限问题
os.environ.setdefault(
    "MPLCONFIGDIR",
    os.path.join(tempfile.gettempdir(), "matplotlib-cache"),
)

import matplotlib

matplotlib.use("Agg")  # 无界面环境绘图

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FormatStrFormatter, MaxNLocator

from .cleaning import CleaningStats
from .config import (
    CJK_FONT_CANDIDATES,
    COLOR_CASING,
    COLOR_GAS,
    COLOR_OIL,
    AXIS_LINE_WIDTH,
    FIG_DPI,
    FIG_SIZE,
    LABEL_FONT_SIZE,
    LEGEND_FONT_SIZE,
    LINEWIDTH_CASING,
    LINEWIDTH_GAS,
    LINEWIDTH_OIL,
    NUM_TICKS,
    TICK_COLOR,
    TICK_DIRECTION,
    TICK_LABEL_FONT_SIZE,
    TICK_LINE_WIDTH,
    UNIT_GAS,
    UNIT_PRESSURE,
    get_logger,
)

logger = get_logger()

# ---------------------------------------------------------------------------
# 时间刻度：固定 6 个等时间间隔刻度
# ---------------------------------------------------------------------------
DAY = pd.Timedelta(days=1)
HOUR = pd.Timedelta(hours=1)


def build_time_ticks(tmin: pd.Timestamp, tmax: pd.Timestamp, num_ticks: int = NUM_TICKS) -> list:
    """生成 num_ticks 个等时间间隔的刻度。

    刻度从数据最小时间到最大时间均匀分布，保证相邻刻度间隔完全相等
    （即“等时间”），首尾正好落在数据边界上。
    """
    return list(pd.date_range(tmin, tmax, periods=num_ticks))


def tick_label_format(span: pd.Timedelta) -> str:
    """根据时间跨度选择刻度标签格式。"""
    if span >= pd.Timedelta(days=365):
        return "%Y-%m"
    if span >= pd.Timedelta(days=60):
        return "%Y-%m-%d"
    if span >= 2 * DAY:
        return "%m-%d"
    return "%H:%M"


# ---------------------------------------------------------------------------
# 纵坐标
# ---------------------------------------------------------------------------
def nice_ceil(value: float) -> float:
    """把最大值向上取整到“合适的刻度”（整数或 1/2/5×10^n）。"""
    if value < 1:
        return 1.0
    if value <= 10:
        return float(math.ceil(value))
    exp = 10 ** math.floor(math.log10(value))
    for m in (1, 2, 5, 10):
        if value <= m * exp:
            return m * exp
    return 10 * exp


def compute_y_range(series: pd.Series) -> tuple[float, float]:
    """计算纵轴范围：最低 0，最高 = 最大值 + 10% 余量后向上取整。

    若最大值 < 1，则最高值取 1（对应需求 3.3）。
    """
    valid = pd.to_numeric(series, errors="coerce").dropna()
    if valid.empty:
        return 0.0, 1.0
    data_max = float(valid.max())
    top = nice_ceil(data_max * 1.1)
    return 0.0, top


# ---------------------------------------------------------------------------
# 图表生成
# ---------------------------------------------------------------------------
def _setup_chinese_fonts() -> None:
    """设置中文字体，避免图中出现方块乱码。"""
    from matplotlib import font_manager

    # 注册项目自带的免费中文字体（部署到云服务器时保证中文正常显示）
    bundled_font = Path(__file__).resolve().parent.parent / "fonts" / "NotoSansCJKsc-Regular.otf"
    if bundled_font.exists():
        try:
            font_manager.fontManager.addfont(str(bundled_font))
        except Exception:  # noqa: BLE001
            logger.warning("内置中文字体注册失败：%s", bundled_font)

    available = {f.name for f in font_manager.fontManager.ttflist}
    used = None
    for name in CJK_FONT_CANDIDATES:
        if name in available:
            used = name
            break
    current = list(plt.rcParams.get("font.sans-serif", []))
    if used:
        plt.rcParams["font.sans-serif"] = [used] + current
        logger.info("图表字体：%s", used)
    else:
        logger.warning("未找到方正大黑简体，使用系统默认中文字体。")
    plt.rcParams["axes.unicode_minus"] = False


def create_chart(
    df_clean: pd.DataFrame,
    well_name: str,
    stats: CleaningStats,
) -> plt.Figure:
    """根据清洗后的数据绘制曲线图，返回 Figure 对象。"""
    _setup_chinese_fonts()
    times = df_clean["日期"]
    tmin, tmax = times.min(), times.max()
    span = tmax - tmin

    fig, ax = plt.subplots(figsize=FIG_SIZE, dpi=FIG_DPI, facecolor="white")
    ax2 = ax.twinx()  # 右侧纵轴：瞬时气量

    # 三条曲线（对应需求 4.1 的颜色与线宽）
    ax.plot(times, df_clean["油压"], color=COLOR_OIL, linewidth=LINEWIDTH_OIL, label=f"油压{UNIT_PRESSURE}")
    ax.plot(times, df_clean["套压"], color=COLOR_CASING, linewidth=LINEWIDTH_CASING, label=f"套压{UNIT_PRESSURE}")
    ax2.plot(times, df_clean["瞬时气量"], color=COLOR_GAS, linewidth=LINEWIDTH_GAS, label=f"瞬时气量{UNIT_GAS}")

    # 纵轴范围
    oil_lo, oil_hi = compute_y_range(df_clean["油压"])
    casing_lo, casing_hi = compute_y_range(df_clean["套压"])
    left_hi = max(oil_hi, casing_hi)
    ax.set_ylim(0, left_hi)

    gas_lo, gas_hi = compute_y_range(df_clean["瞬时气量"])
    ax2.set_ylim(gas_lo, gas_hi)

    # 纵坐标只显示合适的整数刻度，不保留小数
    for axis in (ax, ax2):
        axis.yaxis.set_major_locator(MaxNLocator(nbins=5, integer=True))
        axis.yaxis.set_major_formatter(FormatStrFormatter("%d"))

    # 瞬时气量全为 0 → 0~1 范围并显示“关井”标注
    gas_all_zero = (
        df_clean["瞬时气量"].notna().any()
        and (df_clean["瞬时气量"].dropna() == 0).all()
    )
    if gas_all_zero:
        ax2.set_ylim(0, 1)
        ax2.text(
            tmin + span * 0.5, 0.5, "关井",
            ha="center", va="center", fontsize=18, color="gray",
            transform=ax2.transData,
        )

    # 时间轴：固定 6 个自动取整刻度
    ticks = build_time_ticks(tmin, tmax)
    ax.set_xlim(tmin, tmax)
    ax.set_xticks(ticks)
    ax.xaxis.set_major_formatter(mdates.DateFormatter(tick_label_format(span)))
    # 刻度线：外部、黑色、0.5 磅；刻度文字 8 号
    for axis in (ax, ax2):
        axis.tick_params(
            axis="both", which="major",
            direction=TICK_DIRECTION, color=TICK_COLOR, width=TICK_LINE_WIDTH,
            length=4, labelsize=TICK_LABEL_FONT_SIZE, labelcolor="black",
        )

    # 坐标轴标签（字号 9）；横轴不显示“日期”文字
    ax.set_ylabel("压力（MPa）", fontsize=LABEL_FONT_SIZE)
    ax2.set_ylabel("瞬时气量（万方/天）", fontsize=LABEL_FONT_SIZE)

    # 不要网格线
    ax.grid(False)
    ax2.grid(False)

    # 图例：居中、只占一行（三条曲线横向排列）
    lines = ax.get_lines() + ax2.get_lines()
    labels = [line.get_label() for line in lines]
    ax.legend(lines, labels, loc="upper center", ncol=3, fontsize=LEGEND_FONT_SIZE, frameon=True)

    # 不显示图标题；去掉最上面的横线（上边框）；坐标轴线黑色 0.5 磅
    ax.spines["top"].set_visible(False)
    ax2.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax2.spines["left"].set_visible(False)
    for spine_name in ("left", "bottom"):
        ax.spines[spine_name].set_color("black")
        ax.spines[spine_name].set_linewidth(AXIS_LINE_WIDTH)
    ax2.spines["right"].set_color("black")
    ax2.spines["right"].set_linewidth(AXIS_LINE_WIDTH)

    fig.tight_layout()
    logger.info("图表生成完成：时间范围 %s ~ %s，刻度=%s", tmin, tmax, ticks)
    return fig


def fig_to_png_bytes(fig: plt.Figure) -> bytes:
    """把 Figure 渲染为 PNG 字节（用于 Streamlit 预览与 Excel 嵌入）。"""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=FIG_DPI, bbox_inches="tight", facecolor="white")
    buf.seek(0)
    return buf.getvalue()


def close_fig(fig: plt.Figure) -> None:
    """释放 Figure 占用的内存。"""
    plt.close(fig)
