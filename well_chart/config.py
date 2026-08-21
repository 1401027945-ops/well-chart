# -*- coding: utf-8 -*-
"""全局配置：图表样式、清洗阈值、日志工具等。

所有与业务规则相关的常量集中在这里，便于统一调整和单元测试。
"""

import logging

# ---------------------------------------------------------------------------
# 日志
# ---------------------------------------------------------------------------
LOGGER_NAME = "well_chart"


def get_logger() -> logging.Logger:
    """返回模块级日志器；未初始化时自动添加控制台输出。"""
    logger = logging.getLogger(LOGGER_NAME)
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        )
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    # 避免重复打印
    logger.propagate = False
    return logger


# ---------------------------------------------------------------------------
# 数据清洗阈值（对应需求 3.1）
# ---------------------------------------------------------------------------
CONSTANT_MIN_POINTS = 6        # 连续不变的数据点达到该数量即视为无效段
CONSTANT_MIN_HOURS = 24.0      # 连续不变持续超过该小时数即视为无效段

# ---------------------------------------------------------------------------
# 图表样式（对应需求 4.1，与模板保持一致）
# ---------------------------------------------------------------------------
FIG_SIZE = (12, 6)             # 图表尺寸（英寸）
FIG_DPI = 150                  # 输出分辨率

COLOR_OIL = "#FF0000"          # 油压：红色
COLOR_CASING = "#0000FF"       # 套压：蓝色
COLOR_GAS = "#FFC000"          # 瞬时气量：橙色（与模板一致）

LINEWIDTH_OIL = 1.5            # 油压线宽
LINEWIDTH_CASING = 1.5         # 套压线宽
LINEWIDTH_GAS = 1.0            # 瞬时气量线宽

GRID_COLOR = "gray"            # 网格颜色：灰色
GRID_ALPHA = 0.3               # 网格透明度
GRID_LINESTYLE = "--"          # 网格线型：虚线

LABEL_FONT_SIZE = 9            # 坐标轴标签字号（7-9 号取 9）
TICK_LABEL_FONT_SIZE = 8       # 刻度文字字号（7-9 号取 8）
LEGEND_FONT_SIZE = 11          # 图例字号（比正文大 2 个字号：9 → 11）
TITLE_FONT_SIZE = 15           # 图标题字号（当前不使用）
NUM_TICKS = 6                  # 时间轴固定刻度数量

# 坐标轴样式
AXIS_LINE_WIDTH = 0.5          # 坐标轴线宽（磅）
TICK_LINE_WIDTH = 0.5          # 刻度线宽（磅）
TICK_DIRECTION = "out"         # 刻度线方向：外部
TICK_COLOR = "black"           # 刻度线颜色：黑色

# 字体：优先方正大黑简体，未安装时回退到系统黑体（SimHei 等）
FONT_NAME = "方正大黑简体"
# 中文字体候选（Windows 优先，Linux 容器需安装 Noto CJK 字体）
CJK_FONT_CANDIDATES = [
    FONT_NAME,
    "SimHei",
    "Microsoft YaHei",
    "Noto Sans CJK SC",
    "WenQuanYi Zen Hei",
    "Source Han Sans SC",
    "PingFang SC",
]

# 生成有效图表所需的最少数据条数（对应需求 7.1）
MIN_ROWS = 10
