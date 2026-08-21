# -*- coding: utf-8 -*-
"""数据清洗模块。

核心处理逻辑（对应需求 3.1）：
1. 油压/套压/瞬时气量为负 → 替换为 NaN（瞬时气量的 0 值保留，可能是真实关井）；
2. 油压/套压连续不变超过阈值（6 个点或 24 小时）→ 视为无效段，
   用前后有效值的线性插值替代；无法插值（数据段在首尾）则保留空值；
3. 每列输出“插值标记”：0=原始有效，1=插值修复，2=无法修复（空值）。

所有统计结果集中到 CleaningStats，便于界面展示和日志记录。
"""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from .config import CONSTANT_MIN_HOURS, CONSTANT_MIN_POINTS, get_logger

logger = get_logger()

# 插值标记含义
FLAG_ORIGINAL = 0      # 原始有效值
FLAG_INTERPOLATED = 1  # 插值修复
FLAG_UNFIXED = 2       # 无法修复（保留空值）

FLAG_NAMES = {
    FLAG_ORIGINAL: "原始有效",
    FLAG_INTERPOLATED: "插值修复",
    FLAG_UNFIXED: "无法修复",
}

PRESSURE_COLUMNS = ("油压", "套压")   # 需要固定值检测的列
GAS_COLUMN = "瞬时气量"


@dataclass
class CleaningStats:
    """数据清洗与统计结果。"""

    total_rows: int = 0
    time_min: pd.Timestamp | None = None
    time_max: pd.Timestamp | None = None
    frequency: str = "未知"
    negatives: dict = field(default_factory=dict)      # 各列负值个数
    zero_gas_rows: int = 0                             # 瞬时气量为 0 的行数
    constant_runs: dict = field(default_factory=dict)  # 各列固定值段数
    interpolated: dict = field(default_factory=dict)   # 各列插值点数
    unfixed: dict = field(default_factory=dict)        # 各列无法修复点数
    warnings: list = field(default_factory=list)       # 处理警告（界面展示）

    @property
    def total_anomalies(self) -> int:
        """异常值总数：负值 + 无法修复 + 插值（固定值段按点数估算）。"""
        neg = sum(self.negatives.values())
        unfix = sum(self.unfixed.values())
        interp = sum(self.interpolated.values())
        return int(neg + unfix + interp)


def detect_constant_runs(
    values: pd.Series,
    times: pd.Series,
    min_points: int = CONSTANT_MIN_POINTS,
    min_hours: float = CONSTANT_MIN_HOURS,
) -> tuple[np.ndarray, list]:
    """检测连续不变的数据段。

    返回 (无效掩码, 段列表)；每段记录 (起始下标, 结束下标, 固定值, 点数, 持续小时数)。
    NaN 会切断连续段，因此负值替换为 NaN 后不会参与固定值判断。
    """
    arr = values.to_numpy(dtype=float)
    times_np = times.to_numpy()
    n = len(arr)
    invalid = np.zeros(n, dtype=bool)
    runs: list = []
    i = 0
    while i < n:
        if not np.isfinite(arr[i]):
            i += 1
            continue
        j = i
        while j + 1 < n and np.isfinite(arr[j + 1]) and arr[j + 1] == arr[i]:
            j += 1
        run_len = j - i + 1
        run_hours = 0.0
        if j > i:
            run_hours = (times_np[j] - times_np[i]) / np.timedelta64(1, "h")
        # 连续不变超过阈值（点数或小时数任一满足）→ 无效段
        if run_len >= min_points or run_hours >= min_hours:
            invalid[i : j + 1] = True
            runs.append((int(i), int(j), float(arr[i]), int(run_len), float(run_hours)))
        i = j + 1
    return invalid, runs


def interpolate_series(
    values: pd.Series,
    times: pd.Series,
    flags: np.ndarray,
) -> tuple[pd.Series, np.ndarray, int]:
    """对缺失值（NaN）按时间线性插值。

    只对“前后都有有效锚点”的点进行插值；数据首尾无法找到前后锚点的
    点保留 NaN，并标记为无法修复。
    返回 (插值后的序列, 更新后的标记数组, 无法修复点数)。
    """
    cleaned = values.astype(float).copy()
    missing = np.isnan(cleaned.to_numpy())
    if not missing.any():
        return cleaned, flags, 0

    x_all = times.astype("int64").to_numpy() / 1e9  # 转成秒，便于插值
    valid_idx = np.where(~missing)[0]
    if len(valid_idx) == 0:
        # 没有任何有效锚点，全部无法修复
        flags[missing] = FLAG_UNFIXED
        return cleaned, flags, int(missing.sum())

    x_valid = x_all[valid_idx]
    y_valid = cleaned.to_numpy()[valid_idx]
    first_valid, last_valid = x_valid.min(), x_valid.max()

    unfixed_count = 0
    for k in np.where(missing)[0]:
        t = x_all[k]
        if t < first_valid or t > last_valid:
            # 位于所有有效锚点之前/之后，无法插值
            flags[k] = FLAG_UNFIXED
            unfixed_count += 1
            continue
        # 用 numpy.interp 在前后锚点之间线性插值
        cleaned.iloc[k] = np.interp(t, x_valid, y_valid)
        flags[k] = FLAG_INTERPOLATED
    return cleaned, flags, unfixed_count


def estimate_frequency(times: pd.Series) -> str:
    """根据相邻时间差的中位数估计采样频率（分钟级/小时级/天级）。"""
    uniq = pd.Series(times.unique()).sort_values()
    if len(uniq) < 2:
        return "未知"
    diffs = uniq.diff().dropna().dt.total_seconds()
    median_s = float(diffs.median())
    if median_s < 2700:      # < 45 分钟
        return "分钟级"
    if median_s < 72000:     # < 20 小时（涵盖小时级与分钟级以上的亚日数据）
        return "小时级"
    return "天级"


def clean_well_data(df: pd.DataFrame) -> tuple[pd.DataFrame, CleaningStats]:
    """执行完整清洗流程，返回 (清洗后的 DataFrame, 统计信息)。

    输入 df 必须包含列：日期（datetime64）、油压、套压、瞬时气量。
    输出 DataFrame 列：
    日期 / 油压_原始 / 套压_原始 / 瞬时气量_原始 /
    油压 / 套压 / 瞬时气量 / 油压_插值标记 / 套压_插值标记 / 瞬时气量_插值标记
    """
    df = df.copy()
    times = df["日期"].reset_index(drop=True)
    stats = CleaningStats(
        total_rows=len(df),
        time_min=times.min(),
        time_max=times.max(),
        frequency=estimate_frequency(times),
    )
    logger.info("开始清洗：共 %d 行，时间范围 %s ~ %s，频率 %s",
                len(df), stats.time_min, stats.time_max, stats.frequency)

    out = pd.DataFrame({"日期": times})
    result: dict[str, tuple[pd.Series, np.ndarray]] = {}

    for col in PRESSURE_COLUMNS + (GAS_COLUMN,):
        raw = pd.to_numeric(df[col], errors="coerce").reset_index(drop=True)
        out[f"{col}_原始"] = raw

        # 步骤 1：负值 → NaN
        cleaned = raw.astype(float).copy()
        neg_mask = cleaned < 0
        neg_count = int(neg_mask.sum())
        stats.negatives[col] = neg_count
        flags = np.zeros(len(df), dtype=np.int8)
        if neg_count:
            logger.info("%s 列发现 %d 个负值，已替换为 NaN。", col, neg_count)
        cleaned[neg_mask] = np.nan
        flags[neg_mask.to_numpy()] = FLAG_UNFIXED  # 先标记为待修复

        # 步骤 2：油压/套压固定值段检测（瞬时气量零值保留，不做固定值处理）
        if col in PRESSURE_COLUMNS:
            invalid, runs = detect_constant_runs(cleaned, times)
            stats.constant_runs[col] = len(runs)
            if runs:
                logger.info(
                    "%s 列检测到 %d 段固定不变数据（首段：值=%s，点数=%d，持续 %.1f 小时）",
                    col, len(runs), runs[0][2], runs[0][3], runs[0][4],
                )
                invalid_mask = pd.Series(invalid, index=cleaned.index)
                cleaned[invalid_mask] = np.nan
                flags[invalid_mask.to_numpy()] = FLAG_UNFIXED

                if cleaned.notna().any():
                    stats.warnings.append(
                        f"{col}列检测到 {len(runs)} 段连续不变数据（≥{CONSTANT_MIN_POINTS} 个点"
                        f"或 ≥{CONSTANT_MIN_HOURS:.0f} 小时），已标记为无效并用线性插值修复。"
                    )
                else:
                    stats.warnings.append(
                        f"{col}列全部为固定值/0，没有可用的相邻有效值，该列保留空值。"
                    )

        # 步骤 3：线性插值并记录标记
        cleaned, flags, unfixed = interpolate_series(cleaned, times, flags)
        # 瞬时气量自动保留 4 位小数
        if col == GAS_COLUMN:
            cleaned = cleaned.round(4)
        stats.interpolated[col] = int((flags == FLAG_INTERPOLATED).sum())
        stats.unfixed[col] = int((flags == FLAG_UNFIXED).sum())
        result[col] = (cleaned, flags)

        out[col] = cleaned
        out[f"{col}_插值标记"] = flags

    # 瞬时气量零值统计（按原始值统计，保留）
    gas_raw = out["瞬时气量_原始"]
    stats.zero_gas_rows = int(((gas_raw == 0) & gas_raw.notna()).sum())
    if stats.zero_gas_rows == stats.total_rows and stats.total_rows > 0:
        stats.warnings.append("瞬时气量全部为 0，可能处于长期关井状态，图表将显示“关井”标注。")

    # 全部数据为空检查
    if out["油压"].notna().sum() == 0 and out["套压"].notna().sum() == 0 and out["瞬时气量"].notna().sum() == 0:
        raise ValueError("所有数据列均为空，无法生成曲线图。")

    logger.info("清洗完成：插值 %s，无法修复 %s", stats.interpolated, stats.unfixed)
    return out, stats
