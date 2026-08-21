# -*- coding: utf-8 -*-
"""单元测试：数据清洗、时间刻度、纵轴范围、文件解析。

可直接运行：python tests/test_all.py（无需 pytest，输出简单通过/失败信息）。
"""

from __future__ import annotations

import io
import sys
import traceback
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from well_chart import cleaning, loader, plotting  # noqa: E402

PASSED = 0
FAILED = 0


def check(name: str, condition: bool, detail: str = "") -> None:
    global PASSED, FAILED
    if condition:
        PASSED += 1
        print(f"  ✓ {name}")
    else:
        FAILED += 1
        print(f"  ✗ {name}  {detail}")


def make_df(
    times: list,
    oil: list,
    casing: list,
    gas: list,
) -> pd.DataFrame:
    return pd.DataFrame(
        {"日期": pd.to_datetime(times), "油压": oil, "套压": casing, "瞬时气量": gas}
    )


def test_negative_replacement() -> None:
    print("[1] 负值替换")
    df = make_df(
        ["2026-01-01 00:00", "2026-01-01 01:00", "2026-01-01 02:00", "2026-01-01 03:00"],
        [1.0, -2.0, 3.0, 4.0],
        [2.0, 3.0, -1.0, 5.0],
        [0.0, -5.0, 2.0, 3.0],
    )
    out, stats = cleaning.clean_well_data(df)
    check("油压负值被插值修复", abs(out.loc[1, "油压"] - 2.0) < 1e-9)
    check("套压负值被插值修复", abs(out.loc[2, "套压"] - 4.0) < 1e-9)
    check("瞬时气量负值被插值修复", abs(out.loc[1, "瞬时气量"] - 1.0) < 1e-9)
    check("原始列保留负值", out.loc[1, "油压_原始"] == -2.0)
    check("负值点标记为插值", out.loc[1, "油压_插值标记"] == cleaning.FLAG_INTERPOLATED)
    check("负值统计正确", stats.negatives == {"油压": 1, "套压": 1, "瞬时气量": 1})
    check("瞬时气量零值保留", out.loc[0, "瞬时气量"] == 0.0)


def test_interpolation_interior() -> None:
    print("[2] 内部插值")
    df = make_df(
        ["2026-01-01 00:00", "2026-01-01 01:00", "2026-01-01 02:00", "2026-01-01 03:00"],
        [1.0, np.nan, 3.0, 4.0],
        [1.0, 2.0, 3.0, 4.0],
        [0.0, 0.0, 0.0, 0.0],
    )
    out, stats = cleaning.clean_well_data(df)
    check("油压内部 NaN 被插值", abs(out.loc[1, "油压"] - 2.0) < 1e-9)
    check("插值标记为 1", out.loc[1, "油压_插值标记"] == cleaning.FLAG_INTERPOLATED)
    check("原始点标记为 0", out.loc[0, "油压_插值标记"] == cleaning.FLAG_ORIGINAL)


def test_edge_unfixable() -> None:
    print("[3] 首尾无法插值")
    df = make_df(
        ["2026-01-01 00:00", "2026-01-01 01:00", "2026-01-01 02:00", "2026-01-01 03:00"],
        [np.nan, 2.0, 3.0, np.nan],
        [1.0, 2.0, 3.0, 4.0],
        [0.0, 0.0, 0.0, 0.0],
    )
    out, stats = cleaning.clean_well_data(df)
    check("首点保持空值", pd.isna(out.loc[0, "油压"]))
    check("尾点保持空值", pd.isna(out.loc[3, "油压"]))
    check("无法修复标记为 2", out.loc[0, "油压_插值标记"] == cleaning.FLAG_UNFIXED)


def test_constant_run_detection() -> None:
    print("[4] 固定值段检测")
    # 两段各 6 个点固定为 5.3（>=6 点），前后各有有效锚点可插值
    times = pd.date_range("2026-01-01", periods=15, freq="h")
    oil = [1.0] + [5.3] * 6 + [6.0] + [5.3] * 6 + [9.0]
    casing = [round(2.0 + i * 0.1, 2) for i in range(15)]
    df = make_df(times, oil, casing, [0.0] * 15)
    out, stats = cleaning.clean_well_data(df)
    check("检测到 2 段固定值", stats.constant_runs["油压"] == 2)
    check("固定值段被插值", out.loc[1, "油压_插值标记"] == cleaning.FLAG_INTERPOLATED)
    check("固定值段插值结果非 5.3", abs(out.loc[1, "油压"] - 5.3) > 1e-9)
    check("固定值段插值有梯度", abs(out.loc[1, "油压"] - out.loc[6, "油压"]) > 1e-9)


def test_constant_run_threshold() -> None:
    print("[5] 固定值阈值（<6 点不处理）")
    times = pd.date_range("2026-01-01", periods=10, freq="h")
    oil = [5.3] * 3 + [6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0]
    casing = [round(1.0 + i * 0.1, 2) for i in range(10)]
    df = make_df(times, oil, casing, [0.0] * 10)
    out, stats = cleaning.clean_well_data(df)
    check("短固定段不处理", stats.constant_runs["油压"] == 0)
    check("原始值保留", out.loc[0, "油压"] == 5.3)


def test_all_zero_gas() -> None:
    print("[6] 瞬时气量全为 0")
    df = make_df(
        pd.date_range("2026-01-01", periods=12, freq="h"),
        [1.0] * 12, [2.0] * 12, [0.0] * 12,
    )
    out, stats = cleaning.clean_well_data(df)
    check("零值全部保留", out["瞬时气量"].sum() == 0.0)
    check("关井警告生成", any("关井" in w for w in stats.warnings))


def test_all_constant_pressure_column() -> None:
    print("[7] 整列固定值")
    df = make_df(
        pd.date_range("2026-01-01", periods=12, freq="h"),
        [5.3] * 12, [2.0] * 12, [0.0] * 12,
    )
    out, stats = cleaning.clean_well_data(df)
    check("整列固定值全部为空", out["油压"].isna().all())
    check("整列固定值警告", any("全部为固定值" in w for w in stats.warnings))


def test_frequency_estimation() -> None:
    print("[8] 采样频率识别")
    hourly = pd.date_range("2026-01-01", periods=10, freq="h")
    daily = pd.date_range("2026-01-01", periods=10, freq="D")
    minute = pd.date_range("2026-01-01", periods=10, freq="10min")
    check("小时级", cleaning.estimate_frequency(hourly) == "小时级")
    check("天级", cleaning.estimate_frequency(daily) == "天级")
    check("分钟级", cleaning.estimate_frequency(minute) == "分钟级")


def test_nice_ceil() -> None:
    print("[9] 纵轴取整")
    check("小于 1 取 1", plotting.nice_ceil(0.5) == 1.0)
    check("7.5 取 8", plotting.nice_ceil(7.5) == 8.0)
    check("5.83 取 6", plotting.nice_ceil(5.83) == 6.0)
    check("100 取 100", plotting.nice_ceil(100) == 100.0)
    check("12 取 20", plotting.nice_ceil(12) == 20.0)


def test_y_range() -> None:
    print("[10] 纵轴范围（+10% 余量）")
    lo, hi = plotting.compute_y_range(pd.Series([1.0, 5.0, 5.3]))
    check("最高值含 10% 余量", hi >= 5.3 * 1.1 and hi == 6.0)
    check("最低为 0", lo == 0.0)
    lo2, hi2 = plotting.compute_y_range(pd.Series([0.1]))
    check("最大值<1 取 1", hi2 == 1.0)


def test_time_ticks_equal_spacing() -> None:
    print("[11] 等间隔时间刻度")
    tmin = pd.Timestamp("2026-01-01 18:00:00")
    tmax = pd.Timestamp("2026-08-21 18:00:00")
    ticks = plotting.build_time_ticks(tmin, tmax)
    check("刻度数量为 6", len(ticks) == 6)
    check("首尾为数据边界", ticks[0] == tmin and ticks[-1] == tmax)
    diffs = [(ticks[i + 1] - ticks[i]).total_seconds() for i in range(len(ticks) - 1)]
    check("相邻间隔完全相等", max(diffs) - min(diffs) < 1e-6, diffs)


def test_chart_style() -> None:
    print("[12] 图表样式（无标题/无日期标签/无网格/图例居中单行）")
    df = make_df(
        pd.date_range("2026-01-01", periods=24, freq="h"),
        [3 + i * 0.1 for i in range(24)],
        [4 + i * 0.1 for i in range(24)],
        [i * 0.5 for i in range(24)],
    )
    out, stats = cleaning.clean_well_data(df)
    fig = plotting.create_chart(out, "测试井-1", stats)
    ax, ax2 = fig.axes
    check("无图标题", ax.get_title() == "")
    check("无横轴“日期”标签", ax.get_xlabel() == "")
    check(
        "无网格线",
        not any(gl.get_visible() for gl in ax.get_xgridlines() + ax.get_ygridlines()),
    )
    check("图例居中（upper center）", ax.get_legend()._loc == 9)
    check("图例只占一行（3 列）", ax.get_legend()._ncols == 3)
    check("隐藏最上面的横线", not ax.spines["top"].get_visible() and not ax2.spines["top"].get_visible())
    check("坐标轴标签字号 9", ax.yaxis.get_label().get_fontsize() == 9)
    check("刻度文字字号 8", ax.xaxis.get_ticklabels()[0].get_fontsize() == 8)
    check("图例字号 11", ax.get_legend().get_texts()[0].get_fontsize() == 11)
    check(
        "坐标轴线黑色 0.5 磅",
        tuple(ax.spines["left"].get_edgecolor()) == (0.0, 0.0, 0.0, 1.0)
        and ax.spines["left"].get_linewidth() == 0.5,
    )
    check(
        "右轴坐标线黑色 0.5 磅",
        tuple(ax2.spines["right"].get_edgecolor()) == (0.0, 0.0, 0.0, 1.0)
        and ax2.spines["right"].get_linewidth() == 0.5,
    )
    plotting.close_fig(fig)


def test_loader_sample_xls() -> None:
    print("[13] 示例 .xls 解析")
    sample = Path(__file__).resolve().parents[1] / "sample_data" / "单井历史数据-2026-08-21.xls"
    if not sample.exists():
        check("示例文件存在", False, f"缺少 {sample}")
        return
    result = loader.load_well_data(sample)
    df = result["data"]
    check("井号识别为苏36-13", result["well_name"] == "苏36-13")
    check("列为日期/油压/套压/瞬时气量", list(df.columns) == ["日期", "油压", "套压", "瞬时气量"])
    check("日期为 datetime64", pd.api.types.is_datetime64_any_dtype(df["日期"]))
    check("数据行数大于 1000", len(df) > 1000)
    check("数据按时间排序", df["日期"].is_monotonic_increasing)


def test_loader_missing_date_column() -> None:
    print("[14] 缺少时间列报错")
    buf = io.BytesIO()
    pd.DataFrame({"油压": [1, 2], "套压": [3, 4], "瞬时气量": [0, 0]}).to_excel(buf, index=False)
    buf.seek(0)
    try:
        loader.load_well_data(buf)
        check("应抛出 LoadError", False)
    except loader.LoadError:
        check("应抛出 LoadError", True)


def test_loader_bytesio() -> None:
    print("[15] 内存文件解析")
    buf = io.BytesIO()
    pd.DataFrame(
        {
            "日期": pd.date_range("2026-01-01", periods=20, freq="h"),
            "油压": np.arange(20) * 0.1,
            "套压": np.arange(20) * 0.2,
            "瞬时气量": [0.0] * 20,
        }
    ).to_excel(buf, index=False)
    buf.seek(0)
    result = loader.load_well_data(buf)
    check("BytesIO 解析成功", len(result["data"]) == 20)
    check("默认井号", result["well_name"] == "未知井号")


def main() -> None:
    tests = [
        test_negative_replacement,
        test_interpolation_interior,
        test_edge_unfixable,
        test_constant_run_detection,
        test_constant_run_threshold,
        test_all_zero_gas,
        test_all_constant_pressure_column,
        test_frequency_estimation,
        test_nice_ceil,
        test_y_range,
        test_time_ticks_equal_spacing,
        test_chart_style,
        test_loader_sample_xls,
        test_loader_missing_date_column,
        test_loader_bytesio,
    ]
    for fn in tests:
        try:
            fn()
        except Exception:  # noqa: BLE001
            global FAILED
            FAILED += 1
            print(f"  ✗ {fn.__name__} 异常")
            traceback.print_exc()
    print(f"\n结果：{PASSED} 项通过，{FAILED} 项失败")
    sys.exit(1 if FAILED else 0)


if __name__ == "__main__":
    main()
