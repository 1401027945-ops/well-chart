# -*- coding: utf-8 -*-
"""把下载结果中的原生 Excel 图表实时渲染为图片（本地预览用）。

优先使用本机 WPS（KET.Application），其次 Microsoft Excel；
两者都不可用（例如部署在 Linux 服务器）时返回 None，界面回退到 matplotlib 预览。
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

from .config import get_logger

logger = get_logger()


def render_native_chart(xlsx_bytes: bytes, sheet_index: int = 3, chart_index: int = 1) -> tuple[bytes | None, str]:
    """渲染 Excel 工作簿中指定子表的原生图表为 PNG 字节。

    两个模板的图表都位于第 3 个子表（图片 / 曲线图）。
    返回 (PNG 字节或 None, 说明文字)。当前环境无法渲染时说明文字记录失败原因。
    """
    try:
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        return None, f"缺少本地渲染组件（win32com）：{exc}"

    tmp_dir = tempfile.mkdtemp(prefix="well_chart_render_")
    xlsx_path = os.path.join(tmp_dir, "chart_src.xlsx")
    png_path = os.path.join(tmp_dir, "chart_out.png")
    try:
        Path(xlsx_path).write_bytes(xlsx_bytes)
        # 依次尝试 WPS 表格 / Microsoft Excel，每个最多重试 2 次
        last_error = "未找到可用的本地办公软件"
        for progid in ("KET.Application", "Excel.Application"):
            for attempt in (1, 2):
                app = None
                try:
                    app = win32com.client.Dispatch(progid)
                    app.Visible = False
                    try:
                        app.DisplayAlerts = False
                    except Exception:  # noqa: BLE001
                        pass
                    wb = app.Workbooks.Open(xlsx_path, 0, True)  # 只读打开
                    ws = wb.Sheets.Item(sheet_index)
                    chart = ws.ChartObjects().Item(chart_index).Chart
                    chart.Export(png_path)
                    wb.Close(False)
                    data = Path(png_path).read_bytes()
                    if data:
                        logger.info("已用 %s 渲染图表预览（%d 字节）", progid, len(data))
                        return data, ""
                except Exception as exc:  # noqa: BLE001
                    last_error = f"{progid}：{exc}"
                    logger.warning("%s 渲染图表预览失败（第 %d 次）：%s", progid, attempt, exc)
                finally:
                    if app is not None:
                        try:
                            app.Quit()
                        except Exception:  # noqa: BLE001
                            pass
        return None, last_error
    finally:
        # 清理临时文件
        for path in (xlsx_path, png_path):
            try:
                os.remove(path)
            except Exception:  # noqa: BLE001
                pass
        try:
            os.rmdir(tmp_dir)
        except Exception:  # noqa: BLE001
            pass
