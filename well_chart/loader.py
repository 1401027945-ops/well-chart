# -*- coding: utf-8 -*-
"""文件加载模块。

负责读取用户上传的 .xls/.xlsx 文件，并完成三件事：
1. 在文件开头若干行中自动定位表头（兼容“标题行 + 表头行”的常见格式）；
2. 识别日期列、油压列、套压列、瞬时气量列（支持中文/英文列名）；
3. 从“井号：xxx”标注、表头附近单元格、Sheet 名或文件名中提取井号。

该模块不依赖 Streamlit，便于单元测试。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import BinaryIO, Union

import pandas as pd

from .config import get_logger

logger = get_logger()

# 在文件开头多少行范围内寻找表头（一般数据表前几行都是标题/备注）
HEADER_SCAN_ROWS = 8

# 显式井号标注，如 “井号：苏36-13”“井号: 苏36-5-2”
WELL_LABEL_PATTERN = re.compile(r"井号\s*[:：]\s*([^\s,，;；]+)")

# 形如井号的单元格/Sheet 名：中文或字母开头 + 数字 + “-” + 数字（如 苏36-13、苏36-5-2）
WELL_ID_PATTERN = re.compile(r"^[\u4e00-\u9fa5A-Za-z]{0,6}\d{1,5}[-_—－]\d{1,6}井?$")

# 常见列名关键字（统一转小写后匹配）
DATE_KEYWORDS = ("日期", "时间", "date", "time", "timestamp", "datetime")


class LoadError(Exception):
    """文件加载或解析错误。"""


def _norm(value) -> str:
    """把任意单元格内容转为去除空白的小写字符串。"""
    return "" if value is None else str(value).strip().lower()


def _classify_header(text) -> str | None:
    """把表头文本归类为列类型：date / oil / casing / gas / None。

    注意匹配顺序：先判断“套压”，再判断“油压”，避免“套压”被“压”字误判。
    """
    t = _norm(text)
    if not t:
        return None
    if "套压" in t or "casing" in t:
        return "casing"
    if "油压" in t or "tubing" in t or "oil" in t:
        return "oil"
    if "气量" in t or "gas" in t or "瞬时" in t:
        return "gas"
    if any(k in t for k in DATE_KEYWORDS):
        return "date"
    return None


def _find_header_row(raw: pd.DataFrame) -> int:
    """在文件前几行中找出包含最多列关键字的那一行作为表头行。"""
    best_row, best_score = -1, 0
    limit = min(len(raw), HEADER_SCAN_ROWS)
    for i in range(limit):
        score = 0
        for value in raw.iloc[i].astype(str):
            if _classify_header(value):
                score += 1
        if score > best_score:
            best_row, best_score = i, score
    if best_score < 2:
        raise LoadError(
            "未找到包含“日期/油压/套压/瞬时气量”的表头行，请检查文件格式是否符合要求。"
        )
    return best_row


def _locate_columns(raw: pd.DataFrame, header_row: int) -> dict:
    """在表头行及其上方两行中定位各列位置。

    某些文件把“日期”和“井号”放在表头上一行（例如示例 .xls），
    因此需要向上多找两行，但列号以表头行下方的数据为准。
    """
    search_rows = range(max(0, header_row - 2), header_row + 1)
    date_col: int | None = None
    cols: dict[str, int] = {}
    for i in search_rows:
        for j in range(raw.shape[1]):
            ctype = _classify_header(raw.iat[i, j])
            if ctype == "date" and date_col is None:
                date_col = j
            elif ctype in ("oil", "casing", "gas") and ctype not in cols:
                cols[ctype] = j
    if date_col is None:
        raise LoadError("未找到日期/时间列，请确认数据表包含“日期”列。")
    missing = [n for n, key in (("油压", "oil"), ("套压", "casing"), ("瞬时气量", "gas")) if key not in cols]
    if missing:
        raise LoadError("缺少数据列：" + "、".join(missing))
    return {"date": date_col, **cols}


def extract_well_name(raw: pd.DataFrame, header_row: int, sheet_name: str, file_name: str) -> str:
    """提取井号：显式标注 → 表头附近形似井号的单元格 → Sheet 名 → 文件名。"""
    # 1. 显式“井号：xxx”
    meta_rows = min(header_row + 3, len(raw))
    for i in range(meta_rows):
        for j in range(raw.shape[1]):
            cell_value = raw.iat[i, j]
            cell = "" if pd.isna(cell_value) else str(cell_value)
            m = WELL_LABEL_PATTERN.search(cell)
            if m:
                return m.group(1).strip()
    # 2. 表头附近（含表头行）形似井号的单元格，例如“苏36-13”
    for i in range(0, min(header_row + 1, len(raw))):
        for j in range(raw.shape[1]):
            cell_value = raw.iat[i, j]
            if isinstance(cell_value, str) and WELL_ID_PATTERN.match(cell_value.strip()):
                return cell_value.strip().rstrip("井")
    # 3. Sheet 名（排除 Sheet1/Sheet2 等默认名称）
    if sheet_name and not re.match(r"^sheet\d*$", sheet_name, re.IGNORECASE):
        return sheet_name.strip()
    # 4. 文件名（如“苏36-13_历史数据.xls”）
    stem = Path(file_name).stem
    m = WELL_LABEL_PATTERN.search(stem) or WELL_ID_PATTERN.search(stem)
    if m:
        return (m.group(1) if m.lastindex else m.group(0)).strip()
    return "未知井号"


def _parse_dates(values: pd.Series) -> pd.Series:
    """解析日期列：兼容 Excel 序列号与各种字符串日期格式。"""
    numeric = pd.to_numeric(values, errors="coerce")
    dates = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")

    # 数值型日期 → 按 Excel 序列号换算（1899-12-30 为起点）
    num_ok = numeric.notna()
    if num_ok.any():
        excel_epoch = pd.Timestamp("1899-12-30")
        dates[num_ok] = excel_epoch + pd.to_timedelta(numeric[num_ok], unit="D")

    # 字符串日期 → 统一“年月日”写法后交给 pandas 解析
    str_vals = values[num_ok == False].dropna()  # noqa: E712
    if len(str_vals):
        s = pd.Series([str(v) for v in str_vals], index=str_vals.index)
        s = (
            s.str.replace("年", "-", regex=False)
            .str.replace("月", "-", regex=False)
            .str.replace("日", "", regex=False)
        )
        parsed = pd.to_datetime(s, errors="coerce", format="mixed")
        dates[str_vals.index] = parsed
    return dates


def _load_workbook(source: Union[str, Path, BinaryIO]) -> pd.ExcelFile:
    """读取工作簿，支持路径与文件对象（Streamlit 上传的 BytesIO）。"""
    try:
        return pd.ExcelFile(source)
    except Exception as exc:  # noqa: BLE001
        raise LoadError(f"无法读取 Excel 文件：{exc}") from exc


def load_well_data(source: Union[str, Path, BinaryIO]) -> dict:
    """加载并标准化单井历史数据。

    返回字典：
    - data: 标准化后的 DataFrame，列为 日期/油压/套压/瞬时气量
    - raw:  原始读入的 DataFrame（调试用）
    - well_name: 识别出的井号
    - sheet_name / file_name: 来源信息

    日期无法解析的行会被丢弃并记录日志。
    """
    if isinstance(source, (str, Path)):
        file_name = Path(source).name
    else:
        file_name = getattr(source, "name", "upload.xlsx") or "upload.xlsx"

    xls = _load_workbook(source)
    logger.info("读取工作簿成功，共 %d 个 Sheet：%s", len(xls.sheet_names), xls.sheet_names)

    last_error: Exception | None = None
    for sheet_name in xls.sheet_names:
        try:
            raw = xls.parse(sheet_name=sheet_name, header=None)
            if raw.empty:
                continue
            header_row = _find_header_row(raw)
            cols = _locate_columns(raw, header_row)
            well_name = extract_well_name(raw, header_row, sheet_name, file_name)
            logger.info(
                "Sheet[%s] 识别成功：表头行=%d，列位置=%s，井号=%s",
                sheet_name, header_row, cols, well_name,
            )

            # 取表头下一行开始的所有数据
            data_start = header_row + 1
            data = raw.iloc[data_start:].copy()
            df = pd.DataFrame(
                {
                    "日期": _parse_dates(data.iloc[:, cols["date"]]),
                    "油压": data.iloc[:, cols["oil"]],
                    "套压": data.iloc[:, cols["casing"]],
                    "瞬时气量": data.iloc[:, cols["gas"]],
                }
            )
            # 数值列转为数值类型（无法转换的自动变为 NaN）
            for col in ("油压", "套压", "瞬时气量"):
                df[col] = pd.to_numeric(df[col], errors="coerce")

            # 丢弃日期无法解析的行
            bad_dates = df["日期"].isna()
            if bad_dates.any():
                logger.warning("有 %d 行日期无法解析，已删除。", int(bad_dates.sum()))
                df = df[~bad_dates].reset_index(drop=True)

            if df.empty:
                raise LoadError("解析后没有有效数据（日期列均为空或无法解析）。")

            df = df.sort_values("日期").reset_index(drop=True)
            return {
                "data": df,
                "raw": raw,
                "well_name": well_name,
                "sheet_name": sheet_name,
                "file_name": file_name,
                "header_row": header_row,
            }
        except LoadError as exc:
            last_error = exc
            logger.warning("Sheet[%s] 解析失败：%s", sheet_name, exc)
            continue
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            logger.warning("Sheet[%s] 读取异常：%s", sheet_name, exc)
            continue

    raise LoadError(str(last_error or "所有 Sheet 均无法解析。"))
