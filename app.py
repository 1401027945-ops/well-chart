# -*- coding: utf-8 -*-
"""生产曲线模板门户 - Streamlit 主程序。

当前提供“单井生产曲线模板”，并预留“日期叠合曲线模板”“天数叠合曲线模板”
两个后续模板入口。任何人访问网站后上传数据即可自动生成曲线并下载 Excel。

启动方式：streamlit run app.py
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from well_chart import cleaning, excel_export, loader, plotting
from well_chart.config import MIN_ROWS, get_logger

logger = get_logger()

st.set_page_config(
    page_title="生产曲线模板",
    page_icon="📈",
    layout="wide",
)

# ---------------------------------------------------------------------------
# 模板定义：当前可用 1 个，后续开发 2 个
# ---------------------------------------------------------------------------
TEMPLATES = {
    "单井生产曲线模板": {
        "icon": "🛢️",
        "status": "ready",
        "badge": "当前可用",
        "desc": "上传单井历史数据（日期、油压、套压、瞬时气量），自动清洗异常值、"
                "修复固定值段，生成双纵轴曲线图，并下载包含数据与图表的 Excel。",
    },
    "日期叠合曲线模板": {
        "icon": "📅",
        "status": "ready",
        "badge": "当前可用",
        "desc": "基于单井生产曲线模板，瞬时气量以面积图展示，其余功能与单井模板一致。",
    },
    "天数叠合曲线模板": {
        "icon": "📆",
        "status": "coming",
        "badge": "开发中",
        "desc": "多口井按生产天数叠合对比展示，即将上线。",
    },
}


def inject_css() -> None:
    """注入自定义样式，让页面美观简洁。"""
    st.markdown(
        """
        <style>
        .stApp {
            background: #F7F9FC;
            font-family: "Microsoft YaHei", "SimHei", sans-serif;
        }
        #MainMenu, footer { visibility: hidden; }
        .block-container { padding-top: 2.2rem; padding-bottom: 3rem; }

        .site-hero {
            background: linear-gradient(135deg, #0F6CBD 0%, #3E9BDF 100%);
            color: #fff;
            border-radius: 14px;
            padding: 26px 32px;
            margin-bottom: 18px;
        }
        .site-hero h1 { margin: 0; font-size: 26px; font-weight: 700; }
        .site-hero p { margin: 6px 0 0; opacity: .92; font-size: 14px; }

        .tmpl-grid { display: flex; gap: 16px; flex-wrap: wrap; margin-top: 6px; }
        .tmpl-card {
            flex: 1 1 280px;
            background: #fff;
            border: 1px solid #E4E9F2;
            border-radius: 12px;
            padding: 18px 20px;
            box-shadow: 0 2px 8px rgba(15, 108, 189, .06);
        }
        .tmpl-card .tmpl-icon { font-size: 30px; }
        .tmpl-card .tmpl-name { font-size: 17px; font-weight: 700; margin: 6px 0 4px; color: #1F2328; }
        .tmpl-card .tmpl-desc { font-size: 13px; color: #5A6472; line-height: 1.6; }
        .badge {
            display: inline-block; padding: 2px 10px; border-radius: 20px;
            font-size: 12px; font-weight: 600; margin-top: 6px;
        }
        .badge-ready { background: #E6F6EC; color: #1B9C55; }
        .badge-coming { background: #FFF3E0; color: #D97A06; }

        .coming-box {
            background: #FFFBF0; border: 1px dashed #E5B667;
            border-radius: 12px; padding: 34px 26px; text-align: center;
            color: #8A6D3B; margin-top: 8px;
        }
        .coming-box .big { font-size: 40px; }

        .section-title { font-size: 16px; font-weight: 700; color: #1F2328; margin: 18px 0 8px; }

        [data-testid="stFileUploader"] section {
            background: #fff; border: 1px dashed #B9CFE8; border-radius: 12px;
        }
        [data-testid="stSidebar"] {
            background: #FFFFFF; border-right: 1px solid #EDF0F5;
        }
        [data-testid="stMetric"] {
            background: #fff; border: 1px solid #E4E9F2; border-radius: 10px;
            padding: 12px 16px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_template_cards() -> None:
    """首页展示三个模板卡片（当前可用 / 开发中）。"""
    cards = []
    for name, info in TEMPLATES.items():
        badge_cls = "badge-ready" if info["status"] == "ready" else "badge-coming"
        cards.append(
            f'<div class="tmpl-card">'
            f'<div class="tmpl-icon">{info["icon"]}</div>'
            f'<div class="tmpl-name">{name}</div>'
            f'<div class="tmpl-desc">{info["desc"]}</div>'
            f'<span class="badge {badge_cls}">{info["badge"]}</span>'
            f"</div>"
        )
    st.markdown('<div class="tmpl-grid">' + "".join(cards) + "</div>", unsafe_allow_html=True)


def render_coming_soon(name: str) -> None:
    """展示开发中模板的占位页面。"""
    info = TEMPLATES[name]
    st.markdown(
        f'<div class="coming-box"><div class="big">{info["icon"]}</div>'
        f"<h2>{name}</h2>"
        f"<p>{info['desc']}</p>"
        f"<p>该模板正在开发中，上线后此处将自动开放上传与生成功能。</p></div>",
        unsafe_allow_html=True,
    )


def run_single_well_flow(df: pd.DataFrame, well_name: str, gas_area: bool = False) -> None:
    """单井类模板流程：数据预览 → 清洗 → 曲线图 → 下载。

    gas_area=True 表示日期叠合曲线模板（瞬时气量为面积图）。
    """
    st.success(f"文件解析成功：井号 {well_name}，共 {len(df)} 条数据")

    st.markdown('<div class="section-title">① 数据预览（前 10 行）</div>', unsafe_allow_html=True)
    st.dataframe(df.head(10), width="stretch")

    with st.spinner("正在清洗数据（负值处理、固定值识别、插值修复）..."):
        df_clean, stats = cleaning.clean_well_data(df)

    st.markdown('<div class="section-title">② 数据统计</div>', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("数据量", f"{stats.total_rows} 条")
    col2.metric("时间范围", f"{stats.time_min:%Y-%m-%d %H:%M} ~ {stats.time_max:%m-%d %H:%M}")
    col3.metric("采样频率", stats.frequency)
    col4.metric("异常值数量", f"{stats.total_anomalies} 个")

    with st.expander("查看清洗明细"):
        detail = pd.DataFrame(
            {
                "列": ["油压", "套压", "瞬时气量"],
                "负值数": [
                    stats.negatives.get("油压", 0),
                    stats.negatives.get("套压", 0),
                    stats.negatives.get("瞬时气量", 0),
                ],
                "固定值段数": [
                    stats.constant_runs.get("油压", 0),
                    stats.constant_runs.get("套压", 0),
                    "—",
                ],
                "插值修复点数": [
                    stats.interpolated.get("油压", 0),
                    stats.interpolated.get("套压", 0),
                    stats.interpolated.get("瞬时气量", 0),
                ],
                "无法修复点数": [
                    stats.unfixed.get("油压", 0),
                    stats.unfixed.get("套压", 0),
                    stats.unfixed.get("瞬时气量", 0),
                ],
            }
        )
        st.dataframe(detail, width="stretch")
        st.caption("瞬时气量零值属于真实关井，予以保留。")

    for warning in stats.warnings:
        st.warning(warning)

    if df_clean.shape[0] < MIN_ROWS:
        st.error(
            f"数据量不足（仅 {df_clean.shape[0]} 条，少于 {MIN_ROWS} 条），"
            "无法生成有效图表。请上传包含更多数据点的文件。"
        )
        st.stop()

    st.markdown('<div class="section-title">③ 曲线图预览</div>', unsafe_allow_html=True)
    fig = plotting.create_chart(df_clean, well_name, stats, gas_area=gas_area)
    st.pyplot(fig)
    plotting.close_fig(fig)

    st.markdown('<div class="section-title">④ 下载结果</div>', unsafe_allow_html=True)
    xlsx_bytes = excel_export.export_excel(df, df_clean, stats, well_name, gas_area=gas_area)
    st.download_button(
        label="下载 Excel 文件（含原始数据 / 处理后的数据 / 图片 三个子表）",
        data=xlsx_bytes,
        file_name=f"{well_name}_单井生产曲线.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("「图片」子表内为原生 Excel 图表，双击即可编辑；修改处理后的数据，图表会自动更新。")


def main() -> None:
    """页面主流程：模板选择 → 对应模板内容。"""
    inject_css()

    st.markdown(
        """
        <div class="site-hero">
          <h1>📈 生产曲线模板</h1>
          <p>上传数据 → 自动生成曲线 → 下载结果。当前提供单井生产曲线模板，更多模板陆续上线。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.markdown("#### 🧭 模板选择")
        selected = st.radio(
            "选择作图模板",
            list(TEMPLATES.keys()),
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.caption(
            "版本 1.0\n\n"
            "· 单井生产曲线模板：已上线\n"
            "· 日期叠合曲线模板：已上线（瞬时气量面积图）\n"
            "· 天数叠合曲线模板：开发中"
        )

    info = TEMPLATES[selected]
    if info["status"] == "coming":
        st.title(f"{info['icon']} {selected}")
        render_coming_soon(selected)
        st.stop()

    st.title(f"{info['icon']} {selected}")
    st.caption(info["desc"])

    uploaded = st.file_uploader(
        "上传单井历史数据 Excel 文件（支持 .xls / .xlsx，需包含：日期、油压、套压、瞬时气量）",
        type=["xls", "xlsx"],
    )

    if uploaded is None:
        st.markdown('<div class="section-title">模板总览</div>', unsafe_allow_html=True)
        render_template_cards()
        with st.expander("使用说明"):
            st.markdown(
                "1. 点击上方上传区域，选择本机的 Excel 数据文件；\n\n"
                "2. 系统自动识别表头与井号，完成数据清洗并生成曲线图；\n\n"
                "3. 点击“下载 Excel 文件”，得到包含三个子表的成品（原始数据、处理后的数据、图片）。"
            )
        st.stop()

    try:
        with st.spinner("正在解析文件..."):
            loaded = loader.load_well_data(uploaded)
            df = loaded["data"]
            well_name = loaded["well_name"]
        gas_area = selected == "日期叠合曲线模板"
        run_single_well_flow(df, well_name, gas_area=gas_area)
    except loader.LoadError as exc:
        st.error(f"文件格式错误：{exc}")
        logger.error("文件格式错误：%s", exc)
    except ValueError as exc:
        st.error(f"数据无法处理：{exc}")
        logger.error("数据无法处理：%s", exc)
    except Exception as exc:  # noqa: BLE001
        st.error(f"处理过程中发生未预期错误：{exc}")
        logger.exception("处理过程中发生未预期错误")


main()
