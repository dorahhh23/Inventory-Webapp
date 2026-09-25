from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from forecast_analyzer import AnalysisSettings, analyze, build_views, read_source_excel, to_excel_bytes


@st.cache_data(show_spinner="正在读取并分析工作簿…")
def run_analysis(file_bytes: bytes, stock_method: str):
    source = read_source_excel(io.BytesIO(file_bytes))
    return analyze(source, AnalysisSettings(stock_method=stock_method))


@st.cache_data(show_spinner="正在生成 Excel…")
def build_export(result):
    return to_excel_bytes(result)


def highlight_return_risk(row: pd.Series) -> list[str]:
    if row.get("是否会引起断货问题") == "是":
        return ["background-color: #f4cccc; color: #9c0006"] * len(row)
    return [""] * len(row)


st.set_page_config(page_title="补货预测异常识别", page_icon="📦", layout="wide")

st.title("补货预测异常识别")
st.caption("上传补货预测 Excel，识别降级误补货、退货异常和疑似偶发性大单。")

upload_column, settings_column = st.columns([2, 1], gap="large")

with upload_column:
    with st.container(border=True):
        st.markdown("#### 上传文件")
        uploaded = st.file_uploader("上传 .xlsx 文件", type=["xlsx"], label_visibility="collapsed")
        upload_status = st.empty()
        download_area = st.empty()

with settings_column:
    with st.container(border=True):
        st.markdown("#### 计算设置")
        stock_label = st.selectbox(
            "有效库存口径",
            ["非限制库存 + 在途 + 在产", "当前可用库存 + 在途 + 在产", "仅非限制库存"],
            help="默认口径包含在库、在途和在产，不会把退货扭曲量重复加回库存。",
        )
        stock_method = {
            "非限制库存 + 在途 + 在产": "on_hand_plus_pipeline",
            "当前可用库存 + 在途 + 在产": "available_plus_pipeline",
            "仅非限制库存": "on_hand_only",
        }[stock_label]

if uploaded is None:
    upload_status.info("请上传工作簿。系统会自动读取两行表头，并排除13个月销量中的最后一个当前月。")
    st.stop()

try:
    result = run_analysis(uploaded.getvalue(), stock_method)
    views = build_views(result)
except Exception as exc:
    st.error(f"分析失败：{exc}")
    st.stop()

summary = result["summary"]
upload_status.success(
    f"已读取 {summary['records']:,} 个物料。识别月份："
    f"{' → '.join(summary['monthly_columns'])}；已排除当前月 {summary['monthly_columns'][-1]}。"
)

excel = build_export(result)
download_area.download_button(
    "下载分析结果 (.xlsx)",
    data=excel,
    file_name="补货预测异常分析结果.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# summary metrics
metrics = [
    ("汇总记录", len(views["异常清单汇总"])),
    ("降级误补货", summary["issue1"]),
    ("退货异常", summary["negative_last_month"]),
    ("偶发大单", summary["oneoff"]),
]
metric_columns = st.columns(4)
for column, (label, value) in zip(metric_columns, metrics):
    column.metric(label, f"{value:,}")

tab_summary, tab_details = st.tabs(["异常清单汇总", "细分清单"])

with tab_summary:
    st.dataframe(views["异常清单汇总"], use_container_width=True, hide_index=True)

with tab_details:
    tab_downgrade, tab_return, tab_oneoff = st.tabs(
        ["降级误补货问题", "退货异常问题", "偶发大单分析"]
    )

    with tab_downgrade:
        st.write("该问题适用于上一季度为A/B级但本季度为C级的产品。降级后计算预测月销量的历史数据范围由6个月增加为12个月，可能导致预测月销量过高，从而误补货。")
        st.dataframe(views["降级误补货问题"], use_container_width=True, hide_index=True)

    with tab_return:
        st.write("检查上个月是否发生大额退货/冲销导致预测销量远低于实际月销，从而引起缺货问题。")
        st.caption("红色强调行会引起断货问题，并优先显示；其余退货异常记录显示在下方。判定不考虑海运时间。")
        styled = views["退货异常问题"].style.apply(highlight_return_risk, axis=1).format(precision=2, na_rep="")
        st.dataframe(styled, use_container_width=True, hide_index=True)

    with tab_oneoff:
        st.write("判定最近一个月销量是否为偶发性大单。")
        st.dataframe(views["偶发大单分析"], use_container_width=True, hide_index=True)
