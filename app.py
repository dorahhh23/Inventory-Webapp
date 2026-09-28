from __future__ import annotations

import io
import hashlib

import pandas as pd
import streamlit as st

from forecast_analyzer import AnalysisSettings, analyze, build_views, read_source_excel, to_excel_bytes
from inventory_structure_analysis import SUMMARY_AMOUNT_COLUMN, InventoryStructureAnalyzer

INVENTORY_ANALYSIS_CACHE_VERSION = "inventory-price-confirmation-v3"


@st.cache_data(show_spinner="正在读取并分析工作簿…")
def run_analysis(file_bytes: bytes, stock_method: str):
    source = read_source_excel(io.BytesIO(file_bytes))
    return analyze(source, AnalysisSettings(stock_method=stock_method))


@st.cache_data(show_spinner="正在生成 Excel…")
def build_export(result):
    return to_excel_bytes(result)


@st.cache_data(show_spinner="正在读取库存结构文件…")
def read_inventory_files(replenishment_bytes: bytes, price_bytes: bytes):
    analyzer = InventoryStructureAnalyzer()
    return (
        analyzer.read_replenishment_excel(io.BytesIO(replenishment_bytes)),
        analyzer.read_price_excel(io.BytesIO(price_bytes)),
    )


@st.cache_data(show_spinner="正在计算库存结构…")
def run_inventory_analysis(
    replenishment_bytes: bytes,
    price_bytes: bytes,
    material_column: object,
    price_column: object,
    cache_version: str,
):
    analyzer = InventoryStructureAnalyzer()
    replenishment, prices = read_inventory_files(replenishment_bytes, price_bytes)
    return analyzer.analyze(replenishment, prices, material_column, price_column)


def highlight_return_risk(row: pd.Series) -> list[str]:
    if row.get("是否会引起断货问题") == "是":
        return ["background-color: #f4cccc; color: #9c0006"] * len(row)
    return [""] * len(row)


@st.dialog("确认自动识别的价格列")
def confirm_detected_price_column(column: object, decision_key: str) -> None:
    st.write(f"系统已将“{column}”列识别为价格，请核查是否正确。")
    confirm_column, reject_column = st.columns(2)
    with confirm_column:
        if st.button("正确，使用该列", type="primary", use_container_width=True):
            st.session_state[decision_key] = "confirmed"
            st.rerun()
    with reject_column:
        if st.button("不正确，手动选择", use_container_width=True):
            st.session_state[decision_key] = "manual"
            st.rerun()


def render_anomaly_page() -> None:
    st.title("补货预测异常识别")
    st.caption("上传补货预测 Excel，识别降级误补货、退货异常和疑似偶发性大单。")
    upload_column, settings_column = st.columns([2, 1], gap="large")
    with upload_column:
        with st.container(border=True):
            st.markdown("#### 上传文件")
            uploaded = st.file_uploader("上传 .xlsx 文件", type=["xlsx"], label_visibility="collapsed", key="anomaly_file")
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
        return
    try:
        result = run_analysis(uploaded.getvalue(), stock_method)
        views = build_views(result)
    except Exception as exc:
        st.error(f"分析失败：{exc}")
        return
    summary = result["summary"]
    upload_status.success(
        f"已读取 {summary['records']:,} 个物料。识别月份："
        f"{' → '.join(summary['monthly_columns'])}；已排除当前月 {summary['monthly_columns'][-1]}。"
    )
    download_area.download_button(
        "下载分析结果 (.xlsx)", data=build_export(result), file_name="补货预测异常分析结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
    )
    metrics = [
        ("汇总记录", len(views["异常清单汇总"])), ("降级误补货", summary["issue1"]),
        ("退货异常", summary["negative_last_month"]), ("偶发大单", summary["oneoff"]),
    ]
    for column, (label, value) in zip(st.columns(4), metrics):
        column.metric(label, f"{value:,}")
    tab_summary, tab_details = st.tabs(["异常清单汇总", "细分清单"])
    with tab_summary:
        st.dataframe(views["异常清单汇总"], use_container_width=True, hide_index=True)
    with tab_details:
        tab_downgrade, tab_return, tab_oneoff = st.tabs(["降级误补货问题", "退货异常问题", "偶发大单分析"])
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


def show_inventory_summary(title: str, frame: pd.DataFrame) -> None:
    frame = frame.copy()
    if SUMMARY_AMOUNT_COLUMN not in frame.columns and "可用库存金额" in frame.columns:
        frame["可用库存金额"] = frame["可用库存金额"] / 1_000_000
        frame = frame.rename(columns={"可用库存金额": SUMMARY_AMOUNT_COLUMN})
    st.markdown(f"### {title}")
    table_column, chart_column = st.columns([1, 1], gap="large")
    with table_column:
        st.dataframe(frame.style.format({SUMMARY_AMOUNT_COLUMN: "{:,.2f}", "型号数量": "{:,.0f}"}, na_rep=""), use_container_width=True, hide_index=True)
    with chart_column:
        chart_data = frame.iloc[:-1].copy()
        label_columns = [column for column in frame.columns if column not in {SUMMARY_AMOUNT_COLUMN, "型号数量"}]
        st.bar_chart(chart_data.set_index(label_columns[-1])[[SUMMARY_AMOUNT_COLUMN]])


def render_inventory_page() -> None:
    st.title("库存结构分析")
    st.caption("按可用库存覆盖月数识别长期库存、超目标库存及正常/关注库存，并按产品维度汇总库存金额。")
    st.markdown("### 1. 上传文件")
    replenishment_column, price_area = st.columns(2, gap="large")
    with replenishment_column:
        replenishment_file = st.file_uploader("Replenishment 表（两行表头）", type=["xlsx"], key="inventory_replenishment")
    with price_area:
        price_file = st.file_uploader("价格表", type=["xlsx"], key="inventory_prices")
    if replenishment_file is None or price_file is None:
        st.info("请同时上传 Replenishment 表和价格表。上传完成后会自动识别字段并开始计算。")
        return
    try:
        replenishment, prices = read_inventory_files(replenishment_file.getvalue(), price_file.getvalue())
        analyzer = InventoryStructureAnalyzer()
        detected_material, detected_price = analyzer.detect_price_columns(prices)
    except Exception as exc:
        st.error(f"文件读取失败：{exc}")
        return
    st.markdown("### 2. 字段识别与计算状态")
    price_columns = list(prices.columns)
    if not price_columns:
        st.error("价格表没有可用字段。")
        return

    price_file_id = hashlib.sha256(price_file.getvalue()).hexdigest()[:12]
    decision_key = f"price_column_decision_{price_file_id}"
    manual_price_key = f"manual_price_column_{price_file_id}"
    manual_material_key = f"manual_material_column_{price_file_id}"

    if detected_material in price_columns:
        selected_material = detected_material
        st.caption(f"价格表物料代码列：{selected_material}（自动识别）")
    else:
        st.warning("无法自动识别价格表中的物料代码列，请手动选择。")
        selected_material = st.selectbox(
            "价格表物料代码列",
            [None, *price_columns],
            index=0,
            key=manual_material_key,
            format_func=lambda value: "请选择物料代码列" if value is None else str(value),
        )
        if selected_material is None:
            return

    decision = st.session_state.get(decision_key)
    if detected_price in price_columns and decision is None:
        confirm_detected_price_column(detected_price, decision_key)
        st.info(f"已自动识别价格列“{detected_price}”，请在弹出的窗口中确认后继续计算。")
        return

    if detected_price in price_columns and decision == "confirmed":
        selected_price = detected_price
        confirmation_column, change_column = st.columns([3, 1])
        confirmation_column.success(f"已确认价格列：{selected_price}")
        if change_column.button("改为手动选择", key=f"change_price_{price_file_id}", use_container_width=True):
            st.session_state[decision_key] = "manual"
            st.rerun()
    else:
        if detected_price not in price_columns:
            st.warning("无法自动识别价格列，请手动选择价格列后再计算。")
        else:
            st.warning(f"自动识别的“{detected_price}”列未被确认，请手动选择正确的价格列。")
        selected_price = st.selectbox(
            "价格表单价列",
            [None, *price_columns],
            index=0,
            key=manual_price_key,
            format_func=lambda value: "请选择价格列" if value is None else str(value),
        )
        if selected_price is None:
            return
        st.info(f"当前手动选择的价格列：{selected_price}")

    try:
        result = run_inventory_analysis(
            replenishment_file.getvalue(),
            price_file.getvalue(),
            selected_material,
            selected_price,
            INVENTORY_ANALYSIS_CACHE_VERSION,
        )
    except Exception as exc:
        st.error(f"库存结构分析失败：{exc}")
        return
    diagnostics = result.diagnostics
    st.success(
        f"字段识别及计算完成；仅纳入 {diagnostics['纳入Parent记录数']:,} 条 Parent 记录，"
        f"已忽略 {diagnostics['忽略非Parent记录数']:,} 条 Son 或其他级别记录。"
        f"价格缺失 {diagnostics['价格缺失数量']:,} 个，价格冲突 {diagnostics['价格冲突数量']:,} 个，"
        f"金额核对{'一致' if diagnostics['汇总金额核对一致'] else '不一致'}。"
    )
    st.markdown("### 3. 库存结构概要")
    total_column, amount_column = st.columns(2)
    total_column.metric("物料总数", f"{diagnostics['记录数']:,}")
    amount = diagnostics["有效明细金额"]
    amount_column.metric("有效库存金额（百万美元）", "—" if pd.isna(amount) else f"{amount / 1_000_000:,.2f}")

    inventory_metrics = [
        ("库存24个月以上", diagnostics["库存24个月以上数量"]),
        ("超目标库存", diagnostics["超目标库存数量"]),
        ("正常及关注库存", diagnostics["正常及关注库存数量"]),
        ("数据问题", diagnostics["数据问题数量"]),
    ]
    for column, (label, value) in zip(st.columns(4), inventory_metrics):
        column.metric(label, f"{value:,}")
    st.caption("数据问题包括：  \n"
        "有效预测月销缺失或小于等于 0、当前可用库存缺失，"
        "以及存销比低于 24 个月但目标库存月数缺失，导致库存覆盖或区间无法可靠计算。"
    )
    st.markdown("### 4. 库存区间明细")
    st.caption(
        "库存24个月以上：可用库存存销比 ≥ 24个月；  \n"
        "超目标库存：目标库存月数 + 3 < 可用库存存销比 < 24个月；  \n"
        "正常及关注库存：可用库存存销比 ≤ 目标库存月数 + 3。"
    )
    names = ["库存24个月以上", "超目标库存", "正常及关注库存"]
    for tab, name in zip(st.tabs(names), names):
        with tab:
            with st.expander(f"展开查看{name}明细", expanded=False):
                st.dataframe(result.views[name].style.format(precision=2, na_rep=""), use_container_width=True, hide_index=True)
                if name == "正常及关注库存":
                    st.caption("无法计算存销比的记录标记为“数据问题”，并显示在表格最后。")
    st.markdown("### 5. 汇总分析")
    status_tab, group_tab, category_tab = st.tabs(["Product Status", "Product Group", "国贸产品分类码"])
    with status_tab:
        show_inventory_summary("按 Product Status", result.summaries["按产品状态汇总"])
    with group_tab:
        show_inventory_summary("按 Product Group", result.summaries["按产品组汇总"])
    with category_tab:
        show_inventory_summary("按国贸产品分类码", result.summaries["按国贸分类码汇总"])
    st.markdown("### 6. 导出")
    st.download_button(
        "下载库存结构分析结果 (.xlsx)", data=InventoryStructureAnalyzer().export_excel(result),
        file_name="库存结构分析结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
    )


st.set_page_config(page_title="补货与库存分析", page_icon="📦", layout="wide")
page = st.sidebar.radio("一级页面", ["补货预测异常识别", "库存结构分析"])
if page == "补货预测异常识别":
    render_anomaly_page()
else:
    render_inventory_page()
