from __future__ import annotations

import io
import hashlib

import altair as alt
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from forecast_analyzer import AnalysisSettings, analyze, build_views, read_source_excel, to_excel_bytes
from inventory_structure_analysis import (
    SUMMARY_AMOUNT_COLUMN,
    SUMMARY_SHARE_COLUMN,
    InventoryStructureAnalyzer,
    composition_chart_color,
)

INVENTORY_ANALYSIS_CACHE_VERSION = "status-trade-levels-v17"


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


def show_inventory_summary(title: str, frame: pd.DataFrame, summary_type: str) -> None:
    frame = frame.copy()
    if SUMMARY_AMOUNT_COLUMN not in frame.columns and "可用库存金额" in frame.columns:
        frame["可用库存金额"] = frame["可用库存金额"] / 1_000_000
        frame = frame.rename(columns={"可用库存金额": SUMMARY_AMOUNT_COLUMN})
    st.markdown(f"### {title}")
    table_column, chart_column = st.columns([1, 2], gap="large")
    label_columns = [
        column
        for column in frame.columns
        if column not in {SUMMARY_AMOUNT_COLUMN, SUMMARY_SHARE_COLUMN, "型号数量"}
    ]
    column_config = {
        SUMMARY_AMOUNT_COLUMN: st.column_config.NumberColumn(
            "可用库存金额\n（M USD）", format="%.2f", width="small"
        ),
        "型号数量": st.column_config.NumberColumn("型号数量", format="%d", width="small"),
        SUMMARY_SHARE_COLUMN: st.column_config.NumberColumn(
            SUMMARY_SHARE_COLUMN, format="percent", width="small"
        ),
    }
    for column in label_columns:
        column_config[column] = st.column_config.TextColumn(
            column,
            width="medium" if column == "产品组描述" else "small",
        )
    with table_column:
        st.dataframe(
            frame,
            column_config=column_config,
            use_container_width=True,
            hide_index=True,
        )
    with chart_column:
        chart_data = frame.iloc[:-1].copy().sort_values(
            SUMMARY_AMOUNT_COLUMN, ascending=False, na_position="last", kind="stable"
        )
        label_column = label_columns[-1]
        bar_chart = (
            alt.Chart(chart_data)
            .mark_bar()
            .encode(
                x=alt.X(
                    f"{label_column}:N",
                    sort=None,
                    axis=alt.Axis(
                        title=label_column,
                        labelAngle=-45,
                        labelLimit=260,
                        labelOverlap=False,
                        labelSeparation=8,
                    ),
                ),
                y=alt.Y(
                    f"{SUMMARY_AMOUNT_COLUMN}:Q",
                    axis=alt.Axis(title="可用库存金额（M USD）", format=",.2f"),
                ),
                tooltip=[
                    alt.Tooltip(f"{label_column}:N", title=label_column),
                    alt.Tooltip(f"{SUMMARY_AMOUNT_COLUMN}:Q", title="可用库存金额（M USD）", format=",.2f"),
                    alt.Tooltip(f"{SUMMARY_SHARE_COLUMN}:Q", title="库存金额占比", format=".2%"),
                ],
            )
            .properties(height=400)
        )
        total_amount = frame.iloc[-1][SUMMARY_AMOUNT_COLUMN]
        total_text = "—" if pd.isna(total_amount) else f"{total_amount:,.2f} M USD"

        if summary_type == "product_group":
            secondary_data = chart_data.loc[
                chart_data[SUMMARY_AMOUNT_COLUMN].gt(0),
                [label_column, SUMMARY_AMOUNT_COLUMN, SUMMARY_SHARE_COLUMN],
            ].copy()
            secondary_data = secondary_data.sort_values(SUMMARY_AMOUNT_COLUMN, ascending=False, kind="stable")
            inside_mask = secondary_data[SUMMARY_SHARE_COLUMN].fillna(0).ge(0.05)
            secondary_data["图内标签"] = [
                f"{label}<br>{share:.1%}" if is_inside else ""
                for label, share, is_inside in zip(
                    secondary_data[label_column],
                    secondary_data[SUMMARY_SHARE_COLUMN].fillna(0),
                    inside_mask,
                )
            ]
            secondary_data["图例标签"] = [
                f"{label}  {share:.1%}"
                for label, share in zip(
                    secondary_data[label_column], secondary_data[SUMMARY_SHARE_COLUMN].fillna(0)
                )
            ]
            secondary_chart = go.Figure(
                data=[
                    go.Pie(
                        labels=secondary_data["图例标签"],
                        values=secondary_data[SUMMARY_AMOUNT_COLUMN],
                        hole=0.46,
                        domain=dict(x=[0, 0.64], y=[0, 1]),
                        sort=False,
                        direction="clockwise",
                        text=secondary_data["图内标签"],
                        textinfo="text",
                        textposition="inside",
                        insidetextorientation="horizontal",
                        customdata=secondary_data[label_column],
                        hovertemplate=(
                            f"{label_column}：%{{customdata}}<br>"
                            "可用库存金额：%{value:,.2f} M USD<br>"
                            "库存金额占比：%{percent:.2%}<extra></extra>"
                        ),
                        automargin=True,
                    )
                ]
            )
            secondary_chart.update_layout(
                height=max(520, min(700, 400 + 14 * len(secondary_data))),
                margin=dict(l=30, r=25, t=45, b=45),
                showlegend=True,
                legend=dict(
                    title=dict(text="产品组（占比）"),
                    x=0.68,
                    y=0.5,
                    xanchor="left",
                    yanchor="middle",
                    font=dict(size=11),
                    traceorder="normal",
                ),
                annotations=[
                    dict(
                        text=f"总可用库存金额<br><b>{total_text}</b>",
                        x=0.32,
                        y=0.5,
                        xref="paper",
                        yref="paper",
                        showarrow=False,
                        align="center",
                        font=dict(size=14),
                    )
                ],
            )
            secondary_tab_label = "环形图"
        else:
            secondary_chart = go.Figure(
                data=[
                    go.Pie(
                        labels=chart_data[label_column],
                        values=chart_data[SUMMARY_AMOUNT_COLUMN],
                        sort=False,
                        direction="clockwise",
                        textposition="auto",
                        texttemplate="%{label}<br>%{percent:.1%}",
                        insidetextorientation="horizontal",
                        hovertemplate=(
                            f"{label_column}：%{{label}}<br>"
                            "可用库存金额：%{value:,.2f} M USD<br>"
                            "库存金额占比：%{percent:.2%}<extra></extra>"
                        ),
                        automargin=True,
                    )
                ]
            )
            secondary_chart.update_layout(
                height=480,
                margin=dict(l=20, r=190, t=30, b=20),
                showlegend=False,
                annotations=[
                    dict(
                        text=f"总可用库存金额<br><b>{total_text}</b>",
                        x=1.08,
                        y=0.5,
                        xref="paper",
                        yref="paper",
                        xanchor="left",
                        yanchor="middle",
                        showarrow=False,
                        align="left",
                        font=dict(size=14),
                    )
                ],
            )
            secondary_tab_label = "饼图"

        bar_tab, secondary_tab = st.tabs(["柱状图", secondary_tab_label])
        with bar_tab:
            st.altair_chart(bar_chart, use_container_width=True)
        with secondary_tab:
            st.plotly_chart(secondary_chart, use_container_width=True, config={"displayModeBar": False})


def show_trade_top_eighty_dimension(
    overview: pd.DataFrame,
    full_detail: pd.DataFrame,
    trade_diagnostics: dict[str, object],
    dimension: str,
) -> None:
    overview_tab, detail_tab = st.tabs(["前80%分类概览", "数据明细"])
    with overview_tab:
        metric_columns = st.columns(3)
        metric_columns[0].metric(
            "分组总数", f"{int(trade_diagnostics.get('国贸分类码分组总数', len(full_detail))):,}"
        )
        metric_columns[1].metric(
            "累计达到80%的分组数",
            f"{int(trade_diagnostics.get('国贸分类码80%覆盖数量', len(overview))):,}",
        )
        metric_columns[2].metric(
            "数据问题分组数", f"{int(trade_diagnostics.get('国贸数据问题数量', 0)):,}"
        )

        st.write(
            "由于分类数据呈明显的长尾分布，本页聚焦累计贡献前80%库存金额的分类，"
            "以突出主要库存构成和关键影响项。 \n ")

        st.caption(
            "以下分类按可用库存金额从高到低排列，"
            "并包含累计达到80%所需的最后一个分类。"
        )

        st.dataframe(
            overview,
            column_config={
                "排名": st.column_config.NumberColumn(format="%d", width="small"),
                dimension: st.column_config.TextColumn(width="medium"),
                SUMMARY_AMOUNT_COLUMN: st.column_config.NumberColumn(
                    "库存金额\n（M USD）", format="%.2f", width="small"
                ),
                SUMMARY_SHARE_COLUMN: st.column_config.NumberColumn(format="percent", width="small"),
                "累计占比": st.column_config.NumberColumn(format="percent", width="small"),
                "型号数量": st.column_config.NumberColumn(format="%d", width="small"),
                "单型号库存金额（k USD）": st.column_config.NumberColumn(format="%.2f", width="small"),
            },
            use_container_width=True,
            hide_index=True,
            key=f"trade-top80-overview-{dimension}",
        )

        chart_data = overview.copy()
        amount_values = pd.to_numeric(chart_data[SUMMARY_AMOUNT_COLUMN], errors="coerce")
        share_values = pd.to_numeric(chart_data[SUMMARY_SHARE_COLUMN], errors="coerce")
        model_values = pd.to_numeric(chart_data["型号数量"], errors="coerce")
        bar_labels = [
            "" if pd.isna(amount) or pd.isna(share) else f"{amount:,.2f} M  |  {share:.1%}"
            for amount, share in zip(amount_values, share_values)
        ]
        pivot_chart = go.Figure()
        pivot_chart.add_trace(
            go.Bar(
                x=amount_values,
                y=chart_data[dimension],
                orientation="h",
                name="库存金额（柱形）",
                marker_color="#2F75B5",
                text=bar_labels,
                textposition="auto",
                insidetextfont=dict(color="white"),
                cliponaxis=False,
                customdata=pd.DataFrame(
                    {"占比": share_values, "型号数量": model_values}
                ).to_numpy(),
                hovertemplate=(
                    f"{dimension}：%{{y}}<br>"
                    "库存金额：%{x:,.2f} M USD<br>"
                    "库存金额占比：%{customdata[0]:.2%}<br>"
                    "型号数量：%{customdata[1]:,.0f}<extra></extra>"
                ),
            )
        )
        pivot_chart.add_trace(
            go.Scatter(
                x=model_values,
                y=chart_data[dimension],
                name="型号数量（点）",
                mode="markers",
                marker=dict(
                    color="#ED7D31", size=11, symbol="diamond", line=dict(color="white", width=1)
                ),
                xaxis="x2",
                hovertemplate=f"{dimension}：%{{y}}<br>型号数量：%{{x:,.0f}}<extra></extra>",
            )
        )
        pivot_chart.update_layout(
            title="累计前80%分类透视图",
            height=max(420, 48 * len(chart_data) + 140),
            margin=dict(l=200, r=60, t=75, b=115),
            xaxis=dict(title="库存金额（M USD）", tickformat=",.2f", side="top"),
            xaxis2=dict(
                title="型号数量", overlaying="x", side="bottom", tickformat=",.0f", showgrid=False
            ),
            yaxis=dict(title=None, autorange="reversed"),
            legend=dict(
                x=0.99,
                y=-0.14,
                xanchor="right",
                yanchor="top",
                orientation="h",
                bgcolor="rgba(255,255,255,0.82)",
                bordercolor="#D9D9D9",
                borderwidth=1,
            ),
            bargap=0.32,
        )
        st.plotly_chart(
            pivot_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key=f"trade-top80-pivot-{dimension}",
        )

    with detail_tab:
        st.caption("完整分类明细按可用库存金额从高到低排列。")
        st.write("保留两位小数。双击单元格查看完整数值。")
        st.dataframe(
            full_detail,
            column_config={
                "排名": st.column_config.NumberColumn(format="%d", width="small"),
                SUMMARY_AMOUNT_COLUMN: st.column_config.NumberColumn(
                    "库存金额\n（M USD）", format="%.2f", width="small"
                ),
                SUMMARY_SHARE_COLUMN: st.column_config.NumberColumn(format="percent", width="small"),
                "累计占比": st.column_config.NumberColumn(format="percent", width="small"),
                "型号数量": st.column_config.NumberColumn(format="%d", width="small"),
                "单型号库存金额（k USD）": st.column_config.NumberColumn(format="%.2f", width="small"),
            },
            use_container_width=True,
            hide_index=True,
            key=f"trade-top80-detail-{dimension}",
        )


def show_trade_category_dimension(
    summaries: dict[str, pd.DataFrame], diagnostics: dict[str, object], dimension: str
) -> None:
    overview_key, detail_key = InventoryStructureAnalyzer._trade_keys(dimension)
    overview = summaries[overview_key].copy()
    full_detail = summaries[detail_key].copy()
    diagnostics_by_dimension = diagnostics.get("国贸分类诊断", {})
    trade_diagnostics = (
        diagnostics_by_dimension.get(dimension, {})
        if isinstance(diagnostics_by_dimension, dict)
        else {}
    )
    if not trade_diagnostics:
        trade_diagnostics = diagnostics
    if dimension in {
        "国贸产品分类码L2",
        "国贸产品分类码L3",
        "国贸产品分类码L4",
    }:
        show_trade_top_eighty_dimension(overview, full_detail, trade_diagnostics, dimension)
        return
    for frame in (overview, full_detail):
        if "库存分层分类" not in frame.columns and "金额层级" in frame.columns:
            frame.rename(columns={"金额层级": "库存分层分类"}, inplace=True)

    tier_order = [
        "核心（≥1 M USD）",
        "重点（0.5–1 M USD）",
        "一般（0.1–0.5 M USD）",
        "长尾（0–0.1 M USD）",
        "负金额",
        "合计",
    ]
    if "库存分层分类" in overview.columns:
        overview = overview.loc[overview["库存分层分类"].isin(tier_order)].copy()
        overview["_分层顺序"] = pd.Categorical(
            overview["库存分层分类"], categories=tier_order, ordered=True
        )
        overview = overview.sort_values("_分层顺序").drop(columns="_分层顺序").reset_index(drop=True)

    overview_tab, detail_tab = st.tabs(["库存分类概览", "数据明细"])
    with overview_tab:
        metric_columns = st.columns(3)
        metric_columns[0].metric("达到累计80%的分类码", f"{trade_diagnostics['国贸分类码80%覆盖数量']:,}")
        metric_columns[1].metric("核心及重点分类码（≥0.5 M USD）", f"{trade_diagnostics['国贸重点分类码数量']:,}")
        metric_columns[2].metric("数据问题分类码", f"{trade_diagnostics['国贸数据问题数量']:,}")

        st.dataframe(
            overview,
            column_config={
                "库存分层分类": st.column_config.TextColumn(width="medium"),
                "分类码数量": st.column_config.NumberColumn(format="%d", width="small"),
                "型号数量": st.column_config.NumberColumn(format="%d", width="small"),
                SUMMARY_AMOUNT_COLUMN: st.column_config.NumberColumn(
                    "库存金额\n（M USD）", format="%.2f", width="small"
                ),
                SUMMARY_SHARE_COLUMN: st.column_config.NumberColumn(format="percent", width="small"),
                "单型号平均金额（k USD）": st.column_config.NumberColumn(format="%.2f", width="small"),
            },
            use_container_width=True,
            hide_index=True,
            key=f"trade-overview-{dimension}",
        )

        tier_data = overview.iloc[:-1].copy()
        tier_data = tier_data.loc[
            tier_data[SUMMARY_AMOUNT_COLUMN].notna() | tier_data["型号数量"].notna()
        ].copy()
        st.markdown("#### 库存分层透视图")
        amount_values = pd.to_numeric(tier_data[SUMMARY_AMOUNT_COLUMN], errors="coerce")
        share_values = pd.to_numeric(tier_data[SUMMARY_SHARE_COLUMN], errors="coerce")
        model_values = pd.to_numeric(tier_data["型号数量"], errors="coerce")
        bar_labels = [
            "" if pd.isna(amount) or pd.isna(share) else f"{amount:,.2f} M  |  {share:.1%}"
            for amount, share in zip(amount_values, share_values)
        ]
        tier_colors = {
            "核心（≥1 M USD）": "#2F75B5",
            "重点（0.5–1 M USD）": "#5B9BD5",
            "一般（0.1–0.5 M USD）": "#70AD47",
            "长尾（0–0.1 M USD）": "#A5A5A5",
            "负金额": "#C00000",
        }
        pivot_chart = go.Figure()
        pivot_chart.add_trace(
            go.Bar(
                x=amount_values,
                y=tier_data["库存分层分类"],
                orientation="h",
                name="库存金额（柱形）",
                marker_color=[tier_colors.get(value, "#5B9BD5") for value in tier_data["库存分层分类"]],
                text=bar_labels,
                textposition="auto",
                insidetextfont=dict(color="white"),
                cliponaxis=False,
                customdata=pd.DataFrame(
                    {"占比": share_values, "型号数量": model_values}
                ).to_numpy(),
                hovertemplate=(
                    "库存分层分类：%{y}<br>"
                    "库存金额：%{x:,.2f} M USD<br>"
                    "库存金额占比：%{customdata[0]:.2%}<br>"
                    "型号数量：%{customdata[1]:,.0f}<extra></extra>"
                ),
            )
        )
        pivot_chart.add_trace(
            go.Scatter(
                x=model_values,
                y=tier_data["库存分层分类"],
                name="型号数量（点）",
                mode="markers",
                marker=dict(color="#ED7D31", size=11, symbol="diamond", line=dict(color="white", width=1)),
                xaxis="x2",
                hovertemplate="库存分层分类：%{y}<br>型号数量：%{x:,.0f}<extra></extra>",
            )
        )
        pivot_chart.update_layout(
            height=max(360, 70 * len(tier_data) + 100),
            margin=dict(l=175, r=70, t=70, b=55),
            xaxis=dict(title="库存金额（M USD）", tickformat=",.2f", side="bottom"),
            xaxis2=dict(
                title="型号数量",
                overlaying="x",
                side="top",
                tickformat=",.0f",
                showgrid=False,
            ),
            yaxis=dict(title=None, autorange="reversed"),
            legend=dict(
                x=0.99,
                y=0.99,
                xanchor="right",
                yanchor="top",
                bgcolor="rgba(255,255,255,0.75)",
                bordercolor="#D9D9D9",
                borderwidth=1,
            ),
            bargap=0.35,
        )
        st.plotly_chart(
            pivot_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key=f"trade-pivot-{dimension}",
        )

    detail_column_config = {
        "排名": st.column_config.NumberColumn(format="%d", width="small"),
        SUMMARY_AMOUNT_COLUMN: st.column_config.NumberColumn(
            "库存金额\n（M USD）", format="%.2f", width="small"
        ),
        SUMMARY_SHARE_COLUMN: st.column_config.NumberColumn(format="percent", width="small"),
        "累计占比": st.column_config.NumberColumn(format="percent", width="small"),
        "型号数量": st.column_config.NumberColumn(format="%d", width="small"),
        "单型号库存金额（k USD）": st.column_config.NumberColumn(format="%.2f", width="small"),
    }
    with detail_tab:
        st.caption(
            "按可用库存金额从高到低排列；库存分层分类包含核心、重点、一般、长尾和负金额。"
            "金额缺失的记录不强行分层，并在数据状态中标记。"
        )
        st.dataframe(
            full_detail,
            column_config=detail_column_config,
            use_container_width=True,
            hide_index=True,
            key=f"trade-detail-{dimension}",
        )


def show_trade_l1_summary(
    summaries: dict[str, pd.DataFrame], diagnostics: dict[str, object]
) -> None:
    frame = summaries["国贸分类码L1汇总"].copy()
    chart_data = frame.iloc[:-1].copy()
    chart_data[SUMMARY_AMOUNT_COLUMN] = pd.to_numeric(
        chart_data[SUMMARY_AMOUNT_COLUMN], errors="coerce"
    )
    chart_data[SUMMARY_SHARE_COLUMN] = pd.to_numeric(
        chart_data[SUMMARY_SHARE_COLUMN], errors="coerce"
    )
    chart_data["型号数量"] = pd.to_numeric(chart_data["型号数量"], errors="coerce")

    total_row = frame.iloc[-1]
    metric_columns = st.columns(3)
    diagnostics_by_dimension = diagnostics.get("国贸分类诊断", {})
    l1_diagnostics = (
        diagnostics_by_dimension.get("国贸产品分类码L1", {})
        if isinstance(diagnostics_by_dimension, dict)
        else {}
    )
    metric_columns[0].metric(
        "分组总数", f"{int(l1_diagnostics.get('国贸分类码分组总数', len(chart_data))):,}"
    )
    metric_columns[1].metric(
        "总可用库存金额",
        f"{pd.to_numeric(total_row[SUMMARY_AMOUNT_COLUMN], errors='coerce'):,.2f} M USD",
    )
    metric_columns[2].metric(
        "型号数量", f"{pd.to_numeric(total_row['型号数量'], errors='coerce'):,.0f}"
    )

    st.dataframe(
        frame,
        column_config={
            "国贸产品分类码L1": st.column_config.TextColumn(width="medium"),
            SUMMARY_AMOUNT_COLUMN: st.column_config.NumberColumn(
                "库存金额\n（M USD）", format="%.2f", width="small"
            ),
            SUMMARY_SHARE_COLUMN: st.column_config.NumberColumn(format="percent", width="small"),
            "型号数量": st.column_config.NumberColumn(format="%d", width="small"),
            "单型号库存金额（k USD）": st.column_config.NumberColumn(format="%.2f", width="small"),
        },
        use_container_width=True,
        hide_index=True,
        key="trade-l1-summary",
    )

    chart_data = chart_data.sort_values(
        SUMMARY_AMOUNT_COLUMN, ascending=True, na_position="first", kind="stable"
    )
    labels = [
        "" if pd.isna(amount) or pd.isna(share) else f"{amount:,.2f} M  |  {share:.1%}"
        for amount, share in zip(chart_data[SUMMARY_AMOUNT_COLUMN], chart_data[SUMMARY_SHARE_COLUMN])
    ]
    bar_chart = go.Figure(
        go.Bar(
            x=chart_data[SUMMARY_AMOUNT_COLUMN],
            y=chart_data["国贸产品分类码L1"],
            orientation="h",
            marker_color="#2F75B5",
            text=labels,
            textposition="auto",
            insidetextfont=dict(color="white"),
            cliponaxis=False,
            customdata=chart_data[[SUMMARY_SHARE_COLUMN, "型号数量"]].to_numpy(),
            hovertemplate=(
                "L1分类：%{y}<br>库存金额：%{x:,.2f} M USD<br>"
                "库存金额占比：%{customdata[0]:.2%}<br>型号数量：%{customdata[1]:,.0f}<extra></extra>"
            ),
        )
    )
    bar_chart.update_layout(
        height=max(360, 62 * len(chart_data) + 100),
        margin=dict(l=190, r=80, t=30, b=55),
        xaxis=dict(title="库存金额（M USD）", tickformat=",.2f"),
        yaxis=dict(title=None),
        showlegend=False,
    )

    pie_data = chart_data.loc[chart_data[SUMMARY_AMOUNT_COLUMN].gt(0)].sort_values(
        SUMMARY_AMOUNT_COLUMN, ascending=False, kind="stable"
    )
    pie_positions = [
        "inside" if share >= 0.03 else "outside"
        for share in pie_data[SUMMARY_SHARE_COLUMN].fillna(0)
    ]
    pie_labels = [
        f"{category}<br>{share:.1%}" if share >= 0.01 else f"{share:.1%}"
        for category, share in zip(
            pie_data["国贸产品分类码L1"], pie_data[SUMMARY_SHARE_COLUMN].fillna(0)
        )
    ]
    pie_chart = go.Figure(
        go.Pie(
            labels=pie_data["国贸产品分类码L1"],
            values=pie_data[SUMMARY_AMOUNT_COLUMN],
            sort=False,
            text=pie_labels,
            textinfo="text",
            textposition=pie_positions,
            insidetextorientation="horizontal",
            automargin=True,
            hovertemplate=(
                "L1分类：%{label}<br>库存金额：%{value:,.2f} M USD<br>"
                "库存金额占比：%{percent:.2%}<extra></extra>"
            ),
        )
    )
    pie_chart.update_layout(
        height=500,
        margin=dict(l=45, r=185, t=70, b=45),
        showlegend=True,
        legend=dict(x=1.02, y=0.5, xanchor="left", yanchor="middle"),
        uniformtext_minsize=10,
        uniformtext_mode="hide",
    )

    bar_tab, pie_tab = st.tabs(["柱状图", "饼图"])
    with bar_tab:
        st.plotly_chart(
            bar_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key="trade-l1-bar",
        )
    with pie_tab:
        st.plotly_chart(
            pie_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key="trade-l1-pie",
        )


def show_trade_category_analysis(
    summaries: dict[str, pd.DataFrame], diagnostics: dict[str, object]
) -> None:
    st.markdown("### 按国贸产品分类码分类")
    dimensions = list(diagnostics.get("国贸分类维度", ("国贸产品分类码",)))
    if len(dimensions) == 1:
        if dimensions[0] == "国贸产品分类码L1":
            show_trade_l1_summary(summaries, diagnostics)
        else:
            show_trade_category_dimension(summaries, diagnostics, dimensions[0])
        return

    level_tabs = st.tabs([dimension.removeprefix("国贸产品分类码") for dimension in dimensions])
    for tab, dimension in zip(level_tabs, dimensions):
        with tab:
            if dimension == "国贸产品分类码L1":
                show_trade_l1_summary(summaries, diagnostics)
            else:
                show_trade_category_dimension(summaries, diagnostics, dimension)


def show_status_trade_analysis(result, status: str, dimension: str, selected_groups: list[str]) -> None:
    analyzer = InventoryStructureAnalyzer()
    selected_detail = result.detail.loc[
        result.detail["产品组描述"].astype("string").str.strip().fillna("").replace("", "未分类").isin(selected_groups)
    ]
    summary = analyzer.summarize_excel_status_trade(selected_detail, status, dimension)
    st.caption(f"按产品组和{dimension}汇总；金额单位：M USD。图中显示所有具有正库存金额的分类码，不再合并为“其他分类码”。")
    if summary.empty:
        st.info(f"没有产品状态为 {status} 的记录。")
        return
    st.dataframe(
        summary.style.format({SUMMARY_AMOUNT_COLUMN: "{:,.2f}", "组内金额占比": "{:.2%}"}, na_rep=""),
        width="content",
        hide_index=True,
    )
    pivot = analyzer.build_status_group_composition(summary, dimension)
    if pivot.empty:
        st.info("没有可绘图的正库存金额；负金额和缺失金额仍保留在上方表格中。")
        return
    category_columns = list(pivot.columns[2:])
    group_amounts = pivot[category_columns].sum(axis=1)
    total_amount = group_amounts.sum()
    labels = [
        f"{description}（{amount / total_amount:.1%}）"
        for description, amount in zip(pivot["产品组描述"], group_amounts)
    ]
    chart = go.Figure()
    for index, category in enumerate(category_columns):
        chart.add_trace(go.Bar(
            name=str(category),
            x=pivot[category] / group_amounts,
            y=labels,
            orientation="h",
            marker_color=composition_chart_color(str(category), index),
            customdata=pivot[category],
            hovertemplate="产品组：%{y}<br>分类码：" + str(category)
            + "<br>组内占比：%{x:.1%}<br>库存金额：%{customdata:,.2f} M USD<extra></extra>",
        ))
    chart.update_layout(
        barmode="stack",
        height=max(360, 38 * len(pivot) + 180, 18 * min(len(category_columns), 30) + 80),
        margin=dict(l=180, r=270, t=45, b=70),
        xaxis=dict(title="库存金额构成占比", tickformat=".0%", range=[0, 1]),
        yaxis=dict(title="产品组", autorange="reversed"),
        legend=dict(
            title="分类码", orientation="v", x=1.02, xanchor="left",
            y=1, yanchor="top", maxheight=0.88,
        ),
        title=f"{status} · {dimension}：产品组内分类码构成",
    )
    st.plotly_chart(chart, use_container_width=True, config={"displayModeBar": False}, key=f"status-trade-{status}-{dimension}")


def show_status_inventory_analysis(result, status: str, selected_groups: list[str]) -> None:
    analyzer = InventoryStructureAnalyzer()
    selected_detail = result.detail.loc[
        result.detail["产品组描述"].astype("string").str.strip().fillna("").replace("", "未分类").isin(selected_groups)
    ]
    summary = analyzer.summarize_excel_status_inventory(selected_detail)
    st.caption("按泛欧ABC等级展示三种库存水位；未填写ABC等级或无法归入库存区间的记录不进入此图。库存金额单位：M USD。")
    if summary.empty:
        st.info("所选产品组没有可展示的 A/B/C 库存区间记录。")
        return
    colors = {"库存24个月以上": "#d95f5f", "超目标库存": "#f2b34d", "正常及关注库存": "#4c9e91"}
    for grade_tab, grade in zip(st.tabs(["A级", "B级", "C级"]), ["A", "B", "C"]):
        with grade_tab:
            grade_data = summary.loc[summary["ABC等级"].eq(grade)]
            if grade_data.empty:
                st.info(f"所选产品组没有 {grade} 级库存记录。")
                continue
            st.dataframe(
                grade_data.style.format({SUMMARY_AMOUNT_COLUMN: "{:,.2f}"}, na_rep=""),
                width="content",
                hide_index=True,
            )
            chart = go.Figure()
            displayed_groups = grade_data[["产品组", "产品组描述"]].drop_duplicates().itertuples(index=False, name=None)
            displayed_groups = list(displayed_groups)
            grade_positive_total = grade_data[SUMMARY_AMOUNT_COLUMN].clip(lower=0).sum()
            group_labels = []
            for code, description in displayed_groups:
                group_rows = grade_data.loc[
                    grade_data["产品组"].eq(code) & grade_data["产品组描述"].eq(description)
                ]
                group_positive = group_rows[SUMMARY_AMOUNT_COLUMN].clip(lower=0).sum()
                share = group_positive / grade_positive_total if grade_positive_total > 0 else 0.0
                group_labels.append(f"{description}（{share:.1%}）")
            for band, color in colors.items():
                values = []
                counts = []
                for code, description in displayed_groups:
                    matching = grade_data.loc[
                        grade_data["产品组"].eq(code)
                        & grade_data["产品组描述"].eq(description)
                        & grade_data["库存区间"].eq(band)
                    ]
                    amount = matching[SUMMARY_AMOUNT_COLUMN].sum(min_count=1)
                    values.append(float(amount) if pd.notna(amount) else 0.0)
                    counts.append(int(matching["型号数量"].sum()))
                chart.add_trace(go.Bar(
                    name=band,
                    x=values,
                    y=group_labels,
                    orientation="h",
                    marker_color=color,
                    customdata=counts,
                    hovertemplate="产品组：%{y}<br>库存金额：%{x:,.2f} M USD<br>型号数量：%{customdata:,}<extra></extra>",
                ))
            chart.update_layout(
                barmode="stack",
                height=max(400, 44 * len(displayed_groups) + 100),
                margin=dict(l=160, r=40, t=25, b=60),
                xaxis_title="可用库存金额（M USD）",
                yaxis=dict(title="产品组", autorange="reversed"),
                legend_title="库存水位",
            )
            st.plotly_chart(chart, use_container_width=True, config={"displayModeBar": False}, key=f"status-inventory-{status}-{grade}")


def show_product_status_analysis(result) -> None:
    analyzer = InventoryStructureAnalyzer()
    status_names = ["NEW RELEASE (RECOMMEND)", "NORMAL/PHASING OUT", "EOL"]
    trade_levels = [f"国贸产品分类码L{level}" for level in range(1, 5)]
    for tab, status in zip(st.tabs(status_names), status_names):
        with tab:
            groups = analyzer.available_product_groups(result.detail, status)
            if not groups:
                st.info(f"没有产品状态为 {status} 的记录。")
                continue
            selected_groups = st.multiselect(
                "产品组描述（默认全部）",
                groups,
                default=groups,
                key=f"inventory-groups-{status}",
            )
            if not selected_groups:
                st.info("请选择至少一个产品组。")
                continue
            labels = ["L1", "L2", "L3", "L4"]
            if status == "NORMAL/PHASING OUT":
                labels.append("A/B/C库存等级")
            for analysis_tab, label in zip(st.tabs(labels), labels):
                with analysis_tab:
                    if label == "A/B/C库存等级":
                        show_status_inventory_analysis(result, status, selected_groups)
                        continue
                    dimension = trade_levels[int(label[1]) - 1]
                    if dimension not in result.detail.columns:
                        if label == "L1" and "国贸产品分类码" in result.detail.columns:
                            st.caption("源文件未提供 L1；此处使用原“国贸产品分类码”字段。")
                            dimension = "国贸产品分类码"
                        else:
                            st.info(f"源文件未提供{dimension}，无法展示该层级。")
                            continue
                    show_status_trade_analysis(result, status, dimension, selected_groups)


def show_status_overview(result) -> None:
    summary = InventoryStructureAnalyzer().summarize_status_overview(result.detail)
    st.markdown("#### 按产品状态概览")
    table_column, pie_column = st.columns([1, 1], gap="large")
    with table_column:
        st.dataframe(
            summary.style.format(
                {SUMMARY_AMOUNT_COLUMN: "{:,.2f}", SUMMARY_SHARE_COLUMN: "{:.2%}"}, na_rep=""
            ),
            width="content",
            hide_index=True,
        )
    with pie_column:
        chart_data = summary.iloc[:-1].loc[summary.iloc[:-1][SUMMARY_AMOUNT_COLUMN].gt(0)]
        if chart_data.empty:
            st.info("没有可用于饼图的正库存金额。")
        else:
            pie = go.Figure(go.Pie(
                labels=chart_data["产品状态类别"],
                values=chart_data[SUMMARY_AMOUNT_COLUMN],
                textinfo="label+percent",
                textposition="auto",
                hovertemplate="%{label}<br>库存金额：%{value:,.2f} M USD<br>占比：%{percent:.2%}<extra></extra>",
            ))
            pie.update_layout(
                height=340,
                margin=dict(l=20, r=20, t=15, b=20),
                legend=dict(orientation="h", y=-0.1),
            )
            st.plotly_chart(pie, use_container_width=True, config={"displayModeBar": False}, key="inventory-status-overview")
            if summary.iloc[:-1][SUMMARY_AMOUNT_COLUMN].lt(0).any():
                st.caption("负库存金额保留在表格中，不绘入饼图。")


def render_inventory_page() -> None:
    st.title("库存结构分析")
    st.caption("按可用库存覆盖月数识别长期库存、超目标库存及正常/关注库存，并按产品维度汇总库存金额。")
    st.markdown("### 1. 上传文件")
    replenishment_column, price_area = st.columns(2, gap="large")
    with replenishment_column:
        replenishment_file = st.file_uploader("Replenishment 表", type=["xlsx"], key="inventory_replenishment")
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
    st.write(
    "库存24个月以上：可用库存存销比 ≥ 24个月；  \n"
    "超目标库存：目标库存月数 + 3 < 可用库存存销比 < 24个月；  \n"
    "正常及关注库存：可用库存存销比 ≤ 目标库存月数 + 3。"
    )
    total_column, amount_column = st.columns(2)
    total_column.metric("物料总数", f"{diagnostics['记录数']:,}")
    amount = diagnostics["有效明细金额"]
    amount_column.metric("可用库存金额（百万美元）", "—" if pd.isna(amount) else f"{amount / 1_000_000:,.2f}")

    inventory_metrics = [
        ("库存24个月以上数量", diagnostics["库存24个月以上数量"]),
        ("超目标库存数量", diagnostics["超目标库存数量"]),
        ("正常及关注库存数量", diagnostics["正常及关注库存数量"]),
        ("数据问题数量", diagnostics["数据问题数量"]),
    ]
    for column, (label, value) in zip(st.columns(4), inventory_metrics):
        band = label.removesuffix("数量")
        band_amount = diagnostics["库存区间金额"][band]
        share_text = "—" if pd.isna(band_amount) or pd.isna(amount) or amount == 0 else f"{band_amount / amount:.2%}"
        column.metric(label, f"{value:,}")
        column.metric("金额（M USD）", "—" if pd.isna(band_amount) else f"{band_amount / 1_000_000:,.2f}")
        column.metric("占比", share_text)
    st.caption("数据问题包括：  \n"
        "有效预测月销缺失或小于等于 0、当前可用库存缺失，"
        "以及存销比低于 24 个月但目标库存月数缺失，导致库存覆盖或区间无法可靠计算。"
    )
    st.markdown("### 4. 库存区间明细")

    names = ["库存24个月以上", "超目标库存", "正常及关注库存", "无销量", "无库存等级"]
    for tab, name in zip(st.tabs(names), names):
        with tab:
            with st.expander(f"展开查看{name}明细（{len(result.views[name]):,} 条）", expanded=False):
                web_detail = result.views[name].drop(columns=["销售组织"], errors="ignore")
                st.dataframe(
                    web_detail.style.format(precision=2, na_rep=""),
                    use_container_width=True,
                    hide_index=True,
                )
                if name == "正常及关注库存":
                    st.caption("无法计算存销比的记录标记为“数据问题”，并显示在表格最后。")
                elif name == "无销量":
                    st.caption("有效预测月销缺失或小于等于 0；价格缺失不影响本表的筛选。")
                elif name == "无库存等级":
                    st.caption("当前库存等级为空；价格缺失不影响本表的筛选。")
    st.markdown("### 5. 汇总分析")
    show_status_overview(result)
    show_product_status_analysis(result)
    st.markdown("### 6. 导出")
    st.download_button(
        "下载库存结构分析结果 (.xlsx)", data=InventoryStructureAnalyzer().export_excel(result),
        file_name="库存结构分析结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
    )


st.set_page_config(page_title="补货与库存分析", page_icon="📦", layout="wide")
page = st.sidebar.radio("页面选择：", ["补货预测异常识别", "库存结构分析"])
if page == "补货预测异常识别":
    render_anomaly_page()
else:
    render_inventory_page()
