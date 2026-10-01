from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import BinaryIO, Hashable

import numpy as np
import pandas as pd


DETAIL_COLUMNS = [
    "物料",
    "物料描述",
    "产品状态",
    "产品组",
    "产品组描述",
    "国贸产品分类码",
    "当前库存等级",
    "目标库存月数",
    "当前可用库存",
    "在途+在产",
    "有效预测月销",
    "可用库存存销比",
    "全供应链存销比",
    "单价",
    "可用库存金额",
    "库存区间",
    "计算状态",
]

SUMMARY_AMOUNT_COLUMN = "可用库存金额（M USD）"
SUMMARY_SHARE_COLUMN = "库存金额占比"

SHEET_NAMES = [
    "库存24个月以上",
    "超目标库存",
    "正常及关注库存",
    "按产品状态汇总",
    "按产品组汇总",
    "国贸分类码分层",
    "按国贸分类码汇总",
]


@dataclass(frozen=True)
class InventoryAnalysisResult:
    detail: pd.DataFrame
    views: dict[str, pd.DataFrame]
    summaries: dict[str, pd.DataFrame]
    field_mapping: dict[str, str]
    diagnostics: dict[str, int | float | bool]


class InventoryStructureAnalyzer:
    """Analyze available-stock structure without mutating the uploaded workbooks."""

    FIELD_ALIASES = {
        "物料": ["物料", "物料代码", "material", "materialcode", "sku"],
        "物料描述": ["物料描述", "产品描述", "materialdescription", "description"],
        "产品状态": ["产品状态", "productstatus"],
        "产品组": ["产品组", "productgroup", "productgroupcode"],
        "产品组描述": ["产品组描述", "productgroupdescription", "productgroupdesc"],
        "国贸产品分类码": ["国贸产品分类码", "产品分类码", "internationalproductclassificationcode"],
        "级别": ["级别", "层级", "level", "hierarchylevel"],
        "当前库存等级": ["当前库存等级", "currentstocklevel", "currentstockgrade"],
        "目标库存月数": ["库存月数", "目标库存月数", "targetinventorymonths", "targetmos"],
        "当前可用库存": ["当前可用库存", "可用库存", "availableinventory", "availablestock"],
        "在途+在产": ["在途+在产", "在途在产", "pipeline", "intransit+inproduction"],
        "月预测": ["月预测", "monthlyforecast"],
        "月预测（修改）": ["月预测（修改）", "月预测(修改)", "修改月预测", "revisedmonthlyforecast"],
        "全供应链存销比": ["全供应链存销比", "供应链存销比", "fullsupplychainmos"],
    }
    PRICE_MATERIAL_ALIASES = ["物料", "物料代码", "material", "materialcode", "sku", "skucode"]
    PRICE_ALIASES = ["单价", "价格", "unitprice", "price", "标准价", "标准单价"]

    @staticmethod
    def read_replenishment_excel(source: str | BinaryIO) -> pd.DataFrame:
        return pd.read_excel(source, header=[0, 1])

    @staticmethod
    def read_price_excel(source: str | BinaryIO) -> pd.DataFrame:
        return pd.read_excel(source)

    @staticmethod
    def _norm(value: object) -> str:
        if pd.isna(value):
            return ""
        return re.sub(r"[\s_\-—/\\（）()]+", "", str(value)).casefold()

    @staticmethod
    def _display_column(column: Hashable) -> str:
        if isinstance(column, tuple):
            return " / ".join(str(part) for part in column if not str(part).startswith("Unnamed"))
        return str(column)

    @classmethod
    def _leaf(cls, column: Hashable) -> str:
        if isinstance(column, tuple):
            values = [part for part in column if not str(part).startswith("Unnamed")]
            return cls._norm(values[-1] if values else "")
        return cls._norm(column)

    def detect_replenishment_fields(self, frame: pd.DataFrame) -> dict[str, Hashable]:
        mapping: dict[str, Hashable] = {}
        used: set[Hashable] = set()
        for canonical, aliases in self.FIELD_ALIASES.items():
            alias_norms = {self._norm(alias) for alias in aliases}
            candidates = [column for column in frame.columns if self._leaf(column) in alias_norms and column not in used]
            if candidates:
                mapping[canonical] = candidates[0]
                used.add(candidates[0])
        missing = [name for name in self.FIELD_ALIASES if name not in mapping]
        if missing:
            raise ValueError("Replenishment 表缺少必要字段：" + "、".join(missing))
        return mapping

    def detect_price_columns(self, frame: pd.DataFrame) -> tuple[Hashable | None, Hashable | None]:
        def find(aliases: list[str]) -> Hashable | None:
            norms = {self._norm(alias) for alias in aliases}
            return next((column for column in frame.columns if self._leaf(column) in norms), None)

        return find(self.PRICE_MATERIAL_ALIASES), find(self.PRICE_ALIASES)

    @staticmethod
    def _material_key(value: object) -> object:
        if pd.isna(value):
            return pd.NA
        text = str(value).strip()
        if re.fullmatch(r"[-+]?\d+\.0+", text):
            text = text.split(".", 1)[0]
        return text.casefold()

    @staticmethod
    def _add_status(statuses: list[list[str]], mask: pd.Series | np.ndarray, message: str) -> None:
        for index in np.flatnonzero(np.asarray(mask, dtype=bool)):
            if message not in statuses[index]:
                statuses[index].append(message)

    def _prepare_prices(
        self, frame: pd.DataFrame, material_column: Hashable, price_column: Hashable
    ) -> tuple[pd.DataFrame, set[object]]:
        prices = pd.DataFrame(
            {
                "_物料键": frame[material_column].map(self._material_key),
                "_价格": pd.to_numeric(frame[price_column], errors="coerce"),
            }
        ).dropna(subset=["_物料键"])

        records: list[dict[str, object]] = []
        conflicts: set[object] = set()
        for key, group in prices.groupby("_物料键", sort=False, dropna=False):
            unique_prices = group["_价格"].dropna().unique()
            if len(unique_prices) > 1:
                conflicts.add(key)
                price = np.nan
            elif len(unique_prices) == 1:
                price = float(unique_prices[0])
            else:
                price = np.nan
            records.append({"_物料键": key, "单价": price})
        return pd.DataFrame(records, columns=["_物料键", "单价"]), conflicts

    def analyze(
        self,
        replenishment: pd.DataFrame,
        prices: pd.DataFrame,
        price_material_column: Hashable | None = None,
        price_column: Hashable | None = None,
    ) -> InventoryAnalysisResult:
        field_map = self.detect_replenishment_fields(replenishment)
        detected_material, detected_price = self.detect_price_columns(prices)
        price_material_column = price_material_column if price_material_column is not None else detected_material
        price_column = price_column if price_column is not None else detected_price
        if price_material_column not in prices.columns or price_column not in prices.columns:
            raise ValueError("无法识别价格表中的物料代码列或单价列，请在页面中手动选择。")

        detail = pd.DataFrame({name: replenishment[column] for name, column in field_map.items()})
        source_record_count = len(detail)
        parent_mask = detail["级别"].astype("string").str.strip().str.casefold().eq("parent").fillna(False)
        detail = detail.loc[parent_mask].copy().reset_index(drop=True)
        ignored_record_count = source_record_count - len(detail)
        detail = detail.rename(columns={"月预测（修改）": "_修改预测", "月预测": "_原预测"})
        detail["物料"] = detail["物料"].where(detail["物料"].notna(), pd.NA)
        detail["_物料键"] = detail["物料"].map(self._material_key)

        numeric_columns = [
            "目标库存月数", "当前可用库存", "在途+在产", "_修改预测", "_原预测", "全供应链存销比"
        ]
        for column in numeric_columns:
            detail[column] = pd.to_numeric(detail[column], errors="coerce")

        detail["有效预测月销"] = detail["_修改预测"].where(detail["_修改预测"].notna(), detail["_原预测"])
        statuses: list[list[str]] = [[] for _ in range(len(detail))]

        stock_missing = detail["当前可用库存"].isna()
        forecast_missing = detail["有效预测月销"].isna()
        invalid_forecast = detail["有效预测月销"].le(0) & detail["当前可用库存"].gt(0)
        nonpositive_other = detail["有效预测月销"].le(0) & ~detail["当前可用库存"].gt(0) & ~stock_missing
        self._add_status(statuses, stock_missing, "可用库存缺失")
        self._add_status(statuses, forecast_missing, "预测月销缺失")
        self._add_status(statuses, invalid_forecast, "无有效销量，库存覆盖无法计算")
        self._add_status(statuses, nonpositive_other, "无有效销量，库存覆盖无法计算")

        valid_ratio = detail["当前可用库存"].notna() & detail["有效预测月销"].gt(0)
        detail["可用库存存销比"] = np.where(
            valid_ratio, detail["当前可用库存"] / detail["有效预测月销"], np.nan
        )

        target_missing = detail["目标库存月数"].isna() & detail["可用库存存销比"].lt(24)
        self._add_status(statuses, target_missing, "目标库存月数缺失")
        ratio = detail["可用库存存销比"]
        detail["库存区间"] = pd.Series(pd.NA, index=detail.index, dtype="object")
        detail.loc[ratio.ge(24), "库存区间"] = "库存24个月以上"
        detail.loc[ratio.lt(24) & detail["目标库存月数"].notna() & ratio.gt(detail["目标库存月数"] + 3), "库存区间"] = "超目标库存"
        detail.loc[ratio.notna() & detail["目标库存月数"].notna() & ratio.le(detail["目标库存月数"] + 3), "库存区间"] = "正常及关注库存"
        detail["库存区间"] = detail["库存区间"].fillna("数据问题")

        price_lookup, conflicts = self._prepare_prices(prices, price_material_column, price_column)
        detail = detail.merge(price_lookup, how="left", on="_物料键", validate="many_to_one")
        conflict_mask = detail["_物料键"].isin(conflicts)
        price_missing = detail["单价"].isna() & ~conflict_mask
        self._add_status(statuses, conflict_mask, "价格冲突")
        self._add_status(statuses, price_missing, "价格缺失")
        detail["可用库存金额"] = detail["当前可用库存"] * detail["单价"]
        detail["计算状态"] = ["；".join(items) if items else "正常" for items in statuses]

        detail = detail[DETAIL_COLUMNS + ["_物料键"]]
        views = {
            "库存24个月以上": self._sorted(detail.loc[detail["库存区间"].eq("库存24个月以上")]),
            "超目标库存": self._sorted(detail.loc[detail["库存区间"].eq("超目标库存")]),
        }
        normal = detail.loc[detail["库存区间"].eq("正常及关注库存")]
        problems = detail.loc[detail["库存区间"].eq("数据问题")]
        views["正常及关注库存"] = pd.concat([self._sorted(normal), self._sorted(problems)], ignore_index=True)
        views = {name: frame[DETAIL_COLUMNS].copy() for name, frame in views.items()}

        summaries = {
            "按产品状态汇总": self._summarize(detail, ["产品状态"]),
            "按产品组汇总": self._summarize(detail, ["产品组", "产品组描述"]),
            "按国贸分类码汇总": self._summarize(detail, ["国贸产品分类码"]),
        }
        trade_tables, trade_diagnostics = self._build_trade_category_tables(
            summaries["按国贸分类码汇总"]
        )
        summaries.update(trade_tables)
        detail_amount = detail["可用库存金额"].sum(min_count=1)
        summary_amount_musd = summaries["按产品状态汇总"].iloc[-1][SUMMARY_AMOUNT_COLUMN]
        summary_amount = summary_amount_musd * 1_000_000
        amounts_match = (pd.isna(detail_amount) and pd.isna(summary_amount)) or bool(np.isclose(detail_amount, summary_amount))
        diagnostics = {
            "源文件记录数": int(source_record_count),
            "纳入Parent记录数": int(len(detail)),
            "忽略非Parent记录数": int(ignored_record_count),
            "记录数": int(detail["_物料键"].nunique(dropna=True)),
            "价格缺失数量": int(detail.loc[price_missing, "_物料键"].nunique(dropna=True)),
            "价格冲突数量": int(detail.loc[conflict_mask, "_物料键"].nunique(dropna=True)),
            "库存24个月以上数量": int(detail.loc[detail["库存区间"].eq("库存24个月以上"), "_物料键"].nunique(dropna=True)),
            "超目标库存数量": int(detail.loc[detail["库存区间"].eq("超目标库存"), "_物料键"].nunique(dropna=True)),
            "正常及关注库存数量": int(detail.loc[detail["库存区间"].eq("正常及关注库存"), "_物料键"].nunique(dropna=True)),
            "数据问题数量": int(detail.loc[detail["库存区间"].eq("数据问题"), "_物料键"].nunique(dropna=True)),
            "有效明细金额": float(detail_amount) if pd.notna(detail_amount) else np.nan,
            "汇总金额核对一致": amounts_match,
            **trade_diagnostics,
        }
        readable_mapping = {name: self._display_column(column) for name, column in field_map.items()}
        readable_mapping["价格表物料代码"] = self._display_column(price_material_column)
        readable_mapping["价格表单价"] = self._display_column(price_column)
        return InventoryAnalysisResult(detail[DETAIL_COLUMNS].copy(), views, summaries, readable_mapping, diagnostics)

    @staticmethod
    def _sorted(frame: pd.DataFrame) -> pd.DataFrame:
        return frame.sort_values("可用库存存销比", ascending=False, na_position="last", kind="stable").reset_index(drop=True)

    @staticmethod
    def _summarize(detail: pd.DataFrame, dimensions: list[str]) -> pd.DataFrame:
        source = detail.copy()
        for dimension in dimensions:
            source[dimension] = source[dimension].fillna("未分类").replace("", "未分类")
        grouped = (
            source.groupby(dimensions, dropna=False, sort=False)
            .agg(
                可用库存金额=("可用库存金额", lambda values: values.sum(min_count=1)),
                型号数量=("_物料键", lambda values: values.nunique(dropna=True)),
            )
            .reset_index()
            .sort_values("可用库存金额", ascending=False, na_position="last", kind="stable")
            .reset_index(drop=True)
        )
        grouped["可用库存金额"] = grouped["可用库存金额"] / 1_000_000
        grouped = grouped.rename(columns={"可用库存金额": SUMMARY_AMOUNT_COLUMN})
        total = {dimension: "合计" if index == 0 else "" for index, dimension in enumerate(dimensions)}
        detail_total = detail["可用库存金额"].sum(min_count=1)
        total_amount_musd = detail_total / 1_000_000 if pd.notna(detail_total) else np.nan
        total[SUMMARY_AMOUNT_COLUMN] = total_amount_musd
        if pd.notna(total_amount_musd) and total_amount_musd != 0:
            grouped[SUMMARY_SHARE_COLUMN] = grouped[SUMMARY_AMOUNT_COLUMN] / total_amount_musd
            total[SUMMARY_SHARE_COLUMN] = 1.0
        else:
            grouped[SUMMARY_SHARE_COLUMN] = np.nan
            total[SUMMARY_SHARE_COLUMN] = np.nan
        total["型号数量"] = detail["_物料键"].nunique(dropna=True)
        grouped = grouped[dimensions + [SUMMARY_AMOUNT_COLUMN, SUMMARY_SHARE_COLUMN, "型号数量"]]
        return pd.concat([grouped, pd.DataFrame([total])], ignore_index=True)

    @staticmethod
    def _build_trade_category_tables(
        summary: pd.DataFrame,
    ) -> tuple[dict[str, pd.DataFrame], dict[str, int]]:
        total_amount = pd.to_numeric(summary.iloc[-1][SUMMARY_AMOUNT_COLUMN], errors="coerce")
        total_models = pd.to_numeric(summary.iloc[-1]["型号数量"], errors="coerce")
        detail = summary.iloc[:-1].copy()
        detail[SUMMARY_AMOUNT_COLUMN] = pd.to_numeric(detail[SUMMARY_AMOUNT_COLUMN], errors="coerce")
        detail[SUMMARY_SHARE_COLUMN] = pd.to_numeric(detail[SUMMARY_SHARE_COLUMN], errors="coerce")
        detail["型号数量"] = pd.to_numeric(detail["型号数量"], errors="coerce")
        detail = detail.sort_values(
            SUMMARY_AMOUNT_COLUMN, ascending=False, na_position="last", kind="stable"
        ).reset_index(drop=True)

        positive = detail[SUMMARY_AMOUNT_COLUMN].gt(0)
        detail["排名"] = pd.Series(pd.NA, index=detail.index, dtype="Int64")
        detail.loc[positive, "排名"] = range(1, int(positive.sum()) + 1)
        detail["累计占比"] = np.nan
        detail.loc[positive, "累计占比"] = (
            detail.loc[positive, SUMMARY_SHARE_COLUMN].fillna(0).cumsum()
        )
        detail["单型号库存金额（k USD）"] = np.where(
            detail["型号数量"].gt(0),
            detail[SUMMARY_AMOUNT_COLUMN] * 1000 / detail["型号数量"],
            np.nan,
        )
        detail["库存分层分类"] = np.select(
            [
                detail[SUMMARY_AMOUNT_COLUMN].ge(1),
                detail[SUMMARY_AMOUNT_COLUMN].ge(0.5),
                detail[SUMMARY_AMOUNT_COLUMN].ge(0.1),
                detail[SUMMARY_AMOUNT_COLUMN].ge(0),
                detail[SUMMARY_AMOUNT_COLUMN].lt(0),
            ],
            [
                "核心（≥1 M USD）",
                "重点（0.5–1 M USD）",
                "一般（0.1–0.5 M USD）",
                "长尾（0–0.1 M USD）",
                "负金额",
            ],
            default=None,
        )
        detail["数据状态"] = np.select(
            [
                detail[SUMMARY_AMOUNT_COLUMN].isna(),
                detail[SUMMARY_AMOUNT_COLUMN].lt(0),
                detail[SUMMARY_AMOUNT_COLUMN].ge(0) & detail[SUMMARY_AMOUNT_COLUMN].lt(0.005),
            ],
            ["金额缺失", "负库存金额", "金额接近零"],
            default="正常",
        )

        detail_columns = [
            "排名",
            "国贸产品分类码",
            SUMMARY_AMOUNT_COLUMN,
            SUMMARY_SHARE_COLUMN,
            "累计占比",
            "型号数量",
            "单型号库存金额（k USD）",
            "库存分层分类",
            "数据状态",
        ]
        full_detail = detail[detail_columns].copy()

        tiers = [
            "核心（≥1 M USD）",
            "重点（0.5–1 M USD）",
            "一般（0.1–0.5 M USD）",
            "长尾（0–0.1 M USD）",
            "负金额",
        ]
        overview_rows: list[dict[str, object]] = []
        for tier in tiers:
            subset = detail.loc[detail["库存分层分类"].eq(tier)]
            amount = subset[SUMMARY_AMOUNT_COLUMN].sum(min_count=1)
            models = subset["型号数量"].sum(min_count=1)
            overview_rows.append(
                {
                    "库存分层分类": tier,
                    "分类码数量": int(len(subset)),
                    "型号数量": models,
                    SUMMARY_AMOUNT_COLUMN: amount,
                    SUMMARY_SHARE_COLUMN: amount / total_amount
                    if pd.notna(amount) and pd.notna(total_amount) and total_amount != 0
                    else np.nan,
                    "单型号平均金额（k USD）": amount * 1000 / models
                    if pd.notna(amount) and pd.notna(models) and models != 0
                    else np.nan,
                }
            )
        overview_rows.append(
            {
                "库存分层分类": "合计",
                "分类码数量": int(len(detail)),
                "型号数量": total_models,
                SUMMARY_AMOUNT_COLUMN: total_amount,
                SUMMARY_SHARE_COLUMN: 1.0 if pd.notna(total_amount) else np.nan,
                "单型号平均金额（k USD）": total_amount * 1000 / total_models
                if pd.notna(total_amount) and pd.notna(total_models) and total_models != 0
                else np.nan,
            }
        )
        overview = pd.DataFrame(overview_rows)

        positive_cumulative = full_detail.loc[
            full_detail[SUMMARY_AMOUNT_COLUMN].gt(0), "累计占比"
        ]
        reaches_eighty = positive_cumulative.ge(0.8)
        coverage_count = (
            int(np.flatnonzero(reaches_eighty.to_numpy())[0] + 1)
            if reaches_eighty.any()
            else int(len(positive_cumulative))
        )
        return (
            {
                "国贸分类码分层": overview,
                "按国贸分类码汇总": full_detail,
            },
            {
                "国贸重点分类码数量": int(full_detail[SUMMARY_AMOUNT_COLUMN].ge(0.5).sum()),
                "国贸数据问题数量": int(full_detail["数据状态"].ne("正常").sum()),
                "国贸分类码80%覆盖数量": coverage_count,
            },
        )

    def export_excel(self, result: InventoryAnalysisResult) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook = writer.book
            header_format = workbook.add_format(
                {"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "align": "center"}
            )
            amount_format = workbook.add_format({"num_format": "#,##0.00"})
            quantity_format = workbook.add_format({"num_format": "#,##0"})
            ratio_format = workbook.add_format({"num_format": "0.00"})
            percent_format = workbook.add_format({"num_format": "0.00%"})
            warning_format = workbook.add_format({"bg_color": "#FFF2CC"})
            negative_format = workbook.add_format({"bg_color": "#FCE8E6", "font_color": "#9C0006"})

            frames = {**result.views, **result.summaries}
            for sheet_name in SHEET_NAMES:
                frame = frames[sheet_name].copy()
                if sheet_name in {"国贸分类码分层", "按国贸分类码汇总"}:
                    if "库存分层分类" not in frame.columns and "金额层级" in frame.columns:
                        frame = frame.rename(columns={"金额层级": "库存分层分类"})
                    if sheet_name == "国贸分类码分层" and "库存分层分类" in frame.columns:
                        valid_tiers = {
                            "核心（≥1 M USD）",
                            "重点（0.5–1 M USD）",
                            "一般（0.1–0.5 M USD）",
                            "长尾（0–0.1 M USD）",
                            "负金额",
                            "合计",
                        }
                        frame = frame.loc[frame["库存分层分类"].isin(valid_tiers)].reset_index(drop=True)
                if sheet_name in result.summaries and SUMMARY_AMOUNT_COLUMN not in frame.columns and "可用库存金额" in frame.columns:
                    frame["可用库存金额"] = frame["可用库存金额"] / 1_000_000
                    frame = frame.rename(columns={"可用库存金额": SUMMARY_AMOUNT_COLUMN})
                frame.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
                worksheet = writer.sheets[sheet_name]
                worksheet.hide_gridlines(2)
                worksheet.freeze_panes(1, 0)
                worksheet.autofilter(0, 0, max(len(frame), 1), len(frame.columns) - 1)
                for column_index, column in enumerate(frame.columns):
                    worksheet.write(0, column_index, column, header_format)
                    values = frame[column].astype(str).replace("nan", "")
                    width = min(max(len(str(column)) * 2, int(values.map(len).quantile(0.95) if len(values) else 0) + 2), 36)
                    cell_format = None
                    if column in {"可用库存存销比", "全供应链存销比", "目标库存月数", "有效预测月销", "单价"}:
                        cell_format = ratio_format
                    if column in {"当前可用库存", "在途+在产", "型号数量", "分类码数量", "排名"}:
                        cell_format = quantity_format
                    if column in {
                        "可用库存金额",
                        SUMMARY_AMOUNT_COLUMN,
                        "单型号库存金额（k USD）",
                        "单型号平均金额（k USD）",
                    }:
                        cell_format = amount_format
                    if column in {SUMMARY_SHARE_COLUMN, "累计占比"}:
                        cell_format = percent_format
                    worksheet.set_column(column_index, column_index, width, cell_format)

                if "计算状态" in frame.columns and len(frame):
                    status_index = frame.columns.get_loc("计算状态")
                    first_data_row, last_data_row = 1, len(frame)
                    status_letter = self._excel_column(status_index)
                    worksheet.conditional_format(
                        first_data_row,
                        0,
                        last_data_row,
                        len(frame.columns) - 1,
                        {"type": "formula", "criteria": f'=${status_letter}2<>"正常"', "format": warning_format},
                    )

                if "数据状态" in frame.columns and len(frame):
                    status_index = frame.columns.get_loc("数据状态")
                    status_letter = self._excel_column(status_index)
                    worksheet.conditional_format(
                        1,
                        0,
                        len(frame),
                        len(frame.columns) - 1,
                        {"type": "formula", "criteria": f'=${status_letter}2<>"正常"', "format": warning_format},
                    )
                    worksheet.conditional_format(
                        1,
                        0,
                        len(frame),
                        len(frame.columns) - 1,
                        {"type": "formula", "criteria": f'=${status_letter}2="负库存金额"', "format": negative_format},
                    )

                if sheet_name in {"按产品状态汇总", "按产品组汇总"} and len(frame) > 1:
                    category_column = 1 if sheet_name == "按产品组汇总" else 0
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    chart = workbook.add_chart({"type": "column"})
                    chart.add_series(
                        {
                            "name": SUMMARY_AMOUNT_COLUMN,
                            "categories": [sheet_name, 1, category_column, len(frame) - 1, category_column],
                            "values": [sheet_name, 1, amount_column, len(frame) - 1, amount_column],
                            "fill": {"color": "#5B9BD5"},
                        }
                    )
                    chart.set_title({"name": sheet_name})
                    chart.set_y_axis({"name": "M USD", "num_format": "#,##0.00"})
                    chart.set_legend({"none": True})
                    worksheet.insert_chart(1, len(frame.columns) + 1, chart, {"x_scale": 1.35, "y_scale": 1.2})

                    total_amount = frame.iloc[-1][SUMMARY_AMOUNT_COLUMN]
                    total_text = "—" if pd.isna(total_amount) else f"{total_amount:,.2f} M USD"
                    if sheet_name == "按产品组汇总":
                        chart_rows = frame.iloc[:-1].copy()
                        custom_labels: list[dict[str, object]] = []
                        for _, chart_row in chart_rows.iterrows():
                            amount = pd.to_numeric(chart_row[SUMMARY_AMOUNT_COLUMN], errors="coerce")
                            share = pd.to_numeric(chart_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                            if pd.notna(amount) and amount >= 1:
                                share_text = "—" if pd.isna(share) else f"{share:.1%}"
                                is_large_slice = pd.notna(share) and share >= 0.05
                                custom_labels.append(
                                    {
                                        "value": f"{chart_row.iloc[category_column]}\n{share_text}",
                                        "position": "center" if is_large_slice else "outside_end",
                                        "font": {"size": 9 if is_large_slice else 8},
                                    }
                                )
                            else:
                                custom_labels.append({"delete": True})

                        doughnut_chart = workbook.add_chart({"type": "doughnut"})
                        doughnut_chart.add_series(
                            {
                                "name": SUMMARY_AMOUNT_COLUMN,
                                "categories": [sheet_name, 1, category_column, len(frame) - 1, category_column],
                                "values": [sheet_name, 1, amount_column, len(frame) - 1, amount_column],
                                "data_labels": {
                                    "custom": custom_labels,
                                    "leader_lines": True,
                                },
                            }
                        )
                        doughnut_chart.set_hole_size(46)
                        doughnut_chart.set_title({"name": "按产品组汇总"})
                        doughnut_chart.set_legend({"none": True})
                        chart_start_column = len(frame.columns) + 1
                        worksheet.insert_chart(
                            20,
                            chart_start_column,
                            doughnut_chart,
                            {"x_scale": 1.65, "y_scale": 1.55},
                        )

                        unlabelled = chart_rows.loc[
                            pd.to_numeric(chart_rows[SUMMARY_AMOUNT_COLUMN], errors="coerce").lt(1)
                        ].copy()
                        above_one_percent = unlabelled.loc[
                            pd.to_numeric(unlabelled[SUMMARY_SHARE_COLUMN], errors="coerce").ge(0.01)
                        ]
                        below_one_percent = unlabelled.loc[
                            pd.to_numeric(unlabelled[SUMMARY_SHARE_COLUMN], errors="coerce").lt(0.01)
                        ]

                        def joined_names(rows: pd.DataFrame) -> str:
                            names = rows.iloc[:, category_column].dropna().astype(str).tolist()
                            return "、".join(names) if names else "无"

                        legend_text = (
                            f"总可用库存金额\n{total_text}\n\n"
                            "低于 1 M USD（图中未标出数值）\n\n"
                            f"占比1%以上：{joined_names(above_one_percent)}\n\n"
                            f"占比低于1%：{joined_names(below_one_percent)}"
                        )
                        worksheet.insert_textbox(
                            20,
                            chart_start_column + 14,
                            legend_text,
                            {
                                "width": 330,
                                "height": 330,
                                "font": {"size": 10, "color": "#44546A"},
                                "fill": {"color": "#FFFFFF", "transparency": 100},
                                "line": {"none": True},
                                "align": {"vertical": "middle"},
                            },
                        )
                    else:
                        pie_chart = workbook.add_chart({"type": "pie"})
                        pie_chart.add_series(
                            {
                                "name": SUMMARY_AMOUNT_COLUMN,
                                "categories": [sheet_name, 1, category_column, len(frame) - 1, category_column],
                                "values": [sheet_name, 1, amount_column, len(frame) - 1, amount_column],
                                "data_labels": {
                                    "category": True,
                                    "percentage": True,
                                    "leader_lines": True,
                                    "position": "best_fit",
                                    "font": {"size": 8},
                                },
                            }
                        )
                        pie_chart.set_title({"name": f"{sheet_name}｜总可用库存金额：{total_text}"})
                        pie_chart.set_legend({"none": True})
                        worksheet.insert_chart(
                            20,
                            len(frame.columns) + 1,
                            pie_chart,
                            {"x_scale": 1.35, "y_scale": 1.2},
                        )

                if sheet_name == "国贸分类码分层" and len(frame) > 1:
                    category_column = frame.columns.get_loc("库存分层分类")
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    model_count_column = frame.columns.get_loc("型号数量")
                    last_tier_row = len(frame) - 1
                    share_labels: list[dict[str, object]] = []
                    for _, tier_row in frame.iloc[:-1].iterrows():
                        amount = pd.to_numeric(tier_row[SUMMARY_AMOUNT_COLUMN], errors="coerce")
                        share = pd.to_numeric(tier_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                        if pd.isna(amount) or pd.isna(share):
                            share_labels.append({"delete": True})
                        else:
                            share_labels.append(
                                {"value": f"{share:.1%}", "position": "outside_end", "font": {"size": 9}}
                            )

                    pivot_chart = workbook.add_chart({"type": "column"})
                    pivot_chart.add_series(
                        {
                            "name": SUMMARY_AMOUNT_COLUMN,
                            "categories": [sheet_name, 1, category_column, last_tier_row, category_column],
                            "values": [sheet_name, 1, amount_column, last_tier_row, amount_column],
                            "fill": {"color": "#5B9BD5"},
                            "data_labels": {"custom": share_labels},
                        }
                    )
                    model_chart = workbook.add_chart({"type": "line"})
                    model_chart.add_series(
                        {
                            "name": "型号数量",
                            "categories": [sheet_name, 1, category_column, last_tier_row, category_column],
                            "values": [sheet_name, 1, model_count_column, last_tier_row, model_count_column],
                            "y2_axis": True,
                            "line": {"color": "#ED7D31", "width": 2},
                            "marker": {"type": "circle", "size": 5},
                        }
                    )
                    pivot_chart.combine(model_chart)
                    pivot_chart.set_title({"name": "库存分层透视图"})
                    pivot_chart.set_y_axis({"name": "库存金额（M USD）", "num_format": "0.00"})
                    pivot_chart.set_y2_axis({"name": "型号数量", "num_format": "#,##0"})
                    pivot_chart.set_legend({"position": "bottom"})
                    worksheet.insert_chart(
                        1,
                        len(frame.columns) + 1,
                        pivot_chart,
                        {"x_scale": 1.65, "y_scale": 1.35},
                    )

                if sheet_name == "按国贸分类码汇总" and len(frame) >= 1:
                    category_column = frame.columns.get_loc("国贸产品分类码")
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    cumulative_column = frame.columns.get_loc("累计占比")
                    cumulative = pd.to_numeric(frame["累计占比"], errors="coerce")
                    reaches_eighty = cumulative.ge(0.8)
                    cutoff = (
                        int(np.flatnonzero(reaches_eighty.to_numpy())[0] + 1)
                        if reaches_eighty.any()
                        else int(cumulative.notna().sum())
                    )
                    if cutoff:
                        pareto_chart = workbook.add_chart({"type": "column"})
                        pareto_chart.add_series(
                            {
                                "name": SUMMARY_AMOUNT_COLUMN,
                                "categories": [sheet_name, 1, category_column, cutoff, category_column],
                                "values": [sheet_name, 1, amount_column, cutoff, amount_column],
                                "fill": {"color": "#5B9BD5"},
                            }
                        )
                        cumulative_chart = workbook.add_chart({"type": "line"})
                        cumulative_chart.add_series(
                            {
                                "name": "累计占比",
                                "categories": [sheet_name, 1, category_column, cutoff, category_column],
                                "values": [sheet_name, 1, cumulative_column, cutoff, cumulative_column],
                                "y2_axis": True,
                                "line": {"color": "#ED7D31", "width": 2},
                                "marker": {"type": "circle", "size": 4},
                            }
                        )
                        pareto_chart.combine(cumulative_chart)
                        pareto_chart.set_title({"name": f"累计80%库存金额分类码（前{cutoff}项）"})
                        pareto_chart.set_x_axis({"name": "国贸产品分类码", "label_position": "low"})
                        pareto_chart.set_y_axis({"name": "可用库存金额（M USD）", "num_format": "0.00"})
                        pareto_chart.set_y2_axis(
                            {"name": "累计占比", "num_format": "0%", "min": 0, "max": 1}
                        )
                        pareto_chart.set_legend({"position": "bottom"})
                        worksheet.insert_chart(
                            1,
                            len(frame.columns) + 1,
                            pareto_chart,
                            {"x_scale": 2.1, "y_scale": 1.45},
                        )
        return output.getvalue()

    @staticmethod
    def _excel_column(zero_based_index: int) -> str:
        value = zero_based_index + 1
        letters = ""
        while value:
            value, remainder = divmod(value - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters
