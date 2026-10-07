from __future__ import annotations

import io
import re
from colorsys import hls_to_rgb
from dataclasses import dataclass
from typing import BinaryIO, Hashable

import numpy as np
import pandas as pd


DETAIL_COLUMNS = [
    "销售组织",
    "物料",
    "物料描述",
    "产品状态",
    "产品组",
    "产品组描述",
    "国贸产品分类码",
    "当前库存等级",
    "ABC等级",
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

TRADE_CATEGORY_FIELDS = [
    "国贸产品分类码",
    "国贸产品分类码L1",
    "国贸产品分类码L2",
    "国贸产品分类码L3",
    "国贸产品分类码L4",
]

STATUS_GROUPS = {
    "NEW RELEASE (RECOMMEND)": ("NEW RELEASE (RECOMMEND)", "NEW RELEASE", "NEW"),
    "NORMAL/PHASING OUT": ("NORMAL", "PHASING OUT", "NORMAL/PHASING OUT"),
    "EOL": ("EOL",),
}

STATUS_SHEET_PREFIXES = {
    "NEW RELEASE (RECOMMEND)": "新品推荐",
    "NORMAL/PHASING OUT": "正常渐退",
    "EOL": "停产",
}

SUMMARY_AMOUNT_COLUMN = "可用库存金额（M USD）"
SUMMARY_SHARE_COLUMN = "库存金额占比"


def composition_chart_color(category: str, index: int) -> str:
    """Use muted, repeatable colors in both web and Excel composition charts."""
    hue = (0.58 + index * 0.61803398875) % 1
    red, green, blue = hls_to_rgb(hue, 0.52, 0.40)
    return f"#{round(red * 255):02X}{round(green * 255):02X}{round(blue * 255):02X}"

LEGACY_SHEET_NAMES = [
    "库存24个月以上",
    "超目标库存",
    "正常及关注库存",
    "按产品状态汇总",
    "按产品组汇总",
    "国贸分类码分层",
    "按国贸分类码汇总",
]

SHEET_NAMES = [
    "库存24个月以上", "超目标库存", "正常及关注库存", "无销量", "无库存等级",
    "按产品状态汇总",
    "新品推荐L1", "新品推荐L2", "新品推荐L3", "新品推荐L4",
    "正常渐退L1", "正常渐退L2", "正常渐退L3", "正常渐退L4", "正常渐退ABC",
    "停产L1", "停产L2", "停产L3", "停产L4",
    "产品组分类明细",
]


@dataclass(frozen=True)
class InventoryAnalysisResult:
    detail: pd.DataFrame
    views: dict[str, pd.DataFrame]
    summaries: dict[str, pd.DataFrame]
    field_mapping: dict[str, str]
    diagnostics: dict[str, object]


class InventoryStructureAnalyzer:
    """Analyze available-stock structure without mutating the uploaded workbooks."""

    FIELD_ALIASES = {
        "销售组织": ["销售组织", "salesorganization", "salesorg"],
        "物料": ["物料", "物料代码", "material", "materialcode", "sku"],
        "物料描述": ["物料描述", "产品描述", "materialdescription", "description"],
        "产品状态": ["产品状态", "productstatus"],
        "产品组": ["产品组", "productgroup", "productgroupcode"],
        "产品组描述": ["产品组描述", "productgroupdescription", "productgroupdesc"],
        "国贸产品分类码": ["国贸产品分类码", "产品分类码", "internationalproductclassificationcode"],
        "国贸产品分类码L1": ["国贸产品分类码L1", "产品分类码L1"],
        "国贸产品分类码L2": ["国贸产品分类码L2", "产品分类码L2"],
        "国贸产品分类码L3": ["国贸产品分类码L3", "产品分类码L3"],
        "国贸产品分类码L4": ["国贸产品分类码L4", "产品分类码L4"],
        "级别": ["级别", "层级", "level", "hierarchylevel"],
        "当前库存等级": ["当前库存等级", "currentstocklevel", "currentstockgrade"],
        "ABC等级": ["泛欧ABC", "ABC等级", "ABC", "paneuropeabc"],
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
        optional_fields = {*TRADE_CATEGORY_FIELDS, "全供应链存销比", "销售组织", "ABC等级"}
        missing = [name for name in self.FIELD_ALIASES if name not in mapping and name not in optional_fields]
        if missing:
            raise ValueError("Replenishment 表缺少必要字段：" + "、".join(missing))
        if "国贸产品分类码" not in mapping and not any(
            name in mapping for name in TRADE_CATEGORY_FIELDS[1:]
        ):
            raise ValueError(
                "Replenishment 表缺少国贸分类字段：请提供“国贸产品分类码”，或至少提供国贸产品分类码L1-L4中的一列。"
            )
        return mapping

    @staticmethod
    def _trade_keys(dimension: str) -> tuple[str, str]:
        if dimension == "国贸产品分类码":
            return "国贸分类码分层", "按国贸分类码汇总"
        level = dimension.removeprefix("国贸产品分类码")
        if level in {"L2", "L3", "L4"}:
            return f"国贸分类码{level}前80%", f"国贸分类码{level}明细"
        return f"国贸分类码{level}分层", f"国贸分类码{level}明细"

    @staticmethod
    def _detail_columns(trade_dimensions: list[str]) -> list[str]:
        columns = DETAIL_COLUMNS.copy()
        category_index = columns.index("国贸产品分类码")
        columns[category_index : category_index + 1] = trade_dimensions
        return columns

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
        trade_dimensions = [name for name in TRADE_CATEGORY_FIELDS[1:] if name in field_map]
        if not trade_dimensions:
            trade_dimensions = ["国贸产品分类码"]
        detail_columns = self._detail_columns(trade_dimensions)
        if "销售组织" not in detail.columns:
            detail["销售组织"] = pd.NA
        if "ABC等级" not in detail.columns:
            detail["ABC等级"] = pd.NA
        if "全供应链存销比" not in detail.columns:
            detail["全供应链存销比"] = np.nan
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
        valid_full_chain_ratio = (
            detail["当前可用库存"].notna()
            & detail["在途+在产"].notna()
            & detail["有效预测月销"].gt(0)
        )
        detail["全供应链存销比"] = np.where(
            valid_full_chain_ratio,
            (detail["当前可用库存"] + detail["在途+在产"])
            / detail["有效预测月销"],
            np.nan,
        )
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

        detail = detail[detail_columns + ["_物料键"]]
        views = {
            "库存24个月以上": self._sorted(detail.loc[detail["库存区间"].eq("库存24个月以上")]),
            "超目标库存": self._sorted(detail.loc[detail["库存区间"].eq("超目标库存")]),
        }
        normal = detail.loc[detail["库存区间"].eq("正常及关注库存")]
        problems = detail.loc[detail["库存区间"].eq("数据问题")]
        views["正常及关注库存"] = pd.concat([self._sorted(normal), self._sorted(problems)], ignore_index=True)
        no_sales = detail["有效预测月销"].isna() | detail["有效预测月销"].le(0)
        no_stock_level = detail["当前库存等级"].isna() | detail["当前库存等级"].astype("string").str.strip().eq("").fillna(False)
        views["无销量"] = self._sorted(detail.loc[no_sales])
        views["无库存等级"] = self._sorted(detail.loc[no_stock_level])
        views = {name: frame[detail_columns].copy() for name, frame in views.items()}

        summaries = {
            "按产品状态汇总": self._summarize(detail, ["产品状态"]),
            "按产品组汇总": self._summarize(detail, ["产品组", "产品组描述"]),
        }
        trade_diagnostics_by_dimension: dict[str, dict[str, int]] = {}
        for dimension in trade_dimensions:
            overview_key, detail_key = self._trade_keys(dimension)
            trade_summary = self._summarize(detail, [dimension])
            if dimension == "国贸产品分类码L1":
                l1_summary = trade_summary.copy()
                l1_summary["单型号库存金额（k USD）"] = np.where(
                    pd.to_numeric(l1_summary["型号数量"], errors="coerce").gt(0),
                    pd.to_numeric(l1_summary[SUMMARY_AMOUNT_COLUMN], errors="coerce")
                    * 1000
                    / pd.to_numeric(l1_summary["型号数量"], errors="coerce"),
                    np.nan,
                )
                summaries["国贸分类码L1汇总"] = l1_summary
                l1_detail = l1_summary.iloc[:-1]
                trade_diagnostics_by_dimension[dimension] = {
                    "国贸分类码分组总数": int(len(l1_detail)),
                    "国贸数据问题数量": int(
                        pd.to_numeric(l1_detail[SUMMARY_AMOUNT_COLUMN], errors="coerce").isna().sum()
                    ),
                }
                continue
            if dimension in {
                "国贸产品分类码L2",
                "国贸产品分类码L3",
                "国贸产品分类码L4",
            }:
                trade_tables, trade_diagnostics = self._build_trade_top_eighty_tables(
                    trade_summary,
                    dimension,
                    overview_key,
                    detail_key,
                )
                summaries.update(trade_tables)
                trade_diagnostics_by_dimension[dimension] = trade_diagnostics
                continue
            trade_tables, trade_diagnostics = self._build_trade_category_tables(
                trade_summary,
                dimension,
                overview_key,
                detail_key,
            )
            summaries.update(trade_tables)
            trade_diagnostics_by_dimension[dimension] = trade_diagnostics
        detail_amount = detail["可用库存金额"].sum(min_count=1)
        summary_amount_musd = summaries["按产品状态汇总"].iloc[-1][SUMMARY_AMOUNT_COLUMN]
        summary_amount = summary_amount_musd * 1_000_000
        amounts_match = (pd.isna(detail_amount) and pd.isna(summary_amount)) or bool(np.isclose(detail_amount, summary_amount))
        inventory_bands = ["库存24个月以上", "超目标库存", "正常及关注库存", "数据问题"]
        band_amounts = {
            band: float(amount) if pd.notna(amount) else np.nan
            for band in inventory_bands
            for amount in [detail.loc[detail["库存区间"].eq(band), "可用库存金额"].sum(min_count=1)]
        }
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
            "库存区间金额": band_amounts,
            "有效明细金额": float(detail_amount) if pd.notna(detail_amount) else np.nan,
            "汇总金额核对一致": amounts_match,
            "国贸分类维度": tuple(trade_dimensions),
            "国贸分类诊断": trade_diagnostics_by_dimension,
        }
        if trade_dimensions == ["国贸产品分类码"]:
            diagnostics.update(trade_diagnostics_by_dimension["国贸产品分类码"])
        readable_mapping = {name: self._display_column(column) for name, column in field_map.items()}
        readable_mapping["价格表物料代码"] = self._display_column(price_material_column)
        readable_mapping["价格表单价"] = self._display_column(price_column)
        return InventoryAnalysisResult(detail[detail_columns].copy(), views, summaries, readable_mapping, diagnostics)

    @staticmethod
    def _sorted(frame: pd.DataFrame) -> pd.DataFrame:
        return frame.sort_values("可用库存存销比", ascending=False, na_position="last", kind="stable").reset_index(drop=True)

    @staticmethod
    def _status_rows(detail: pd.DataFrame, status: str) -> pd.DataFrame:
        normalized = detail["产品状态"].astype("string").str.upper().str.replace(r"[^A-Z0-9]+", "", regex=True)
        accepted = STATUS_GROUPS.get(status, (status,))
        targets = {re.sub(r"[^A-Z0-9]+", "", value.upper()) for value in accepted}
        return detail.loc[normalized.isin(targets)].copy()

    def summarize_status_trade(
        self, detail: pd.DataFrame, status: str, dimension: str,
        selected_groups: list[str] | None = None,
    ) -> pd.DataFrame:
        source = self._status_rows(detail, status)
        if dimension not in source.columns:
            return pd.DataFrame(columns=[dimension, SUMMARY_AMOUNT_COLUMN, SUMMARY_SHARE_COLUMN, "型号数量"])
        source["产品组描述"] = source["产品组描述"].astype("string").str.strip().fillna("").replace("", "未分类")
        if selected_groups is not None:
            source = source.loc[source["产品组描述"].isin(selected_groups)].copy()
        source[dimension] = source[dimension].astype("string").str.strip().fillna("").replace("", "未分类")
        source["_物料键"] = source["物料"].map(self._material_key)
        grouped = source.groupby(dimension, dropna=False, sort=False).agg(
            amount=("可用库存金额", lambda values: values.sum(min_count=1)),
            型号数量=("_物料键", lambda values: values.nunique(dropna=True)),
        ).reset_index()
        total = source["可用库存金额"].sum(min_count=1)
        grouped[SUMMARY_SHARE_COLUMN] = grouped["amount"] / total if pd.notna(total) and total != 0 else np.nan
        grouped[SUMMARY_AMOUNT_COLUMN] = grouped.pop("amount") / 1_000_000
        return grouped.sort_values(SUMMARY_AMOUNT_COLUMN, ascending=False, na_position="last", kind="stable").reset_index(drop=True)

    def summarize_status_inventory(
        self, detail: pd.DataFrame, status: str, selected_groups: list[str]
    ) -> pd.DataFrame:
        source = self._status_rows(detail, status)
        source["产品组描述"] = source["产品组描述"].astype("string").str.strip().fillna("").replace("", "未分类")
        source = source.loc[source["产品组描述"].isin(selected_groups)].copy()
        source["ABC等级"] = source["ABC等级"].astype("string").str.strip().str.upper()
        source = source.loc[source["ABC等级"].isin(["A", "B", "C"])]
        source = source.loc[source["库存区间"].isin(["库存24个月以上", "超目标库存", "正常及关注库存"])]
        source["_物料键"] = source["物料"].map(self._material_key)
        grouped = source.groupby(["产品组描述", "ABC等级", "库存区间"], sort=False).agg(
            amount=("可用库存金额", lambda values: values.sum(min_count=1)),
            型号数量=("_物料键", lambda values: values.nunique(dropna=True)),
        ).reset_index()
        grouped[SUMMARY_AMOUNT_COLUMN] = grouped.pop("amount") / 1_000_000
        return grouped

    def available_product_groups(self, detail: pd.DataFrame, status: str) -> list[str]:
        source = self._status_rows(detail, status)
        groups = source["产品组描述"].astype("string").str.strip().fillna("").replace("", "未分类")
        return sorted(groups.unique().tolist())

    def summarize_status_overview(self, detail: pd.DataFrame) -> pd.DataFrame:
        total_amount = detail["可用库存金额"].sum(min_count=1)
        rows: list[dict[str, object]] = []
        matched_indices: set[int] = set()
        for status in STATUS_GROUPS:
            source = self._status_rows(detail, status)
            matched_indices.update(source.index.tolist())
            amount = source["可用库存金额"].sum(min_count=1) if len(source) else 0.0
            rows.append({
                "产品状态类别": status,
                SUMMARY_AMOUNT_COLUMN: amount / 1_000_000 if pd.notna(amount) else np.nan,
                "型号数量": source["物料"].map(self._material_key).nunique(dropna=True),
            })
        other = detail.loc[~detail.index.isin(matched_indices)]
        if not other.empty:
            amount = other["可用库存金额"].sum(min_count=1)
            rows.append({
                "产品状态类别": "其他状态",
                SUMMARY_AMOUNT_COLUMN: amount / 1_000_000 if pd.notna(amount) else np.nan,
                "型号数量": other["物料"].map(self._material_key).nunique(dropna=True),
            })
        summary = pd.DataFrame(rows)
        summary[SUMMARY_SHARE_COLUMN] = (
            summary[SUMMARY_AMOUNT_COLUMN] / (total_amount / 1_000_000)
            if pd.notna(total_amount) and total_amount != 0 else np.nan
        )
        total = pd.DataFrame([{
            "产品状态类别": "合计",
            SUMMARY_AMOUNT_COLUMN: total_amount / 1_000_000 if pd.notna(total_amount) else np.nan,
            "型号数量": detail["物料"].map(self._material_key).nunique(dropna=True),
            SUMMARY_SHARE_COLUMN: 1.0 if pd.notna(total_amount) and total_amount != 0 else np.nan,
        }])
        return pd.concat([summary, total], ignore_index=True)[
            ["产品状态类别", SUMMARY_AMOUNT_COLUMN, SUMMARY_SHARE_COLUMN, "型号数量"]
        ]

    def summarize_status_group_trade(
        self, detail: pd.DataFrame, status: str, dimension: str
    ) -> pd.DataFrame:
        source = self._status_rows(detail, status)
        if dimension not in source.columns:
            return pd.DataFrame(columns=["产品组描述", dimension, SUMMARY_AMOUNT_COLUMN, "型号数量"])
        for column in ["产品组描述", dimension]:
            source[column] = source[column].astype("string").str.strip().fillna("").replace("", "未分类")
        source["_物料键"] = source["物料"].map(self._material_key)
        grouped = source.groupby(["产品组描述", dimension], dropna=False, sort=False).agg(
            amount=("可用库存金额", lambda values: values.sum(min_count=1)),
            型号数量=("_物料键", lambda values: values.nunique(dropna=True)),
        ).reset_index()
        grouped[SUMMARY_AMOUNT_COLUMN] = grouped.pop("amount") / 1_000_000
        return grouped.sort_values(
            SUMMARY_AMOUNT_COLUMN, ascending=False, na_position="last", kind="stable"
        ).reset_index(drop=True)

    def summarize_all_status_group_trade(self, detail: pd.DataFrame) -> pd.DataFrame:
        columns = ["产品状态类别", "产品组描述", "国贸分类码层级", "国贸产品分类码", "型号数量", SUMMARY_AMOUNT_COLUMN]
        frames = []
        for status in STATUS_GROUPS:
            for level in range(1, 5):
                dimension = f"国贸产品分类码L{level}"
                if dimension not in detail.columns:
                    if level == 1 and "国贸产品分类码" in detail.columns:
                        dimension = "国贸产品分类码"
                    else:
                        continue
                summary = self.summarize_status_group_trade(detail, status, dimension)
                if summary.empty:
                    continue
                summary = summary.rename(columns={dimension: "国贸产品分类码"})
                summary.insert(0, "产品状态类别", status)
                summary.insert(2, "国贸分类码层级", f"L{level}")
                frames.append(summary[columns])
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=columns)

    def summarize_excel_status_trade(
        self, detail: pd.DataFrame, status: str, dimension: str
    ) -> pd.DataFrame:
        columns = ["产品组", "产品组描述", dimension, SUMMARY_AMOUNT_COLUMN, "组内金额占比", "型号数量"]
        source = self._status_rows(detail, status)
        if dimension not in source.columns:
            return pd.DataFrame(columns=columns)
        for column in ["产品组", "产品组描述", dimension]:
            source[column] = source[column].astype("string").str.strip().fillna("").replace("", "未分类")
        source["_物料键"] = source["物料"].map(self._material_key)
        grouped = source.groupby(["产品组", "产品组描述", dimension], dropna=False, sort=False).agg(
            amount=("可用库存金额", lambda values: values.sum(min_count=1)),
            型号数量=("_物料键", lambda values: values.nunique(dropna=True)),
        ).reset_index()
        grouped[SUMMARY_AMOUNT_COLUMN] = grouped.pop("amount") / 1_000_000
        group_amount = grouped.groupby(["产品组", "产品组描述"], dropna=False)[SUMMARY_AMOUNT_COLUMN].transform(
            lambda values: values.sum(min_count=1)
        )
        grouped["组内金额占比"] = np.where(
            group_amount.notna() & group_amount.ne(0), grouped[SUMMARY_AMOUNT_COLUMN] / group_amount, np.nan
        )
        grouped["_产品组正金额"] = grouped.groupby(["产品组", "产品组描述"], dropna=False)[
            SUMMARY_AMOUNT_COLUMN
        ].transform(lambda values: values.clip(lower=0).sum())
        grouped["_编号数值"] = pd.to_numeric(grouped["产品组"], errors="coerce")
        return grouped.sort_values(
            ["_产品组正金额", "_编号数值", "产品组", SUMMARY_AMOUNT_COLUMN],
            ascending=[False, True, True, False], na_position="last", kind="stable",
        )[columns].reset_index(drop=True)

    @classmethod
    def build_status_group_composition(
        cls, group_summary: pd.DataFrame, dimension: str
    ) -> pd.DataFrame:
        columns = ["产品组", "产品组描述"]
        positive = group_summary.loc[group_summary[SUMMARY_AMOUNT_COLUMN].gt(0)].copy()
        if positive.empty:
            return pd.DataFrame(columns=columns)
        category_amounts = positive.groupby(dimension, sort=False)[SUMMARY_AMOUNT_COLUMN].sum()
        ranked_categories = category_amounts.sort_values(ascending=False, kind="stable").index.tolist()
        pivot = positive.pivot_table(
            index=columns, columns=dimension, values=SUMMARY_AMOUNT_COLUMN,
            aggfunc="sum", fill_value=0, sort=False,
        ).reset_index()
        category_columns = [name for name in ranked_categories if name in pivot.columns]
        pivot["_产品组正金额"] = pivot[category_columns].sum(axis=1)
        pivot["_编号数值"] = pd.to_numeric(pivot["产品组"], errors="coerce")
        return pivot.sort_values(
            ["_产品组正金额", "_编号数值", "产品组"],
            ascending=[False, True, True], na_position="last", kind="stable",
        )[columns + category_columns].reset_index(drop=True)

    def summarize_excel_status_inventory(self, detail: pd.DataFrame) -> pd.DataFrame:
        """Keep product-group codes in the exported ABC table without changing web summaries."""
        source = self._status_rows(detail, "NORMAL/PHASING OUT")
        for column in ("产品组", "产品组描述"):
            source[column] = source[column].astype("string").str.strip().fillna("").replace("", "未分类")
        source["ABC等级"] = source["ABC等级"].astype("string").str.strip().str.upper()
        source = source.loc[
            source["ABC等级"].isin(["A", "B", "C"])
            & source["库存区间"].isin(["库存24个月以上", "超目标库存", "正常及关注库存"])
        ].copy()
        source["_物料键"] = source["物料"].map(self._material_key)
        grouped = source.groupby(["产品组", "产品组描述", "ABC等级", "库存区间"], sort=False).agg(
            amount=("可用库存金额", lambda values: values.sum(min_count=1)),
            型号数量=("_物料键", lambda values: values.nunique(dropna=True)),
        ).reset_index()
        grouped[SUMMARY_AMOUNT_COLUMN] = grouped.pop("amount") / 1_000_000
        grouped["_产品组正金额"] = grouped.groupby(
            ["ABC等级", "产品组", "产品组描述"], dropna=False
        )[SUMMARY_AMOUNT_COLUMN].transform(lambda values: values.clip(lower=0).sum())
        grouped["_编号数值"] = pd.to_numeric(grouped["产品组"], errors="coerce")
        band_order = {"库存24个月以上": 0, "超目标库存": 1, "正常及关注库存": 2}
        grouped["_库存区间顺序"] = grouped["库存区间"].map(band_order)
        return grouped.sort_values(
            ["ABC等级", "_产品组正金额", "_编号数值", "产品组", "_库存区间顺序"],
            ascending=[True, False, True, True, True], na_position="last", kind="stable",
        ).drop(columns=["_产品组正金额", "_编号数值", "_库存区间顺序"]).reset_index(drop=True)

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
        dimension: str,
        overview_key: str,
        detail_key: str,
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
            dimension,
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
                overview_key: overview,
                detail_key: full_detail,
            },
            {
                "国贸重点分类码数量": int(full_detail[SUMMARY_AMOUNT_COLUMN].ge(0.5).sum()),
                "国贸数据问题数量": int(full_detail["数据状态"].ne("正常").sum()),
                "国贸分类码80%覆盖数量": coverage_count,
            },
        )

    @staticmethod
    def _build_trade_top_eighty_tables(
        summary: pd.DataFrame,
        dimension: str,
        overview_key: str,
        detail_key: str,
    ) -> tuple[dict[str, pd.DataFrame], dict[str, int]]:
        detail = summary.iloc[:-1].copy()
        detail[SUMMARY_AMOUNT_COLUMN] = pd.to_numeric(
            detail[SUMMARY_AMOUNT_COLUMN], errors="coerce"
        )
        detail[SUMMARY_SHARE_COLUMN] = pd.to_numeric(
            detail[SUMMARY_SHARE_COLUMN], errors="coerce"
        )
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
            dimension,
            SUMMARY_AMOUNT_COLUMN,
            SUMMARY_SHARE_COLUMN,
            "累计占比",
            "型号数量",
            "单型号库存金额（k USD）",
            "数据状态",
        ]
        full_detail = detail[detail_columns].copy()
        positive_cumulative = full_detail.loc[
            full_detail[SUMMARY_AMOUNT_COLUMN].gt(0), "累计占比"
        ]
        reaches_eighty = positive_cumulative.ge(0.8)
        coverage_count = (
            int(np.flatnonzero(reaches_eighty.to_numpy())[0] + 1)
            if reaches_eighty.any()
            else int(len(positive_cumulative))
        )
        top_eighty = full_detail.loc[
            full_detail[SUMMARY_AMOUNT_COLUMN].gt(0)
        ].iloc[:coverage_count].reset_index(drop=True)

        return (
            {overview_key: top_eighty, detail_key: full_detail},
            {
                "国贸分类码分组总数": int(len(full_detail)),
                "国贸数据问题数量": int(full_detail["数据状态"].ne("正常").sum()),
                "国贸分类码80%覆盖数量": coverage_count,
            },
        )

    @staticmethod
    def summarize_inventory_bands(result: InventoryAnalysisResult) -> pd.DataFrame:
        diagnostics = result.diagnostics
        total = diagnostics["有效明细金额"]
        rows = []
        for band in ["库存24个月以上", "超目标库存", "正常及关注库存", "数据问题"]:
            amount = diagnostics["库存区间金额"][band]
            rows.append({
                "库存区间": band,
                "物料数量": diagnostics[f"{band}数量"],
                SUMMARY_AMOUNT_COLUMN: amount / 1_000_000 if pd.notna(amount) else np.nan,
                SUMMARY_SHARE_COLUMN: amount / total if pd.notna(amount) and pd.notna(total) and total != 0 else np.nan,
            })
        rows.append({
            "库存区间": "合计",
            "物料数量": diagnostics["记录数"],
            SUMMARY_AMOUNT_COLUMN: total / 1_000_000 if pd.notna(total) else np.nan,
            SUMMARY_SHARE_COLUMN: 1.0 if pd.notna(total) and total != 0 else np.nan,
        })
        return pd.DataFrame(rows)

    def export_excel(self, result: InventoryAnalysisResult) -> bytes:
        """Export the inventory page's detail and status based summary views."""
        output = io.BytesIO()
        detail_sheets = ["库存24个月以上", "超目标库存", "正常及关注库存", "无销量", "无库存等级"]
        status_sheet = "按产品状态汇总"
        grade_sheet = "正常渐退ABC"
        status_level_sheets = {
            f"{prefix}L{level}": (status, f"国贸产品分类码L{level}")
            for status, prefix in STATUS_SHEET_PREFIXES.items()
            for level in range(1, 5)
        }
        frames: dict[str, pd.DataFrame] = {name: result.views[name].copy() for name in detail_sheets}
        frames[status_sheet] = self.summarize_status_overview(result.detail)
        for sheet_name, (status, dimension) in status_level_sheets.items():
            if dimension in result.detail.columns:
                frames[sheet_name] = self.summarize_excel_status_trade(result.detail, status, dimension)
            elif dimension.endswith("L1") and "国贸产品分类码" in result.detail.columns:
                status_level_sheets[sheet_name] = (status, "国贸产品分类码")
                frames[sheet_name] = self.summarize_excel_status_trade(result.detail, status, "国贸产品分类码")
            else:
                frames[sheet_name] = pd.DataFrame({"提示": [f"源文件未提供{dimension}"]})
        frames[grade_sheet] = self.summarize_excel_status_inventory(result.detail)
        group_detail_sheet = "产品组分类明细"
        frames[group_detail_sheet] = self.summarize_all_status_group_trade(result.detail)
        sheet_names = [*detail_sheets, status_sheet]
        for status, prefix in STATUS_SHEET_PREFIXES.items():
            sheet_names.extend(f"{prefix}L{level}" for level in range(1, 5))
            if status == "NORMAL/PHASING OUT":
                sheet_names.append(grade_sheet)
        sheet_names.append(group_detail_sheet)

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook = writer.book
            header_format = workbook.add_format({
                "bold": True, "font_color": "white", "bg_color": "#1F4E78",
                "border": 1, "align": "center",
            })
            amount_format = workbook.add_format({"num_format": "#,##0.00"})
            quantity_format = workbook.add_format({"num_format": "#,##0"})
            ratio_format = workbook.add_format({"num_format": "0.00"})
            percent_format = workbook.add_format({"num_format": "0.00%"})
            warning_format = workbook.add_format({"bg_color": "#FFF2CC"})
            note_format = workbook.add_format({"font_color": "#44546A", "text_wrap": True})
            chart_scale = 2 / 3
            for sheet_name in sheet_names:
                frame = frames[sheet_name]
                header_row = 1 if sheet_name == group_detail_sheet else 0
                frame.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_row + 1, header=False)
                worksheet = writer.sheets[sheet_name]
                worksheet.hide_gridlines(2)
                worksheet.freeze_panes(header_row + 1, 0)
                worksheet.autofilter(header_row, 0, header_row + max(len(frame), 1), len(frame.columns) - 1)
                if sheet_name == group_detail_sheet:
                    worksheet.merge_range(
                        0, 0, 0, len(frame.columns) - 1,
                        "每行汇总一个产品状态类别、产品组和国贸分类码层级下的一个分类码。型号数量为去重物料数；金额单位为 M USD。L1–L4 是不同口径，不要跨层级相加。",
                        note_format,
                    )
                    worksheet.set_row(0, 34)
                for column_index, column in enumerate(frame.columns):
                    worksheet.write(header_row, column_index, column, header_format)
                    values = frame[column].astype(str).replace("nan", "")
                    width = min(max(len(str(column)) * 2, int(values.map(len).quantile(0.95) if len(values) else 0) + 2), 38)
                    column_format = None
                    if column in {"可用库存存销比", "全供应链存销比", "目标库存月数", "有效预测月销", "单价"}:
                        column_format = ratio_format
                    elif column in {"当前可用库存", "在途+在产", "型号数量", "物料数量", "分类码数量", "排名"}:
                        column_format = quantity_format
                    elif column in {"可用库存金额", SUMMARY_AMOUNT_COLUMN}:
                        column_format = amount_format
                    elif column in {SUMMARY_SHARE_COLUMN, "组内金额占比"}:
                        column_format = percent_format
                    worksheet.set_column(column_index, column_index, width, column_format)

                if sheet_name != status_sheet and SUMMARY_AMOUNT_COLUMN in frame.columns:
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    for row_index, amount in enumerate(frame[SUMMARY_AMOUNT_COLUMN], start=header_row + 1):
                        if pd.notna(amount):
                            worksheet.write_number(row_index, amount_column, float(amount), amount_format)

                if "计算状态" in frame.columns and len(frame):
                    status_letter = self._excel_column(frame.columns.get_loc("计算状态"))
                    worksheet.conditional_format(
                        1, 0, len(frame), len(frame.columns) - 1,
                        {"type": "formula", "criteria": f'=${status_letter}2<>"正常"', "format": warning_format},
                    )

                if sheet_name == status_sheet:
                    positive = frame.iloc[:-1].loc[frame.iloc[:-1][SUMMARY_AMOUNT_COLUMN].gt(0)]
                    helper_column = len(frame.columns) + 2
                    worksheet.write(0, helper_column, "产品状态类别", header_format)
                    worksheet.write(0, helper_column + 1, SUMMARY_AMOUNT_COLUMN, header_format)
                    for row_index, (_, row) in enumerate(positive.iterrows(), start=1):
                        worksheet.write(row_index, helper_column, row["产品状态类别"])
                        worksheet.write_number(row_index, helper_column + 1, row[SUMMARY_AMOUNT_COLUMN], amount_format)
                    if len(positive):
                        pie = workbook.add_chart({"type": "pie"})
                        pie.add_series({
                            "name": SUMMARY_AMOUNT_COLUMN,
                            "categories": [sheet_name, 1, helper_column, len(positive), helper_column],
                            "values": [sheet_name, 1, helper_column + 1, len(positive), helper_column + 1],
                            "data_labels": {"category": True, "percentage": True, "position": "best_fit"},
                        })
                        pie.set_title({"name": "按产品状态概览"})
                        pie.set_legend({"position": "right"})
                        worksheet.insert_chart(1, helper_column + 3, pie, {
                            "x_scale": 1.5 * chart_scale, "y_scale": 1.25 * chart_scale,
                        })

                if sheet_name in status_level_sheets and SUMMARY_AMOUNT_COLUMN in frame.columns:
                    status, dimension = status_level_sheets[sheet_name]
                    worksheet.write(0, len(frame.columns) + 2, f"{status} · {dimension}：产品组内分类码构成")
                    if len(frame):
                        pivot = self.build_status_group_composition(frame, dimension)
                        if pivot.empty:
                            worksheet.write(2, len(frame.columns) + 2, "没有可绘图的正库存金额。")
                        else:
                            category_columns = list(pivot.columns[2:])
                            positive_total = pivot[category_columns].to_numpy(dtype=float).sum()
                            group_totals = pivot[category_columns].sum(axis=1)
                            pivot.insert(2, "图表产品组标签", [
                                f"{description}（{amount / positive_total:.1%}）"
                                for description, amount in zip(pivot["产品组描述"], group_totals)
                            ])
                            pivot_header_row = 0
                            first_pivot_row = pivot_header_row + 1
                            last_pivot_row = pivot_header_row + len(pivot)
                            helper_start_column = max(30, len(frame.columns) + 20)
                            for column_index, column in enumerate(pivot.columns):
                                worksheet.write(pivot_header_row, helper_start_column + column_index, column, header_format)
                            for row_index, row in enumerate(pivot.itertuples(index=False, name=None), start=first_pivot_row):
                                for column_index, value in enumerate(row):
                                    if column_index < 3:
                                        worksheet.write(row_index, helper_start_column + column_index, value)
                                    else:
                                        worksheet.write_number(row_index, helper_start_column + column_index, float(value), amount_format)
                            worksheet.set_column(
                                helper_start_column, helper_start_column + len(pivot.columns) - 1,
                                None, None, {"hidden": True},
                            )
                            chart = workbook.add_chart({"type": "bar", "subtype": "percent_stacked"})
                            chart.show_hidden_data()
                            for column_index in range(3, len(pivot.columns)):
                                category = str(pivot.columns[column_index])
                                chart.add_series({
                                    "name": [sheet_name, pivot_header_row, helper_start_column + column_index],
                                    "categories": [sheet_name, first_pivot_row, helper_start_column + 2, last_pivot_row, helper_start_column + 2],
                                    "values": [sheet_name, first_pivot_row, helper_start_column + column_index, last_pivot_row, helper_start_column + column_index],
                                    "fill": {"color": composition_chart_color(category, column_index - 3)},
                                })
                            chart.set_title({"none": True})
                            chart.set_x_axis({"num_format": "0%"})
                            chart.set_y_axis({"reverse": True})
                            plot_height = int(max(340, 27 * len(pivot) + 100) * chart_scale)
                            legend_rows = (len(category_columns) + 1) // 2
                            legend_height = max(80, 22 * legend_rows + 20)
                            chart_height = plot_height + legend_height + 80
                            chart.set_plotarea({"layout": {
                                "x": 0.24, "y": 0.04, "width": 0.72,
                                "height": (plot_height - 30) / chart_height,
                            }})
                            chart.set_legend({
                                "position": "bottom", "font": {"size": 8},
                                "layout": {
                                    "x": 0.06, "y": (plot_height + 45) / chart_height,
                                    "width": 0.90, "height": legend_height / chart_height,
                                },
                            })
                            chart.set_size({
                                "width": int(960 * chart_scale),
                                "height": chart_height,
                            })
                            worksheet.insert_chart(2, len(frame.columns) + 2, chart)

                if sheet_name == grade_sheet and len(frame):
                    bands = ["库存24个月以上", "超目标库存", "正常及关注库存"]
                    colors = ["#D95F5F", "#F2B34D", "#4C9E91"]
                    helper_column = len(frame.columns) + 2
                    chart_column = helper_column + 6
                    start_row = 0
                    for grade in "ABC":
                        grade_data = frame.loc[frame["ABC等级"].eq(grade)]
                        groups = grade_data[["产品组", "产品组描述"]].drop_duplicates().itertuples(index=False, name=None)
                        groups = list(groups)
                        if not groups:
                            continue
                        grade_positive_total = grade_data[SUMMARY_AMOUNT_COLUMN].clip(lower=0).sum()
                        worksheet.write(start_row, helper_column, f"{grade}级产品组")
                        for offset, heading in enumerate(["产品组", "产品组描述（占比）", *bands]):
                            worksheet.write(start_row + 1, helper_column + offset, heading, header_format)
                        for group_index, (group_code, group_description) in enumerate(groups):
                            data_row = start_row + 2 + group_index
                            worksheet.write(data_row, helper_column, group_code)
                            group_rows = grade_data.loc[
                                grade_data["产品组"].eq(group_code)
                                & grade_data["产品组描述"].eq(group_description)
                            ]
                            group_positive = group_rows[SUMMARY_AMOUNT_COLUMN].clip(lower=0).sum()
                            group_share = group_positive / grade_positive_total if grade_positive_total > 0 else 0.0
                            worksheet.write(data_row, helper_column + 1, f"{group_description}（{group_share:.1%}）")
                            for band_index, band in enumerate(bands):
                                matching = grade_data.loc[
                                    grade_data["产品组"].eq(group_code)
                                    & grade_data["产品组描述"].eq(group_description)
                                    & grade_data["库存区间"].eq(band),
                                    SUMMARY_AMOUNT_COLUMN,
                                ]
                                value = matching.sum(min_count=1) if len(matching) else 0.0
                                if pd.notna(value):
                                    worksheet.write_number(data_row, helper_column + band_index + 2, float(value), amount_format)
                        chart = workbook.add_chart({"type": "bar", "subtype": "stacked"})
                        for band_index, band in enumerate(bands):
                            chart.add_series({
                                "name": band,
                                "categories": [sheet_name, start_row + 2, helper_column + 1, start_row + 1 + len(groups), helper_column + 1],
                                "values": [sheet_name, start_row + 2, helper_column + band_index + 2, start_row + 1 + len(groups), helper_column + band_index + 2],
                                "fill": {"color": colors[band_index]},
                            })
                        chart.set_title({"name": f"{grade}级产品组库存水位"})
                        chart.set_x_axis({"name": "可用库存金额（M USD）", "num_format": "#,##0.00"})
                        chart.set_y_axis({"name": "产品组", "reverse": True})
                        chart.set_legend({"position": "bottom"})
                        chart_height = int(max(420, 34 * len(groups) + 100) * chart_scale)
                        chart.set_size({"width": int(760 * chart_scale), "height": chart_height})
                        worksheet.insert_chart(start_row, chart_column, chart)
                        start_row += max(len(groups) + 2, (chart_height + 19) // 20) + 5
        return output.getvalue()

    def _export_legacy_excel(self, result: InventoryAnalysisResult) -> bytes:
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
            chart_scale = 2 / 3

            frames = {**result.views, **result.summaries}
            trade_dimensions = list(result.diagnostics.get("国贸分类维度", ("国贸产品分类码",)))
            trade_sheet_dimensions: dict[str, str] = {}
            sheet_names = LEGACY_SHEET_NAMES[:5].copy()
            for dimension in trade_dimensions:
                if dimension == "国贸产品分类码L1":
                    sheet_names.append("国贸分类码L1汇总")
                    continue
                overview_key, detail_key = self._trade_keys(dimension)
                if dimension in {
                    "国贸产品分类码L2",
                    "国贸产品分类码L3",
                    "国贸产品分类码L4",
                }:
                    sheet_names.append(detail_key)
                else:
                    sheet_names.extend([overview_key, detail_key])
                    trade_sheet_dimensions[overview_key] = dimension
                trade_sheet_dimensions[detail_key] = dimension

            for sheet_name in sheet_names:
                frame = frames[sheet_name].copy()
                if sheet_name in trade_sheet_dimensions:
                    if "库存分层分类" not in frame.columns and "金额层级" in frame.columns:
                        frame = frame.rename(columns={"金额层级": "库存分层分类"})
                    if sheet_name == self._trade_keys(trade_sheet_dimensions[sheet_name])[0] and "库存分层分类" in frame.columns:
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
                    worksheet.insert_chart(
                        1,
                        len(frame.columns) + 1,
                        chart,
                        {"x_scale": 1.35 * chart_scale, "y_scale": 1.2 * chart_scale},
                    )

                    total_amount = frame.iloc[-1][SUMMARY_AMOUNT_COLUMN]
                    total_text = "—" if pd.isna(total_amount) else f"{total_amount:,.2f} M USD"
                    if sheet_name == "按产品组汇总":
                        chart_rows = frame.iloc[:-1].copy()
                        custom_labels: list[dict[str, object]] = []
                        for _, chart_row in chart_rows.iterrows():
                            amount = pd.to_numeric(chart_row[SUMMARY_AMOUNT_COLUMN], errors="coerce")
                            share = pd.to_numeric(chart_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                            if pd.notna(amount) and amount > 0 and pd.notna(share):
                                share_text = "—" if pd.isna(share) else f"{share:.1%}"
                                is_large_slice = pd.notna(share) and share >= 0.05
                                custom_labels.append(
                                    {
                                        "value": f"{chart_row.iloc[category_column]}\n{share_text}",
                                        "position": "center",
                                        "font": {"size": 9},
                                    }
                                    if is_large_slice
                                    else {"delete": True}
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
                                },
                            }
                        )
                        doughnut_chart.set_hole_size(46)
                        doughnut_chart.set_title({"name": "按产品组汇总"})
                        doughnut_chart.set_legend({"position": "right"})
                        chart_start_column = len(frame.columns) + 1
                        worksheet.insert_chart(
                            20,
                            chart_start_column,
                            doughnut_chart,
                            {"x_scale": 1.65 * chart_scale, "y_scale": 1.55 * chart_scale},
                        )

                        worksheet.insert_textbox(
                            20,
                            chart_start_column + 10,
                            f"总可用库存金额\n{total_text}",
                            {
                                "width": 180,
                                "height": 75,
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
                            {"x_scale": 1.35 * chart_scale, "y_scale": 1.2 * chart_scale},
                        )

                if sheet_name == "国贸分类码L1汇总" and len(frame) > 1:
                    category_column = frame.columns.get_loc("国贸产品分类码L1")
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    last_category_row = len(frame) - 1
                    share_labels: list[dict[str, object]] = []
                    for _, category_row in frame.iloc[:-1].iterrows():
                        amount = pd.to_numeric(category_row[SUMMARY_AMOUNT_COLUMN], errors="coerce")
                        share = pd.to_numeric(category_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                        if pd.isna(amount) or pd.isna(share):
                            share_labels.append({"delete": True})
                        else:
                            share_labels.append(
                                {"value": f"{share:.1%}", "position": "outside_end", "font": {"size": 9}}
                            )

                    l1_bar_chart = workbook.add_chart({"type": "column"})
                    l1_bar_chart.add_series(
                        {
                            "name": SUMMARY_AMOUNT_COLUMN,
                            "categories": [sheet_name, 1, category_column, last_category_row, category_column],
                            "values": [sheet_name, 1, amount_column, last_category_row, amount_column],
                            "fill": {"color": "#2F75B5"},
                            "data_labels": {"custom": share_labels},
                        }
                    )
                    l1_bar_chart.set_title({"name": "国贸产品分类码L1库存金额"})
                    l1_bar_chart.set_x_axis({"name": "国贸产品分类码L1", "label_position": "low"})
                    l1_bar_chart.set_y_axis({"name": "库存金额（M USD）", "num_format": "0.00"})
                    l1_bar_chart.set_legend({"none": True})
                    worksheet.insert_chart(
                        1,
                        len(frame.columns) + 1,
                        l1_bar_chart,
                        {"x_scale": 1.55 * chart_scale, "y_scale": 1.3 * chart_scale},
                    )

                    l1_pie_chart = workbook.add_chart({"type": "pie"})
                    pie_labels: list[dict[str, object]] = []
                    for _, category_row in frame.iloc[:-1].iterrows():
                        category = str(category_row["国贸产品分类码L1"])
                        share = pd.to_numeric(category_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                        if pd.isna(share) or share <= 0:
                            pie_labels.append({"delete": True})
                        else:
                            pie_labels.append(
                                {
                                    "value": f"{category}\n{share:.1%}" if share >= 0.01 else f"{share:.1%}",
                                    "position": "center" if share >= 0.03 else "outside_end",
                                    "font": {"size": 9 if share >= 0.03 else 8},
                                }
                            )
                    l1_pie_chart.add_series(
                        {
                            "name": SUMMARY_AMOUNT_COLUMN,
                            "categories": [sheet_name, 1, category_column, last_category_row, category_column],
                            "values": [sheet_name, 1, amount_column, last_category_row, amount_column],
                            "data_labels": {
                                "custom": pie_labels,
                                "leader_lines": True,
                            },
                        }
                    )
                    l1_pie_chart.set_title({"name": "国贸产品分类码L1库存金额占比"})
                    l1_pie_chart.set_legend({"position": "right"})
                    worksheet.insert_chart(
                        20,
                        len(frame.columns) + 1,
                        l1_pie_chart,
                        {"x_scale": 1.55 * chart_scale, "y_scale": 1.35 * chart_scale},
                    )

                if (
                    sheet_name in trade_sheet_dimensions
                    and trade_sheet_dimensions[sheet_name] in {
                        "国贸产品分类码L2",
                        "国贸产品分类码L3",
                        "国贸产品分类码L4",
                    }
                    and sheet_name == self._trade_keys(trade_sheet_dimensions[sheet_name])[0]
                    and len(frame) >= 1
                ):
                    trade_dimension = trade_sheet_dimensions[sheet_name]
                    category_column = frame.columns.get_loc(trade_dimension)
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    model_count_column = frame.columns.get_loc("型号数量")
                    last_category_row = len(frame)
                    share_labels: list[dict[str, object]] = []
                    for _, category_row in frame.iterrows():
                        amount = pd.to_numeric(category_row[SUMMARY_AMOUNT_COLUMN], errors="coerce")
                        share = pd.to_numeric(category_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                        if pd.isna(amount) or pd.isna(share):
                            share_labels.append({"delete": True})
                        else:
                            share_labels.append(
                                {"value": f"{share:.1%}", "position": "outside_end", "font": {"size": 9}}
                            )

                    pivot_chart = workbook.add_chart({"type": "column"})
                    pivot_chart.add_series(
                        {
                            "name": "库存金额（柱形）",
                            "categories": [sheet_name, 1, category_column, last_category_row, category_column],
                            "values": [sheet_name, 1, amount_column, last_category_row, amount_column],
                            "fill": {"color": "#2F75B5"},
                            "data_labels": {"custom": share_labels},
                        }
                    )
                    model_chart = workbook.add_chart({"type": "line"})
                    model_chart.add_series(
                        {
                            "name": "型号数量（点）",
                            "categories": [sheet_name, 1, category_column, last_category_row, category_column],
                            "values": [sheet_name, 1, model_count_column, last_category_row, model_count_column],
                            "y2_axis": True,
                            "line": {"none": True},
                            "marker": {
                                "type": "diamond",
                                "size": 7,
                                "border": {"color": "#FFFFFF"},
                                "fill": {"color": "#ED7D31"},
                            },
                        }
                    )
                    pivot_chart.combine(model_chart)
                    pivot_chart.set_title({"name": f"{trade_dimension}累计前80%分类透视图"})
                    pivot_chart.set_x_axis({"name": trade_dimension, "label_position": "low"})
                    pivot_chart.set_y_axis({"name": "库存金额（M USD）", "num_format": "0.00"})
                    pivot_chart.set_y2_axis({"name": "型号数量", "num_format": "#,##0"})
                    pivot_chart.set_legend({"position": "bottom"})
                    worksheet.insert_chart(
                        1,
                        len(frame.columns) + 1,
                        pivot_chart,
                        {"x_scale": 1.75 * chart_scale, "y_scale": 1.4 * chart_scale},
                    )

                if (
                    sheet_name in trade_sheet_dimensions
                    and trade_sheet_dimensions[sheet_name] == "国贸产品分类码"
                    and sheet_name == self._trade_keys("国贸产品分类码")[0]
                    and len(frame) > 1
                ):
                    category_column = frame.columns.get_loc("库存分层分类")
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    model_count_column = frame.columns.get_loc("型号数量")
                    last_tier_row = len(frame) - 1
                    legacy_share_labels: list[dict[str, object]] = []
                    for _, tier_row in frame.iloc[:-1].iterrows():
                        share = pd.to_numeric(tier_row[SUMMARY_SHARE_COLUMN], errors="coerce")
                        legacy_share_labels.append(
                            {"delete": True}
                            if pd.isna(share)
                            else {
                                "value": f"{share:.1%}",
                                "position": "outside_end",
                                "font": {"size": 9},
                            }
                        )
                    legacy_chart = workbook.add_chart({"type": "column"})
                    legacy_chart.add_series(
                        {
                            "name": "库存金额（柱形）",
                            "categories": [sheet_name, 1, category_column, last_tier_row, category_column],
                            "values": [sheet_name, 1, amount_column, last_tier_row, amount_column],
                            "fill": {"color": "#2F75B5"},
                            "data_labels": {"custom": legacy_share_labels},
                        }
                    )
                    legacy_model_chart = workbook.add_chart({"type": "line"})
                    legacy_model_chart.add_series(
                        {
                            "name": "型号数量（点）",
                            "categories": [sheet_name, 1, category_column, last_tier_row, category_column],
                            "values": [sheet_name, 1, model_count_column, last_tier_row, model_count_column],
                            "y2_axis": True,
                            "line": {"none": True},
                            "marker": {
                                "type": "diamond",
                                "size": 7,
                                "fill": {"color": "#ED7D31"},
                            },
                        }
                    )
                    legacy_chart.combine(legacy_model_chart)
                    legacy_chart.set_title({"name": "国贸产品分类码库存分层透视图"})
                    legacy_chart.set_y_axis({"name": "库存金额（M USD）", "num_format": "0.00"})
                    legacy_chart.set_y2_axis({"name": "型号数量", "num_format": "#,##0"})
                    legacy_chart.set_legend({"position": "bottom"})
                    worksheet.insert_chart(
                        1,
                        len(frame.columns) + 1,
                        legacy_chart,
                        {"x_scale": 1.75 * chart_scale, "y_scale": 1.4 * chart_scale},
                    )

                if sheet_name in trade_sheet_dimensions and sheet_name == self._trade_keys(trade_sheet_dimensions[sheet_name])[1] and len(frame) >= 1:
                    trade_dimension = trade_sheet_dimensions[sheet_name]
                    category_column = frame.columns.get_loc(trade_dimension)
                    amount_column = frame.columns.get_loc(SUMMARY_AMOUNT_COLUMN)
                    model_count_column = frame.columns.get_loc("型号数量")
                    cumulative_column = frame.columns.get_loc("累计占比")
                    cumulative = pd.to_numeric(frame["累计占比"], errors="coerce")
                    reaches_eighty = cumulative.ge(0.8)
                    cutoff = (
                        int(np.flatnonzero(reaches_eighty.to_numpy())[0] + 1)
                        if reaches_eighty.any()
                        else int(cumulative.notna().sum())
                    )
                    if cutoff:
                        if trade_dimension in {
                            "国贸产品分类码L2",
                            "国贸产品分类码L3",
                            "国贸产品分类码L4",
                        }:
                            top_eighty_labels: list[dict[str, object]] = []
                            for _, category_row in frame.iloc[:cutoff].iterrows():
                                share = pd.to_numeric(
                                    category_row[SUMMARY_SHARE_COLUMN], errors="coerce"
                                )
                                top_eighty_labels.append(
                                    {"delete": True}
                                    if pd.isna(share)
                                    else {
                                        "value": f"{share:.1%}",
                                        "position": "outside_end",
                                        "font": {"size": 9},
                                    }
                                )
                            top_eighty_chart = workbook.add_chart({"type": "column"})
                            top_eighty_chart.add_series(
                                {
                                    "name": "库存金额（柱形）",
                                    "categories": [sheet_name, 1, category_column, cutoff, category_column],
                                    "values": [sheet_name, 1, amount_column, cutoff, amount_column],
                                    "fill": {"color": "#2F75B5"},
                                    "data_labels": {"custom": top_eighty_labels},
                                }
                            )
                            top_eighty_model_chart = workbook.add_chart({"type": "line"})
                            top_eighty_model_chart.add_series(
                                {
                                    "name": "型号数量（点）",
                                    "categories": [sheet_name, 1, category_column, cutoff, category_column],
                                    "values": [sheet_name, 1, model_count_column, cutoff, model_count_column],
                                    "y2_axis": True,
                                    "line": {"none": True},
                                    "marker": {
                                        "type": "diamond",
                                        "size": 7,
                                        "border": {"color": "#FFFFFF"},
                                        "fill": {"color": "#ED7D31"},
                                    },
                                }
                            )
                            top_eighty_chart.combine(top_eighty_model_chart)
                            top_eighty_chart.set_title(
                                {"name": f"{trade_dimension}累计前80%分类透视图"}
                            )
                            top_eighty_chart.set_x_axis(
                                {"name": trade_dimension, "label_position": "low"}
                            )
                            top_eighty_chart.set_y_axis(
                                {"name": "库存金额（M USD）", "num_format": "0.00"}
                            )
                            top_eighty_chart.set_y2_axis(
                                {"name": "型号数量", "num_format": "#,##0"}
                            )
                            top_eighty_chart.set_legend({"position": "bottom"})
                            worksheet.insert_chart(
                                1,
                                len(frame.columns) + 1,
                                top_eighty_chart,
                                {"x_scale": 2.1 * chart_scale, "y_scale": 1.45 * chart_scale},
                            )

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
                        pareto_chart.set_x_axis({"name": trade_dimension, "label_position": "low"})
                        pareto_chart.set_y_axis({"name": "可用库存金额（M USD）", "num_format": "0.00"})
                        pareto_chart.set_y2_axis(
                            {"name": "累计占比", "num_format": "0%", "min": 0, "max": 1}
                        )
                        pareto_chart.set_legend({"position": "bottom"})
                        worksheet.insert_chart(
                            24
                            if trade_dimension
                            in {
                                "国贸产品分类码L2",
                                "国贸产品分类码L3",
                                "国贸产品分类码L4",
                            }
                            else 1,
                            len(frame.columns) + 1,
                            pareto_chart,
                            {"x_scale": 2.1 * chart_scale, "y_scale": 1.45 * chart_scale},
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
