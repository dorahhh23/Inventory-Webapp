import io
import unittest
import zipfile

import numpy as np
import openpyxl
import pandas as pd

from inventory_structure_analysis import (
    DETAIL_COLUMNS,
    SHEET_NAMES,
    SUMMARY_AMOUNT_COLUMN,
    SUMMARY_SHARE_COLUMN,
    InventoryStructureAnalyzer,
)


REPLENISHMENT_COLUMNS = pd.MultiIndex.from_tuples(
    [
        ("物料", "物料"),
        ("物料", "物料描述"),
        ("产品基本信息", "产品状态"),
        ("产品基本信息", "产品组"),
        ("产品基本信息", "产品组描述"),
        ("产品基本信息", "国贸产品分类码"),
        ("产品基本信息", "级别"),
        ("产品基本信息", "Current Stock level"),
        ("产品基本信息", "库存月数"),
        ("库存信息", "当前可用库存"),
        ("在途在产信息", "在途+在产"),
        ("预测销量", "月预测"),
        ("预测销量", "月预测（修改）"),
        ("服务水平信息", "全供应链存销比"),
    ]
)


def replenishment(rows):
    return pd.DataFrame(rows, columns=REPLENISHMENT_COLUMNS)


def row(code, stock=100, forecast=10, revised=np.nan, target=6, status="Active", group="G1", group_desc="Group 1", category="C1", level="Parent"):
    return [code, f"Item {code}", status, group, group_desc, category, level, "RA", target, stock, 5, forecast, revised, 12]


class InventoryStructureAnalyzerTests(unittest.TestCase):
    def setUp(self):
        self.analyzer = InventoryStructureAnalyzer()

    def analyze(self, rows, prices=None):
        prices = prices if prices is not None else pd.DataFrame({"物料代码": [r[0] for r in rows], "单价": [2] * len(rows)})
        return self.analyzer.analyze(replenishment(rows), prices)

    def test_revised_forecast_has_priority(self):
        result = self.analyze([row("A", stock=90, forecast=10, revised=30)])
        self.assertEqual(result.detail.loc[0, "有效预测月销"], 30)
        self.assertEqual(result.detail.loc[0, "可用库存存销比"], 3)

    def test_price_columns_are_auto_detected_only_from_known_names(self):
        recognized = pd.DataFrame({"物料代码": ["A"], "销售组织": [8374], "单价": [2.5]})
        self.assertEqual(self.analyzer.detect_price_columns(recognized), ("物料代码", "单价"))

        unrecognized = pd.DataFrame({"物料代码": ["A"], "销售组织": [8374]})
        self.assertEqual(self.analyzer.detect_price_columns(unrecognized), ("物料代码", None))

    def test_only_parent_level_is_included_in_all_analysis(self):
        result = self.analyze([row("A", level=" parent "), row("B", level="SON"), row("C", level=np.nan)])
        self.assertEqual(result.detail["物料"].tolist(), ["A"])
        self.assertEqual(result.diagnostics["源文件记录数"], 3)
        self.assertEqual(result.diagnostics["纳入Parent记录数"], 1)
        self.assertEqual(result.diagnostics["忽略非Parent记录数"], 2)
        self.assertEqual(result.summaries["按产品状态汇总"].iloc[-1]["型号数量"], 1)

    def test_original_forecast_used_when_revised_is_empty(self):
        result = self.analyze([row("A", stock=90, forecast=15)])
        self.assertEqual(result.detail.loc[0, "有效预测月销"], 15)
        self.assertEqual(result.detail.loc[0, "可用库存存销比"], 6)

    def test_exactly_24_months_is_long_term(self):
        result = self.analyze([row("A", stock=240, forecast=10)])
        self.assertEqual(result.detail.loc[0, "库存区间"], "库存24个月以上")

    def test_exactly_target_plus_three_is_normal(self):
        result = self.analyze([row("A", stock=90, forecast=10, target=6)])
        self.assertEqual(result.detail.loc[0, "库存区间"], "正常及关注库存")

    def test_intervals_are_mutually_exclusive_and_complete_for_valid_rows(self):
        result = self.analyze([row("A", 250), row("B", 100), row("C", 90)])
        codes = []
        for name in ["库存24个月以上", "超目标库存", "正常及关注库存"]:
            codes.extend(result.views[name]["物料"].tolist())
        self.assertCountEqual(codes, ["A", "B", "C"])
        self.assertEqual(len(codes), len(set(codes)))

    def test_nonpositive_forecast_does_not_divide_by_zero(self):
        result = self.analyze([row("A", forecast=0), row("B", forecast=-2)])
        self.assertTrue(result.detail["可用库存存销比"].isna().all())
        self.assertTrue(result.detail["计算状态"].str.contains("无有效销量").all())
        self.assertTrue(result.detail["库存区间"].eq("数据问题").all())

    def test_missing_forecast_and_stock_are_not_coerced_to_zero(self):
        result = self.analyze([row("A", forecast=np.nan), row("B", stock=np.nan)])
        self.assertIn("预测月销缺失", result.detail.loc[0, "计算状态"])
        self.assertIn("可用库存缺失", result.detail.loc[1, "计算状态"])
        self.assertTrue(result.detail["可用库存存销比"].isna().all())

    def test_missing_price(self):
        result = self.analyze([row("A")], pd.DataFrame({"物料代码": ["B"], "单价": [2]}))
        self.assertTrue(pd.isna(result.detail.loc[0, "单价"]))
        self.assertTrue(pd.isna(result.detail.loc[0, "可用库存金额"]))
        self.assertIn("价格缺失", result.detail.loc[0, "计算状态"])
        self.assertEqual(result.diagnostics["价格缺失数量"], 1)

    def test_duplicate_material_with_same_price_is_deduplicated(self):
        prices = pd.DataFrame({"物料代码": ["A", "A"], "单价": [2, 2]})
        result = self.analyze([row("A")], prices)
        self.assertEqual(result.detail.loc[0, "单价"], 2)
        self.assertNotIn("价格冲突", result.detail.loc[0, "计算状态"])

    def test_duplicate_material_with_conflicting_prices_is_flagged(self):
        prices = pd.DataFrame({"物料代码": ["A", "A"], "单价": [2, 3]})
        result = self.analyze([row("A")], prices)
        self.assertTrue(pd.isna(result.detail.loc[0, "单价"]))
        self.assertIn("价格冲突", result.detail.loc[0, "计算状态"])
        self.assertEqual(result.diagnostics["价格冲突数量"], 1)

    def test_model_counts_are_distinct_material_codes(self):
        result = self.analyze([row("A"), row("A"), row("B")])
        summary = result.summaries["按产品状态汇总"]
        self.assertEqual(summary.iloc[-1]["型号数量"], 2)

    def test_three_summary_totals_match_detail_amount(self):
        result = self.analyze([row("A", stock=10), row("B", stock=20), row("C", stock=30)])
        detail_total = result.detail["可用库存金额"].sum(min_count=1)
        for name in ["按产品状态汇总", "按产品组汇总"]:
            summary = result.summaries[name]
            self.assertEqual(summary.iloc[-1][SUMMARY_AMOUNT_COLUMN], detail_total / 1_000_000)
            self.assertAlmostEqual(summary.iloc[:-1][SUMMARY_SHARE_COLUMN].sum(), 1.0)
            self.assertEqual(summary.iloc[-1][SUMMARY_SHARE_COLUMN], 1.0)
        trade_detail = result.summaries["按国贸分类码汇总"]
        self.assertEqual(trade_detail[SUMMARY_AMOUNT_COLUMN].sum(), detail_total / 1_000_000)
        self.assertAlmostEqual(trade_detail[SUMMARY_SHARE_COLUMN].sum(), 1.0)
        self.assertTrue(result.diagnostics["汇总金额核对一致"])

    def test_trade_category_tables_include_tiers_and_single_detail(self):
        rows = [
            row("A", stock=2_000_000, category="C1"),
            row("B", stock=750_000, category="C2"),
            row("C", stock=200_000, category="C3"),
            row("D", stock=1_000, category="C4"),
            row("E", stock=-10_000, category="C5"),
            row("F", stock=np.nan, category="C6"),
        ]
        prices = pd.DataFrame({"物料代码": list("ABCDEF"), "单价": [1] * 6})
        result = self.analyze(rows, prices)
        overview = result.summaries["国贸分类码分层"]
        detail = result.summaries["按国贸分类码汇总"]

        self.assertNotIn("国贸重点分类码", result.summaries)
        self.assertNotIn("国贸数据问题", result.summaries)
        self.assertIn("单型号库存金额（k USD）", detail.columns)
        self.assertIn("累计占比", detail.columns)
        self.assertIn("库存分层分类", detail.columns)
        self.assertIn("负库存金额", detail["数据状态"].tolist())
        self.assertIn("金额缺失", detail["数据状态"].tolist())
        self.assertIn("金额接近零", detail["数据状态"].tolist())
        self.assertEqual(
            overview.iloc[:-1]["库存分层分类"].tolist(),
            [
                "核心（≥1 M USD）",
                "重点（0.5–1 M USD）",
                "一般（0.1–0.5 M USD）",
                "长尾（0–0.1 M USD）",
                "负金额",
            ],
        )
        self.assertEqual(overview.iloc[-1]["库存分层分类"], "合计")

    def test_missing_categories_are_labeled_unclassified(self):
        result = self.analyze([row("A", status=np.nan, group=np.nan, group_desc=np.nan, category=np.nan)])
        self.assertEqual(result.summaries["按产品状态汇总"].iloc[0]["产品状态"], "未分类")
        self.assertEqual(result.summaries["按国贸分类码汇总"].iloc[0]["国贸产品分类码"], "未分类")

    def test_excel_has_required_sheets_column_order_and_charts(self):
        result = self.analyze([row("A", stock=250), row("B", stock=100), row("C", stock=90)])
        excel_bytes = self.analyzer.export_excel(result)
        workbook = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        self.assertEqual(workbook.sheetnames, SHEET_NAMES)
        for sheet_name in SHEET_NAMES[:3]:
            headers = [cell.value for cell in workbook[sheet_name][1]]
            self.assertLess(headers.index("可用库存存销比"), headers.index("全供应链存销比"))
            self.assertEqual(headers, DETAIL_COLUMNS)
        for sheet_name in ["按产品状态汇总", "按产品组汇总"]:
            headers = [cell.value for cell in workbook[sheet_name][1]]
            self.assertIn(SUMMARY_AMOUNT_COLUMN, headers)
            self.assertIn(SUMMARY_SHARE_COLUMN, headers)
            self.assertGreaterEqual(len(workbook[sheet_name]._charts), 2)
        self.assertEqual(len(workbook["国贸分类码分层"]._charts), 1)
        self.assertGreaterEqual(len(workbook["按国贸分类码汇总"]._charts), 1)
        self.assertIn("累计占比", [cell.value for cell in workbook["按国贸分类码汇总"][1]])
        self.assertNotIn("国贸重点分类码", workbook.sheetnames)
        self.assertNotIn("国贸数据问题", workbook.sheetnames)
        with zipfile.ZipFile(io.BytesIO(excel_bytes)) as archive:
            chart_xml = "".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("xl/charts/chart")
            )
        self.assertIn("100.0%", chart_xml)
        self.assertNotIn("核心与重点产品库存金额占比", chart_xml)

    def test_product_group_excel_chart_matches_web_doughnut_rules(self):
        rows = [
            row("A", stock=100, group="G1", group_desc="Large Group"),
            row("B", stock=100, group="G2", group_desc="Outside Label Group"),
            row("C", stock=100, group="G3", group_desc="Unlabelled Group"),
        ]
        prices = pd.DataFrame(
            {"物料代码": ["A", "B", "C"], "单价": [300_000, 12_000, 2_000]}
        )
        result = self.analyze(rows, prices)
        excel_bytes = self.analyzer.export_excel(result)
        workbook = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        charts = workbook["按产品组汇总"]._charts
        self.assertEqual(type(charts[1]).__name__, "DoughnutChart")

        with zipfile.ZipFile(io.BytesIO(excel_bytes)) as archive:
            chart_xml = "".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("xl/charts/chart")
            )
            drawing_xml = "".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("xl/drawings/drawing") and name.endswith(".xml")
            )
        self.assertIn("Large Group", chart_xml)
        self.assertIn("Outside Label Group", chart_xml)
        self.assertIn("<c:dLblPos val=\"outEnd\"/>", chart_xml)
        self.assertIn("<c:showLeaderLines val=\"1\"/>", chart_xml)
        self.assertIn("<c:delete val=\"1\"/>", chart_xml)
        self.assertIn("总可用库存金额", drawing_xml)
        self.assertIn("低于 1 M USD", drawing_xml)

    def test_excel_accepts_legacy_trade_tier_column_name(self):
        result = self.analyze([row("A", stock=2_000_000), row("B", stock=750_000)])
        result.summaries["国贸分类码分层"].rename(
            columns={"库存分层分类": "金额层级"}, inplace=True
        )
        result.summaries["按国贸分类码汇总"].rename(
            columns={"库存分层分类": "金额层级"}, inplace=True
        )
        workbook = openpyxl.load_workbook(io.BytesIO(self.analyzer.export_excel(result)))
        for sheet_name in ["国贸分类码分层", "按国贸分类码汇总"]:
            headers = [cell.value for cell in workbook[sheet_name][1]]
            self.assertIn("库存分层分类", headers)
            self.assertNotIn("金额层级", headers)


if __name__ == "__main__":
    unittest.main()
