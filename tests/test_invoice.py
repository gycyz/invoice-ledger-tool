# -*- coding: utf-8 -*-
"""回归测试：钉住 2026-11-19 修的发票级核对逻辑。

    python -m unittest discover -s tests -v
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "脚本"))

import invoice_ledger as tool  # noqa: E402

DETAIL_COLS = ["序号", "数电发票号码", "开票日期", "销方名称", "销方识别号", "购买方名称",
               "购方识别号", "发票票种", "特定业务类型", "发票状态", "是否正数发票", "税率",
               "金额", "税额", "价税合计", "货物或应税劳务名称", "备注"]

SELLER_ID = "91340000MA2XXXXX1X"


def detail_frame(rows) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=DETAIL_COLS)


def base_rows():
    """一张 13% 专票 + 一行「不征税」明细（民航发展基金那样）。"""
    return [
        [1, "1234567890123456789", "2026-09-01 10:00", "某供应商", SELLER_ID,
         "本公司", "91340000MA2YYYYY2Y", "增值税专用发票", "", "正常", "是", "13%",
         100.0, 13.0, 113.0, "办公用品", ""],
        [2, "1234567890123456789", "2026-09-01 10:00", "某供应商", SELLER_ID,
         "本公司", "91340000MA2YYYYY2Y", "增值税专用发票", "", "正常", "是", "不征税",
         0.0, 0.0, 50.0, "民航发展基金", ""],
    ]


def invoice_sheet(rows):
    cols = ["序号", "数电发票号码", "销方识别号", "开票日期", "价税合计"]
    return pd.DataFrame(rows, columns=cols)


class BuildTests(unittest.TestCase):
    def test_two_totals_are_separated(self):
        inv = tool.build_invoices(detail_frame(base_rows()))
        self.assertEqual(len(inv), 1)
        row = inv.iloc[0]
        self.assertAlmostEqual(row["价税合计"], 163.0)
        self.assertAlmostEqual(row["其中不征税"], 50.0)
        self.assertAlmostEqual(row["票面价税合计"], 113.0)   # 与税局票面一致
        self.assertAlmostEqual(row["税额_明细合计"], 13.0)


class CheckR03Tests(unittest.TestCase):
    """R03 曾经是死代码：匹配条件里带金额容差，导致「金额不符」永不触发。"""

    def _run(self, sheet_amount):
        detail = detail_frame(base_rows())
        inv = tool.build_invoices(detail)
        data = {"sheets": {"detail": "信息汇总表", "invoice": "发票基础信息"},
                "invoice": invoice_sheet([[1, "1234567890123456789", SELLER_ID,
                                           pd.Timestamp("2026-09-01 10:00"), sheet_amount]])}
        return tool.run_checks(data, inv, detail)

    def test_detects_amount_mismatch(self):
        issues = self._run(120.0)          # 票面应为 113，发票级表写 120
        hits = [i for i in issues if i["检查项"] == "发票级金额与票面不符"]
        self.assertEqual(len(hits), 1, issues)
        self.assertEqual(hits[0]["级别"], "高")
        self.assertAlmostEqual(hits[0]["金额影响"], -7.0)

    def test_no_false_alarm_when_amount_matches(self):
        issues = self._run(113.0)
        self.assertFalse([i for i in issues if i["检查项"] == "发票级金额与票面不符"])

    def test_unmatched_invoice_reported(self):
        detail = detail_frame(base_rows())
        inv = tool.build_invoices(detail)
        data = {"sheets": {"detail": "信息汇总表", "invoice": "发票基础信息"},
                "invoice": invoice_sheet([[1, "9999999999999999999", "91340000MA2ZZZZZ9Z",
                                           pd.Timestamp("2026-10-20 09:00"), 113.0]])}
        issues = tool.run_checks(data, inv, detail)
        self.assertTrue([i for i in issues if i["检查项"] == "发票级表有票未能匹配明细表"])


class CheckRedTests(unittest.TestCase):
    def test_red_invoice_without_blue_number(self):
        rows = [[1, "1111111111111111111", "2026-09-02 09:00", "某供应商", SELLER_ID,
                 "本公司", "91340000MA2YYYYY2Y", "增值税专用发票", "", "正常", "否", "13%",
                 -100.0, -13.0, -113.0, "办公用品", "无蓝票信息"]]
        detail = detail_frame(rows)
        inv = tool.build_invoices(detail)
        issues = tool.run_checks({"sheets": {"detail": "信息汇总表"}}, inv, detail)
        self.assertTrue([i for i in issues if i["检查项"] == "红字发票未见蓝票号"])


class DeductibleTests(unittest.TestCase):
    def test_constant_is_used(self):
        """可抵扣口径必须来自 DEDUCTIBLE_KINDS，不能再硬编码一份。"""
        src = (ROOT / "脚本" / "invoice_ledger.py").read_text(encoding="utf-8")
        self.assertIn("any(k in s for k in DEDUCTIBLE_KINDS)", src)


if __name__ == "__main__":
    unittest.main()
