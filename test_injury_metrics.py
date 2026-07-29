from pathlib import Path
import unittest
from datetime import date

import pandas as pd

from app import normalize_date_selection
from injury_metrics import aggregate_scope, build_monthly_detail, fiscal_year_start
from sanitize_workbook import PII_HEADERS, clean_header


WORKBOOK = Path(__file__).with_name(
    "AIE_Ergo_Dashboard_Template NICOLES VERSION.xlsx"
)

INJURY_COLUMNS = {
    "Site",
    "System Case ID",
    "Case Date",
    "U.S. OSHA Recordable?",
    "Inj. DAFW?",
    "Inj. RWA?",
    "Inj. Other Recordable Case?",
    "Ill. DAFW?",
    "Ill. RWA?",
    "Ill. Other Recordable Case?",
}


class InjuryMetricTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.ehs = pd.read_excel(WORKBOOK, sheet_name="EHS_Hours_Injuries")
        cls.ergo = pd.read_excel(
            WORKBOOK,
            sheet_name="Ergo_Injuries",
            usecols=lambda column: " ".join(str(column).split()) in INJURY_COLUMNS,
        )
        cls.detail, cls.latest_month = build_monthly_detail(cls.ehs, cls.ergo)

    def test_fiscal_calendar_anchor_and_history(self) -> None:
        self.assertEqual(fiscal_year_start(2026), pd.Timestamp("2025-12-28"))
        self.assertEqual(fiscal_year_start(2025), pd.Timestamp("2024-12-29"))
        self.assertEqual(fiscal_year_start(2024), pd.Timestamp("2023-12-31"))
        self.assertEqual(fiscal_year_start(2023), pd.Timestamp("2023-01-01"))

    def test_only_completed_months_are_published(self) -> None:
        self.assertEqual(self.latest_month, pd.Timestamp("2026-06-01"))
        self.assertEqual(self.detail["Date"].max(), pd.Timestamp("2026-06-01"))

    def test_ge_appliances_latest_rolling_metrics(self) -> None:
        metrics = aggregate_scope(self.detail, "GE Appliances")
        latest = metrics.iloc[-1]
        self.assertEqual(latest["Date"], pd.Timestamp("2026-06-01"))
        self.assertAlmostEqual(latest["Hours R12"], 22_318_366)
        self.assertAlmostEqual(latest["EHS Recordables R12"], 329)
        self.assertAlmostEqual(latest["Ergo All Cases R12"], 702)
        self.assertAlmostEqual(latest["Ergo Recordable R12"], 133)
        self.assertAlmostEqual(latest["Ergo DART R12"], 98)
        self.assertAlmostEqual(
            latest["Ergo Share of EHS Recordables"], 133 / 329 * 100
        )
        self.assertAlmostEqual(
            latest["Ergo All Rate R12"], 702 * 200_000 / 22_318_366
        )

    def test_supply_chain_excludes_bodewell(self) -> None:
        gea = aggregate_scope(self.detail, "GE Appliances").iloc[-1]
        supply_chain = aggregate_scope(self.detail, "Supply Chain").iloc[-1]
        bode = aggregate_scope(self.detail, "BODE").iloc[-1]
        self.assertAlmostEqual(
            gea["Hours R12"] - supply_chain["Hours R12"], bode["Hours R12"]
        )
        self.assertAlmostEqual(
            gea["EHS Recordables R12"] - supply_chain["EHS Recordables R12"],
            bode["EHS Recordables R12"],
        )

    def test_direct_identifiers_are_blank_in_bundled_workbook(self) -> None:
        pii = pd.read_excel(
            WORKBOOK,
            sheet_name="Ergo_Injuries",
            usecols=lambda column: clean_header(column) in PII_HEADERS,
        )
        self.assertEqual(len(pii.columns), len(PII_HEADERS))
        self.assertEqual(int(pii.notna().sum().sum()), 0)

    def test_transient_single_date_tuple_is_supported(self) -> None:
        start, end = normalize_date_selection(
            (date(2026, 5, 1),),
            fallback_start=pd.Timestamp("2023-12-01"),
            fallback_end=pd.Timestamp("2026-06-01"),
        )
        self.assertEqual(start, pd.Timestamp("2026-05-01"))
        self.assertEqual(end, pd.Timestamp("2026-05-01"))


if __name__ == "__main__":
    unittest.main()
