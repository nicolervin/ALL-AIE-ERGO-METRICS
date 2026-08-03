from pathlib import Path
import unittest
from datetime import date

import pandas as pd

from app import (
    apply_factory_date_filter,
    apply_scope_ytd_filter,
    available_npi_programs,
    build_assessment_df,
    build_npi_metrics_df,
    default_site_risk_reduction_goals,
    filter_injury_history,
    filter_through_end_month,
    has_npi_values,
    make_npi_detail_snapshot_chart,
    normalize_date_selection,
    previous_calendar_month,
    risk_reduction_goal_for_scope,
)
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

    def test_previous_month_default_and_trailing_injury_history(self) -> None:
        default_start, default_end = previous_calendar_month("2026-07-30")
        self.assertEqual(default_start, pd.Timestamp("2026-06-01"))
        self.assertEqual(default_end, pd.Timestamp("2026-06-30"))

        monthly = pd.DataFrame(
            {"Date": pd.date_range("2024-01-01", "2026-06-01", freq="MS")}
        )
        history = filter_injury_history(monthly, default_end)
        self.assertEqual(len(history), 12)
        self.assertEqual(history["Date"].min(), pd.Timestamp("2025-07-01"))
        self.assertEqual(history["Date"].max(), pd.Timestamp("2026-06-01"))

        summary_history = filter_through_end_month(monthly, default_end)
        self.assertIn(pd.Timestamp("2025-06-01"), summary_history["Date"].tolist())
        self.assertEqual(summary_history["Date"].max(), pd.Timestamp("2026-06-01"))

    def test_critical_assessments_remain_separate_from_red(self) -> None:
        source = pd.DataFrame(
            {
                "Factory": ["AP1"],
                "Year": [2026],
                "Month": [6],
                "Critical_Red": [2],
                "Red": [3],
                "Yellow": [4],
                "Green": [5],
            }
        )
        assessments = build_assessment_df(source)
        counts = assessments.set_index("Risk Level")["Count"].to_dict()
        self.assertEqual(counts["Critical"], 2)
        self.assertEqual(counts["Red"], 3)
        self.assertEqual(sum(counts.values()), 14)

    def test_site_goals_and_risk_reduction_are_year_to_date(self) -> None:
        projects = pd.DataFrame(
            {
                "Factory": ["AP1", "AP1", "AP1", "AP3", "AP3"],
                "Date": pd.to_datetime(
                    [
                        "2025-12-01",
                        "2026-01-01",
                        "2026-06-01",
                        "2026-01-01",
                        "2026-07-01",
                    ]
                ),
                "Reduction Points": [99, 10, 20, 30, 40],
                "Target Points": [100, 120, 120, 180, 180],
            }
        )
        ytd = apply_scope_ytd_filter(
            projects,
            "GE Appliances",
            pd.Timestamp("2026-06-30"),
        )
        self.assertEqual(ytd["Reduction Points"].sum(), 60)
        self.assertEqual(ytd["Date"].min(), pd.Timestamp("2026-01-01"))
        self.assertEqual(ytd["Date"].max(), pd.Timestamp("2026-06-01"))

        goals = default_site_risk_reduction_goals(
            projects,
            ["AP1", "AP3", "BODE"],
            2026,
        )
        self.assertEqual(goals, {"AP1": 120.0, "AP3": 180.0, "BODE": 0.0})
        self.assertEqual(
            risk_reduction_goal_for_scope(goals, "GE Appliances"),
            300.0,
        )
        self.assertEqual(
            risk_reduction_goal_for_scope(goals, "AP1"),
            120.0,
        )

    def test_npi_ergo_red_status_colors_follow_status_names(self) -> None:
        reds = pd.DataFrame(
            {
                "Status": [
                    "Closed Engineering",
                    "Started",
                    "Validation",
                    "Closed Administrative",
                ],
                "Count": [4, 1, 2, 3],
            }
        )
        chart = make_npi_detail_snapshot_chart(reds)
        colors = dict(zip(chart.data[0].x, chart.data[0].marker.color))
        self.assertEqual(colors["Started"], "#9CA3AF")
        self.assertEqual(colors["Validation"], "#3B82F6")
        self.assertEqual(colors["Closed Admin"], "#A7D7A9")
        self.assertEqual(colors["Closed Eng"], "#2E7D32")

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

    def test_npi_program_options_include_aie_and_ergo_only_programs(self) -> None:
        aie = pd.DataFrame({"Program": ["CPO Electric Units"]})
        ergo = pd.DataFrame({"Program": ["AP1 Summit Washer"]})
        self.assertEqual(
            available_npi_programs(aie, ergo),
            ["AP1 Summit Washer", "CPO Electric Units"],
        )

    def test_incomplete_npi_aie_row_is_retained_with_missing_values(self) -> None:
        frame = pd.DataFrame(
            {
                "NPI_Name": ["CPO Electric Units"],
                "Baseline_Headcount": [None],
                "Future_Headcount": [None],
            }
        )
        metrics = build_npi_metrics_df(frame)
        self.assertEqual(metrics["Program"].tolist(), ["CPO Electric Units"])
        self.assertFalse(has_npi_values(metrics, "EAU Target", "EAU Actual"))
        self.assertTrue(metrics["EAU Target"].isna().all())

    def test_undated_incomplete_npi_row_survives_date_filter(self) -> None:
        frame = pd.DataFrame(
            {
                "Program": ["CPO Electric Units"],
                "Factory": ["Other"],
                "Date": [pd.NaT],
            }
        )
        filtered = apply_factory_date_filter(
            frame,
            [],
            pd.Timestamp("2026-07-01"),
            pd.Timestamp("2026-07-31"),
            include_undated=True,
        )
        self.assertEqual(filtered["Program"].tolist(), ["CPO Electric Units"])


if __name__ == "__main__":
    unittest.main()
