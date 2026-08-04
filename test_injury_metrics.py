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
    build_risk_reduction_goals_df,
    default_combined_pdf_sites,
    default_site_risk_reduction_goals,
    filter_injury_history,
    filter_through_end_month,
    has_npi_values,
    injury_trend_status,
    latest_npi_snapshots,
    make_ergo_classification_chart,
    make_npi_detail_snapshot_chart,
    npi_reporting_month,
    normalize_date_selection,
    open_ergo_reds_status,
    previous_calendar_month,
    reduction_status,
    risk_reduction_goal_for_scope,
    _snapshot_operational_items,
    _snapshot_case_table_rows,
    _snapshot_factory_kpi_items,
    _snapshot_factory_kpi_statuses,
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

    def test_combined_pdf_defaults_to_all_scopes_except_other(self) -> None:
        scopes = [
            "GE Appliances",
            "Supply Chain",
            "AP1",
            "AP2",
            "Other",
        ]
        self.assertEqual(
            default_combined_pdf_sites(scopes),
            ["GE Appliances", "Supply Chain", "AP1", "AP2"],
        )

    def test_pdf_operational_strip_uses_medical_request_labels(self) -> None:
        items = _snapshot_operational_items(
            {
                "closed_requests": 12,
                "past_due_requests": 3,
                "latest_year": 2026,
                "current_workers_comp": 1234,
            }
        )
        self.assertEqual(
            [item[0] for item in items],
            [
                "Closed Medical Requests",
                "Past-Due Medical Requests",
                "2026 Workers Comp",
            ],
        )
        self.assertNotIn("Total Assessments", [item[0] for item in items])

    def test_pdf_factory_kpis_show_prior_year_change_and_exclude_hours(self) -> None:
        items = _snapshot_factory_kpi_items(
            {
                "total_case_rate": 3.41,
                "total_case_change": "-5.0%",
                "current_recordable": 0.44,
                "recordable_change": "+2.0%",
                "current_dart": 0.22,
                "dart_change": "-1.0%",
                "current_ehs_recordable": 2.67,
                "ehs_recordable_change": "+4.0%",
                "ergo_share": 16.7,
                "ergo_share_change": "-3.0%",
            }
        )
        self.assertNotIn("Hours", [item[0] for item in items])
        self.assertTrue(all("vs prior year" in item[2] for item in items))

    def test_pdf_injury_kpi_colors_follow_trends(self) -> None:
        self.assertEqual(injury_trend_status(2.0, 3.0), "green")
        self.assertEqual(injury_trend_status(4.0, 3.0), "red")
        self.assertEqual(injury_trend_status(3.0, 3.0), "yellow")
        self.assertEqual(injury_trend_status(0.0, 0.0, False), "neutral")

        summary = {
            "total_case_rate": 2.0,
            "prior_total_case_rate": 3.0,
            "has_total_case_comparison": True,
            "current_recordable": 4.0,
            "prior_recordable": 3.0,
            "has_recordable_comparison": True,
            "current_dart": 3.0,
            "prior_dart": 3.0,
            "has_dart_comparison": True,
            "current_ehs_recordable": 0.0,
            "prior_ehs_recordable": 0.0,
            "has_ehs_recordable_comparison": False,
            "ergo_share": 1.0,
            "prior_ergo_share": 2.0,
            "has_ergo_share_comparison": True,
        }
        self.assertEqual(
            _snapshot_factory_kpi_statuses(summary),
            ["green", "red", "yellow", "neutral", "green"],
        )

    def test_pdf_case_table_includes_rates_next_to_counts(self) -> None:
        summary = {
            "rolling_hours": 2_000_000,
            "ergo_all_cases_r12": 40,
            "ergo_first_aid_r12": 20,
            "ergo_recordable_r12": 10,
            "ergo_dafw_r12": 4,
            "ergo_restricted_r12": 6,
            "ergo_dart_r12": 10,
            "ehs_recordables_r12": 30,
        }
        rows = _snapshot_case_table_rows(summary)
        self.assertEqual(rows[0], ["Rolling-12 cases", "Count", "Rate"])
        self.assertEqual(rows[1], ["All ergonomic cases", "40", "4.00"])
        self.assertEqual(rows[4], ["DAFW", "4", "0.40"])

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

        goal_sheet = build_risk_reduction_goals_df(
            pd.DataFrame(
                {
                    "Site": ["AP1", "AP3"],
                    "Year": [2026, 2026],
                    "Risk_Reduction_Goals": [170, 300],
                }
            )
        )
        goals = default_site_risk_reduction_goals(
            goal_sheet,
            ["AP1", "AP3", "BODE"],
            2026,
        )
        self.assertEqual(goals, {"AP1": 170.0, "AP3": 300.0, "BODE": 0.0})
        self.assertEqual(
            risk_reduction_goal_for_scope(goals, "GE Appliances"),
            470.0,
        )
        self.assertEqual(
            risk_reduction_goal_for_scope(goals, "AP1"),
            170.0,
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

    def test_open_ergo_reds_status_is_green_only_at_zero(self) -> None:
        self.assertEqual(open_ergo_reds_status(0), "green")
        self.assertEqual(open_ergo_reds_status(1), "red")
        self.assertEqual(open_ergo_reds_status(12), "red")
        self.assertEqual(open_ergo_reds_status(0, has_ergo_data=False), "neutral")

    def test_chart_legend_is_reserved_below_plot_area(self) -> None:
        chart = make_ergo_classification_chart(self.detail)
        self.assertEqual(chart.layout.legend.orientation, "h")
        self.assertEqual(chart.layout.legend.yanchor, "top")
        self.assertLess(chart.layout.legend.y, 0)
        self.assertGreaterEqual(chart.layout.margin.b, 80)

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

    def test_npi_timestamp_on_end_date_is_included(self) -> None:
        frame = pd.DataFrame(
            {
                "Program": ["AP1 Nexus Dryer", "Too Late"],
                "Factory": ["AP1", "AP1"],
                "Date": [
                    pd.Timestamp("2026-07-31 13:46"),
                    pd.Timestamp("2026-08-01 00:00"),
                ],
            }
        )
        filtered = apply_factory_date_filter(
            frame,
            [],
            pd.Timestamp("2026-07-01"),
            pd.Timestamp("2026-07-31"),
        )
        self.assertEqual(filtered["Program"].tolist(), ["AP1 Nexus Dryer"])

    def test_npi_reporting_month_allows_early_and_late_submissions(self) -> None:
        submissions = pd.Series(
            pd.to_datetime(
                [
                    "2026-07-27 08:00",
                    "2026-07-31 13:46",
                    "2026-08-01 09:00",
                    "2026-08-07 17:00",
                    "2026-08-08 09:00",
                    "2027-01-07 09:00",
                ]
            )
        )
        reporting_months = npi_reporting_month(submissions)
        self.assertEqual(
            reporting_months.tolist(),
            [
                pd.Timestamp("2026-07-01"),
                pd.Timestamp("2026-07-01"),
                pd.Timestamp("2026-07-01"),
                pd.Timestamp("2026-07-01"),
                pd.Timestamp("2026-08-01"),
                pd.Timestamp("2026-12-01"),
            ],
        )

    def test_reduction_status_treats_lower_values_as_favorable(self) -> None:
        self.assertEqual(reduction_status(100, 90), "green")
        self.assertEqual(reduction_status(100, 110), "red")
        self.assertEqual(reduction_status(100, 100), "yellow")
        self.assertEqual(reduction_status(float("nan"), 100), "neutral")

    def test_latest_npi_snapshot_replaces_earlier_monthly_submission(self) -> None:
        frame = pd.DataFrame(
            {
                "Program": ["AP1 Nexus Dryer", "AP1 Nexus Dryer"],
                "Date": [pd.Timestamp("2026-07-01")] * 2,
                "Submission Date": [
                    pd.Timestamp("2026-07-31 13:46"),
                    pd.Timestamp("2026-08-03 09:15"),
                ],
                "Part Count Actual": [264, 248],
            }
        )
        latest = latest_npi_snapshots(frame)
        self.assertEqual(latest["Part Count Actual"].tolist(), [248])

    def test_latest_ergo_snapshot_keeps_all_status_rows(self) -> None:
        frame = pd.DataFrame(
            {
                "Program": ["AP1 Nexus Dryer"] * 4,
                "Date": [pd.Timestamp("2026-07-01")] * 4,
                "Submission Date": [
                    pd.Timestamp("2026-07-31 13:46"),
                    pd.Timestamp("2026-07-31 13:46"),
                    pd.Timestamp("2026-08-03 09:15"),
                    pd.Timestamp("2026-08-03 09:15"),
                ],
                "Status": ["Open", "Pending", "Open", "Pending"],
                "Count": [2, 7, 0, 4],
            }
        )
        latest = latest_npi_snapshots(frame)
        self.assertEqual(latest["Status"].tolist(), ["Open", "Pending"])
        self.assertEqual(latest["Count"].tolist(), [0, 4])


if __name__ == "__main__":
    unittest.main()
