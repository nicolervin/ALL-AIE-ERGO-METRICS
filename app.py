import io
import importlib
import html
import re
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from injury_metrics import (
    aggregate_scope,
    available_scopes,
    build_monthly_detail,
    normalize_site,
)

st.set_page_config(page_title="Ergonomics & Industrial Engineering Dashboard", layout="wide")

SHEET_ALIASES = {
    "ergo_assessments": ["ergo_assessments", "Ergo_Assessments_Risk_Reduction"],
    "risk_projects": ["risk_projects", "Ergo_Assessments_Risk_Reduction"],
    "risk_reduction_goals": ["risk_reduction_goals", "Ergo_Risk_Reduction"],
    "medical_requests": ["medical_requests", "Medical_Requests"],
    "workers_comp": ["workers_comp", "Ergo_Workers_Comp"],
    "dashboard_goals": ["dashboard_goals", "Dashboard_Goals"],
    "npi_metrics": ["npi_metrics", "NPI_AIE", "AIE_NPI"],
    "npi_ergo_reds": ["npi_ergo_reds", "NPI_Ergo", "Ergo_NPI"],
}

REQUIRED_COLUMNS = {
    "production_injuries": ["Date", "Factory", "Hours", "Ergo All Cases"],
    "ergo_assessments": ["Factory", "Risk Level", "Count"],
    "risk_projects": ["Factory", "Date", "Reduction Points", "Target Points"],
    "risk_reduction_goals": ["Factory", "Year", "Goal"],
    "medical_requests": ["Factory", "Status", "Count"],
    "workers_comp": ["Factory", "Date", "Amount"],
    "dashboard_goals": ["Year", "Metric", "Goal"],
    "npi_metrics": ["Program", "Factory", "Headcount Actual", "Headcount Target"],
    "npi_ergo_reds": ["Program", "Status", "Count"],
}

OPTIONAL_SHEETS = {"dashboard_goals", "risk_reduction_goals"}


@st.cache_data(show_spinner=False)
def load_workbook_data(file_bytes: bytes | None, file_name: str | None = None) -> Tuple[Dict[str, pd.DataFrame], str | None]:
    if file_bytes is None:
        candidate_paths = [
            Path(__file__).with_name("AIE_Ergo_Dashboard_Template NICOLES VERSION.xlsx"),
            Path.cwd() / "AIE_Ergo_Dashboard_Template NICOLES VERSION.xlsx",
        ]
        selected_path = next((path for path in candidate_paths if path.exists()), None)
        if selected_path is None:
            return {}, "No workbook was provided and the attached template was not found."
        bytes_data = selected_path.read_bytes()
        source_name = selected_path.name
    else:
        bytes_data = file_bytes
        source_name = file_name or "uploaded_workbook.xlsx"

    try:
        excel_file = pd.ExcelFile(io.BytesIO(bytes_data), engine="openpyxl")
    except Exception as exc:
        return {}, f"Unable to read workbook: {exc}"

    normalized: Dict[str, pd.DataFrame] = {}
    errors: list[str] = []

    if "EHS_Hours_Injuries" not in excel_file.sheet_names:
        errors.append("Missing required sheet: EHS_Hours_Injuries")
    if "Ergo_Injuries" not in excel_file.sheet_names:
        errors.append("Missing required sheet: Ergo_Injuries")
    if not errors:
        try:
            ehs_frame = pd.read_excel(
                io.BytesIO(bytes_data),
                sheet_name="EHS_Hours_Injuries",
                engine="openpyxl",
            )
            injury_columns = {
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
            ergo_frame = pd.read_excel(
                io.BytesIO(bytes_data),
                sheet_name="Ergo_Injuries",
                engine="openpyxl",
                usecols=lambda column: re.sub(r"\s+", " ", str(column).strip())
                in injury_columns,
            )
            detail, latest_complete_month = build_monthly_detail(ehs_frame, ergo_frame)
            detail.attrs["latest_complete_month"] = latest_complete_month
            normalized["production_injuries"] = detail
        except Exception as exc:
            errors.append(f"Unable to build injury metrics: {exc}")

    sheet_names_by_key = {
        str(sheet_name).strip().casefold(): sheet_name
        for sheet_name in excel_file.sheet_names
    }

    for dataset_name, aliases in SHEET_ALIASES.items():
        matched_sheet = None
        for alias in aliases:
            matched_sheet = sheet_names_by_key.get(str(alias).strip().casefold())
            if matched_sheet is not None:
                break
        if matched_sheet is None:
            if dataset_name in OPTIONAL_SHEETS:
                continue
            errors.append(f"Missing required sheet: {dataset_name}")
            continue

        try:
            frame = pd.read_excel(io.BytesIO(bytes_data), sheet_name=matched_sheet, engine="openpyxl")

            if dataset_name == "ergo_assessments":
                normalized[dataset_name] = build_assessment_df(frame)
            elif dataset_name == "risk_projects":
                normalized[dataset_name] = build_risk_project_df(frame)
            elif dataset_name == "risk_reduction_goals":
                normalized[dataset_name] = build_risk_reduction_goals_df(frame)
            elif dataset_name == "medical_requests":
                normalized[dataset_name] = build_medical_requests_df(frame)
            elif dataset_name == "workers_comp":
                normalized[dataset_name] = build_workers_comp_df(frame)
            elif dataset_name == "dashboard_goals":
                normalized[dataset_name] = build_dashboard_goals_df(frame)
            elif dataset_name == "npi_metrics":
                normalized[dataset_name] = build_npi_metrics_df(frame)
            elif dataset_name == "npi_ergo_reds":
                normalized[dataset_name] = build_npi_ergo_reds_df(frame)
        except Exception as exc:
            errors.append(f"Unable to load sheet {matched_sheet}: {exc}")

    if errors:
        return normalized, "\n".join(errors)
    return normalized, None


def find_column(frame: pd.DataFrame, candidates: list[str]) -> str | None:
    lowered = {str(col).strip().lower(): col for col in frame.columns}
    for name in candidates:
        if str(name).strip().lower() in lowered:
            return lowered[str(name).strip().lower()]
    return None


def coerce_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.extract(r"(-?\d+(?:\.\d+)?)")[0]
    return pd.to_numeric(cleaned, errors="coerce")


def year_month_date(
    frame: pd.DataFrame, year_column: str, month_column: str
) -> pd.Series:
    return pd.to_datetime(
        {
            "year": pd.to_numeric(frame[year_column], errors="coerce"),
            "month": pd.to_numeric(frame[month_column], errors="coerce"),
            "day": 1,
        },
        errors="coerce",
    )


def npi_reporting_month(submitted_at: pd.Series) -> pd.Series:
    """Assign NPI form submissions to their intended monthly reporting period.

    Submissions during the first seven calendar days are treated as late entries
    for the previous month. Entries in the final five days remain in the current
    month; submissions outside either edge window also retain their calendar month.
    """

    submitted_dates = pd.to_datetime(submitted_at, errors="coerce")
    reporting_months = submitted_dates.dt.to_period("M").dt.to_timestamp()
    submitted_during_grace_period = submitted_dates.dt.day.between(1, 7)
    reporting_months.loc[submitted_during_grace_period] = (
        reporting_months.loc[submitted_during_grace_period]
        - pd.DateOffset(months=1)
    )
    return reporting_months


def build_production_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    injuries_col = find_column(frame, ["Number_of_Ergo_Recordable_Injuries_This_Month", "Injuries"])
    recordable_col = find_column(frame, ["EHS_Recordable_Rate_Rolling", "Recordable Rate"])
    dart_col = find_column(frame, ["EHS_DART_Rate_Rolling", "DART Rate"])

    if not all([factory_col, year_col, month_col, injuries_col, recordable_col, dart_col]):
        raise ValueError("Production injuries sheet is missing one or more required columns.")

    df = frame[[factory_col, year_col, month_col, injuries_col, recordable_col, dart_col]].copy()
    df.columns = ["Factory", "Year", "Month", "Injuries", "Recordable Rate", "DART Rate"]
    df["Factory"] = df["Factory"].astype(str)
    df["Date"] = year_month_date(df, "Year", "Month")
    df["Injuries"] = pd.to_numeric(df["Injuries"], errors="coerce").fillna(0).astype(int)
    df["Recordable Rate"] = pd.to_numeric(df["Recordable Rate"], errors="coerce")
    df["DART Rate"] = pd.to_numeric(df["DART Rate"], errors="coerce")
    return df[["Date", "Factory", "Injuries", "Recordable Rate", "DART Rate"]]


def build_assessment_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    critical_red_col = find_column(frame, ["Critical_Red", "Critical Red"])
    red_col = find_column(frame, ["Red"])
    yellow_col = find_column(frame, ["Yellow"])
    green_col = find_column(frame, ["Green"])
    if not all([factory_col, red_col, yellow_col, green_col]):
        raise ValueError("Ergo assessments sheet is missing one or more required columns.")
    df = pd.DataFrame(
        {
            "Factory": frame[factory_col],
            "Critical": frame[critical_red_col] if critical_red_col else 0,
            "Red": frame[red_col],
            "Yellow": frame[yellow_col],
            "Green": frame[green_col],
        }
    )
    df["Factory"] = df["Factory"].map(normalize_site)
    if year_col and month_col:
        dates = year_month_date(frame, year_col, month_col)
    else:
        dates = pd.Series(pd.NaT, index=frame.index)
    df["Critical"] = pd.to_numeric(df["Critical"], errors="coerce").fillna(0)
    df["Red"] = pd.to_numeric(df["Red"], errors="coerce").fillna(0)
    df["Yellow"] = pd.to_numeric(df["Yellow"], errors="coerce").fillna(0)
    df["Green"] = pd.to_numeric(df["Green"], errors="coerce").fillna(0)
    rows = []
    for _, row in df.iterrows():
        rows.append({"Factory": row["Factory"], "Date": dates.loc[_], "Risk Level": "Critical", "Count": int(row["Critical"])})
        rows.append({"Factory": row["Factory"], "Date": dates.loc[_], "Risk Level": "Red", "Count": int(row["Red"])})
        rows.append({"Factory": row["Factory"], "Date": dates.loc[_], "Risk Level": "Yellow", "Count": int(row["Yellow"])})
        rows.append({"Factory": row["Factory"], "Date": dates.loc[_], "Risk Level": "Green", "Count": int(row["Green"])})
    return pd.DataFrame(rows)


def build_risk_project_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    date_col = find_column(frame, ["Date"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    reduction_col = find_column(frame, ["Risk_Reduction_Points_This_Month", "Reduction Points"])
    target_col = find_column(frame, ["Target Points", "Target_Points", "Target"])
    if not all([factory_col, reduction_col]):
        raise ValueError("Risk projects sheet is missing required columns.")
    df = frame[[factory_col, reduction_col]].copy()
    df.columns = ["Factory", "Reduction Points"]
    df["Factory"] = df["Factory"].map(normalize_site)
    df["Reduction Points"] = pd.to_numeric(df["Reduction Points"], errors="coerce").fillna(0)
    if year_col and month_col:
        df["Date"] = year_month_date(frame, year_col, month_col)
    elif date_col:
        df["Date"] = pd.to_datetime(frame[date_col], errors="coerce")
    else:
        df["Date"] = pd.Series([pd.Timestamp.today().normalize()] * len(frame))
    if target_col:
        df["Target Points"] = pd.to_numeric(frame[target_col], errors="coerce").fillna(0)
    else:
        monthly_average = df["Reduction Points"].mean() if not df.empty else 0
        df["Target Points"] = monthly_average * 12
    return df[["Factory", "Date", "Reduction Points", "Target Points"]]


def build_medical_requests_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    completed_col = find_column(frame, ["Completed", "Closed"])
    completed_past_due_col = find_column(frame, ["Completed_Past_Due", "Completed Past Due"])
    past_due_col = find_column(frame, ["Past_Due", "Past Due"])
    open_col = find_column(frame, ["Open"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    if not all([factory_col, completed_col]) or not (past_due_col or open_col):
        raise ValueError("Medical requests sheet is missing required columns.")
    df = pd.DataFrame(index=frame.index)
    df["Factory"] = frame[factory_col].map(normalize_site)
    df["Closed"] = pd.to_numeric(frame[completed_col], errors="coerce").fillna(0)
    if completed_past_due_col:
        df["Closed"] += pd.to_numeric(frame[completed_past_due_col], errors="coerce").fillna(0)
    source_past_due = past_due_col or open_col
    df["Past Due"] = pd.to_numeric(frame[source_past_due], errors="coerce").fillna(0)
    df["Open"] = (
        pd.to_numeric(frame[open_col], errors="coerce").fillna(0)
        if open_col and past_due_col
        else 0
    )
    if year_col and month_col:
        df["Date"] = year_month_date(frame, year_col, month_col)
    else:
        df["Date"] = pd.NaT
    rows = []
    for _, row in df.iterrows():
        rows.append({"Factory": row["Factory"], "Date": row["Date"], "Status": "Closed", "Count": int(row["Closed"])})
        rows.append({"Factory": row["Factory"], "Date": row["Date"], "Status": "Past Due", "Count": int(row["Past Due"])})
        rows.append({"Factory": row["Factory"], "Date": row["Date"], "Status": "Open", "Count": int(row["Open"])})
    return pd.DataFrame(rows)


def build_workers_comp_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    amount_col = find_column(frame, ["Money_Spent_This_Month", "Amount"])
    if not all([factory_col, year_col, month_col, amount_col]):
        raise ValueError("Workers comp sheet is missing required columns.")
    df = frame[[factory_col, year_col, month_col, amount_col]].copy()
    df.columns = ["Factory", "Year", "Month", "Amount"]
    df["Factory"] = df["Factory"].map(normalize_site)
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Date"] = year_month_date(df, "Year", "Month")
    return df[["Factory", "Date", "Amount"]]


def build_dashboard_goals_df(frame: pd.DataFrame) -> pd.DataFrame:
    year_col = find_column(frame, ["Year"])
    metric_col = find_column(frame, ["Metric"])
    goal_col = find_column(frame, ["Goal", "Value"])
    if not all([year_col, metric_col, goal_col]):
        raise ValueError("Dashboard goals sheet is missing required columns.")
    df = frame[[year_col, metric_col, goal_col]].copy()
    df.columns = ["Year", "Metric", "Goal"]
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce").fillna(0).astype(int)
    df["Metric"] = df["Metric"].fillna("").astype(str).str.strip()
    df["Goal"] = coerce_numeric(df["Goal"]).fillna(0)
    return df[df["Metric"] != ""]


def build_risk_reduction_goals_df(frame: pd.DataFrame) -> pd.DataFrame:
    """Normalize annual site goals from the Ergo_Risk_Reduction sheet."""

    site_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    goal_col = find_column(
        frame,
        ["Risk_Reduction_Goals", "Risk Reduction Goals", "Annual Goal", "Goal"],
    )
    if not all([site_col, year_col, goal_col]):
        raise ValueError(
            "Ergo_Risk_Reduction sheet requires Site, Year, and "
            "Risk_Reduction_Goals columns."
        )
    df = frame[[site_col, year_col, goal_col]].copy()
    df.columns = ["Factory", "Year", "Goal"]
    df["Factory"] = df["Factory"].map(normalize_site)
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Goal"] = coerce_numeric(df["Goal"])
    df = df.dropna(subset=["Year", "Goal"])
    df["Year"] = df["Year"].astype(int)
    return df[df["Factory"] != "Other"]


def build_npi_metrics_df(frame: pd.DataFrame) -> pd.DataFrame:
    program_col = find_column(frame, ["NPI_Name", "NPI", "Program"])
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    date_col = find_column(frame, ["Completion_Time", "Completion Time", "Start_Time", "Start Time", "Date"])
    actual_eau_col = find_column(frame, ["Future_EAU", "Future_EAU_Volume", "EAU Actual", "EAU Current"])
    target_eau_col = find_column(frame, ["Baseline_EAU", "Baseline_EAU-Volume", "EAU Target", "EAU Baseline"])
    actual_headcount_col = find_column(frame, ["Future_Headcount", "Headcount Actual"])
    target_headcount_col = find_column(frame, ["Baseline_Headcount", "Headcount Target"])
    actual_part_col = find_column(frame, ["Future_Part_Count", "Part Count Actual"])
    target_part_col = find_column(frame, ["Baseline_Part_Count", "Part Count Target"])
    fishbone_col = find_column(frame, ["Fishbone_Status", "Fishbone Status"])
    paag_col = find_column(frame, ["PAAG_Status", "At-a-Glance_Status", "At-a-Glance Status", "PAAG Status"])
    pin_map_col = find_column(frame, ["Pin_Map_Status", "Pin Map Status"])
    wins_col = find_column(frame, ["AIE_Wins_Opportunities", "AIE Wins Opportunities"])
    if not all([program_col, actual_headcount_col, target_headcount_col]):
        raise ValueError("NPI metrics sheet is missing required columns.")
    df = pd.DataFrame(
        {
            "Program": frame[program_col],
            "Factory": frame[factory_col] if factory_col else pd.Series("Other", index=frame.index),
            "EAU Actual": frame[actual_eau_col] if actual_eau_col else pd.Series(float("nan"), index=frame.index),
            "EAU Target": frame[target_eau_col] if target_eau_col else pd.Series(float("nan"), index=frame.index),
            "Headcount Actual": frame[actual_headcount_col],
            "Headcount Target": frame[target_headcount_col],
            "Part Count Actual": frame[actual_part_col] if actual_part_col else pd.Series(float("nan"), index=frame.index),
            "Part Count Target": frame[target_part_col] if target_part_col else pd.Series(float("nan"), index=frame.index),
            "Fishbone Status": frame[fishbone_col] if fishbone_col else pd.Series("", index=frame.index),
            "PAAG Status": frame[paag_col] if paag_col else pd.Series("", index=frame.index),
            "Pin Map Status": frame[pin_map_col] if pin_map_col else pd.Series("", index=frame.index),
            "AIE Wins Opportunities": frame[wins_col] if wins_col else pd.Series("", index=frame.index),
        },
        index=frame.index,
    )
    if date_col:
        df["Submission Date"] = pd.to_datetime(frame[date_col], errors="coerce")
        df["Date"] = npi_reporting_month(df["Submission Date"])
    elif year_col and month_col:
        df["Date"] = year_month_date(frame, year_col, month_col)
    else:
        df["Date"] = pd.Series(pd.NaT, index=frame.index, dtype="datetime64[ns]")
    df["Program"] = df["Program"].fillna("").astype(str).str.strip()
    df["Factory"] = df["Factory"].map(normalize_site)
    for col in ["EAU Actual", "EAU Target", "Headcount Actual", "Headcount Target", "Part Count Actual", "Part Count Target"]:
        df[col] = coerce_numeric(df[col])
    for col in ["Fishbone Status", "PAAG Status", "Pin Map Status", "AIE Wins Opportunities"]:
        df[col] = df[col].fillna("").astype(str)
    return df[df["Program"] != ""].reset_index(drop=True)


def build_npi_ergo_reds_df(frame: pd.DataFrame) -> pd.DataFrame:
    program_col = find_column(frame, ["NPI_Name", "NPI", "Program"])
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    date_col = find_column(frame, ["Completion_Time", "Completion Time", "Start_Time", "Start Time", "Date"])
    started_col = find_column(frame, ["Started_Ergo_Assessments", "Started"])
    open_col = find_column(frame, ["Open_Ergo_Reds", "Open_Ergo_Assessments", "Open"])
    pending_col = find_column(frame, ["Pending_Ergo_Reds", "Pending_Ergo_Assessments", "Pending"])
    validation_col = find_column(frame, ["Ergo_Assessments_Needing_Final_Validation", "Validation_Ergo_Assessments", "Validation"])
    closed_admin_col = find_column(frame, ["Ergo_Greens_Admin_solutions", "Closed_Ergo_Assessments_Admin", "Closed Administrative"])
    closed_engineering_col = find_column(frame, ["Ergo_Greens_Engineering_solutions", "Closed_Ergo_Assessments_Engineering", "Closed Engineering"])
    wins_col = find_column(frame, ["Ergo_Wins_Opportunities", "Ergo Wins Opportunities"])
    if not all([program_col, started_col, open_col, pending_col, validation_col, closed_admin_col, closed_engineering_col]):
        raise ValueError("NPI ergo reds sheet is missing required columns.")
    df = pd.DataFrame(
        {
            "Program": frame[program_col],
            "Factory": frame[factory_col] if factory_col else pd.Series("Other", index=frame.index),
            "Started": frame[started_col],
            "Open": frame[open_col],
            "Pending": frame[pending_col],
            "Validation": frame[validation_col],
            "Closed Administrative": frame[closed_admin_col],
            "Closed Engineering": frame[closed_engineering_col],
            "Ergo Wins Opportunities": frame[wins_col] if wins_col else pd.Series("", index=frame.index),
        },
        index=frame.index,
    )
    if date_col:
        df["Submission Date"] = pd.to_datetime(frame[date_col], errors="coerce")
        df["Date"] = npi_reporting_month(df["Submission Date"])
    elif year_col and month_col:
        df["Date"] = year_month_date(frame, year_col, month_col)
    else:
        df["Date"] = pd.Series(pd.NaT, index=frame.index, dtype="datetime64[ns]")
    df["Program"] = df["Program"].fillna("").astype(str).str.strip()
    df["Factory"] = df["Factory"].map(normalize_site)
    for col in ["Started", "Open", "Pending", "Validation", "Closed Administrative", "Closed Engineering"]:
        df[col] = coerce_numeric(df[col]).fillna(0).astype(int)
    df["Ergo Wins Opportunities"] = df["Ergo Wins Opportunities"].fillna("").astype(str)
    rows = []
    for _, row in df[df["Program"] != ""].iterrows():
        for status_col in ["Started", "Open", "Pending", "Validation", "Closed Administrative", "Closed Engineering"]:
            rows.append(
                {
                    "Program": row["Program"],
                    "Factory": row["Factory"],
                    "Date": row["Date"],
                    "Submission Date": row.get("Submission Date", pd.NaT),
                    "Status": status_col,
                    "Count": int(row[status_col]),
                    "Ergo Wins Opportunities": row["Ergo Wins Opportunities"],
                }
            )
    return pd.DataFrame(rows)


def apply_factory_date_filter(
    frame: pd.DataFrame,
    selected_factories: list[str],
    start_date: pd.Timestamp | None,
    end_date: pd.Timestamp | None,
    *,
    include_undated: bool = False,
) -> pd.DataFrame:
    if frame.empty:
        return frame
    result = frame.copy()
    if selected_factories:
        result = result[result["Factory"].isin(selected_factories)]
    if "Date" in result.columns and start_date is not None and end_date is not None:
        range_start = pd.Timestamp(start_date).normalize()
        range_end_exclusive = pd.Timestamp(end_date).normalize() + pd.Timedelta(days=1)
        dates = pd.to_datetime(result["Date"], errors="coerce")
        in_range = (dates >= range_start) & (dates < range_end_exclusive)
        if include_undated:
            in_range = in_range | dates.isna()
        result = result[in_range]
    return result


def apply_scope_date_filter(
    frame: pd.DataFrame,
    scope: str,
    start_date: pd.Timestamp | None,
    end_date: pd.Timestamp | None,
) -> pd.DataFrame:
    if frame.empty:
        return frame
    result = frame.copy()
    if "Factory" in result:
        if scope == "GE Appliances":
            result = result[result["Factory"] != "GEA"]
        elif scope == "Supply Chain":
            result = result[~result["Factory"].isin({"GEA", "BODE"})]
        else:
            result = result[result["Factory"] == scope]
    if "Date" in result and start_date is not None and end_date is not None:
        result = result[
            (result["Date"] >= start_date) & (result["Date"] <= end_date)
        ]
    return result


def apply_scope_ytd_filter(
    frame: pd.DataFrame,
    scope: str,
    end_date: pd.Timestamp | None,
) -> pd.DataFrame:
    """Filter site data from January through the selected reporting month."""

    if end_date is None:
        return apply_scope_date_filter(frame, scope, None, None)
    reporting_end = pd.Timestamp(end_date).to_period("M").end_time.normalize()
    reporting_start = pd.Timestamp(year=reporting_end.year, month=1, day=1)
    return apply_scope_date_filter(
        frame,
        scope,
        reporting_start,
        reporting_end,
    )


def default_site_risk_reduction_goals(
    frame: pd.DataFrame,
    sites: list[str],
    year: int,
) -> dict[str, float]:
    """Return one annual goal per site from the goals worksheet."""

    defaults = {site: 0.0 for site in sites}
    if (
        frame.empty
        or "Factory" not in frame
        or "Year" not in frame
        or "Goal" not in frame
    ):
        return defaults
    yearly = frame[frame["Year"] == year].copy()
    if yearly.empty:
        return defaults
    workbook_goals = yearly.groupby("Factory")["Goal"].max().to_dict()
    return {
        site: float(workbook_goals.get(site, defaults[site]))
        for site in sites
    }


def risk_reduction_goal_for_scope(
    site_goals: dict[str, float],
    scope: str,
) -> float:
    """Roll site goals into the selected site or aggregate reporting scope."""

    if scope == "GE Appliances":
        included_sites = site_goals
    elif scope == "Supply Chain":
        included_sites = {
            site: goal for site, goal in site_goals.items() if site != "BODE"
        }
    else:
        included_sites = {scope: site_goals.get(scope, 0.0)}
    return float(sum(included_sites.values()))


def default_combined_pdf_sites(scopes: list[str]) -> list[str]:
    """Select every factory and aggregate scope except Other by default."""

    return [scope for scope in scopes if scope != "Other"]


def apply_program_filter(frame: pd.DataFrame, selected_programs: list[str]) -> pd.DataFrame:
    if frame.empty or not selected_programs:
        return frame
    return frame[frame["Program"].isin(selected_programs)]


def available_npi_programs(*frames: pd.DataFrame) -> list[str]:
    programs: set[str] = set()
    for frame in frames:
        if frame.empty or "Program" not in frame:
            continue
        for value in frame["Program"].dropna().astype(str):
            program = value.strip()
            if program and program.lower() not in {"nan", "none"}:
                programs.add(program)
    return sorted(programs)


def has_npi_values(frame: pd.DataFrame, *columns: str) -> bool:
    return bool(
        not frame.empty
        and columns
        and all(column in frame and frame[column].notna().any() for column in columns)
    )


def latest_npi_snapshots(frame: pd.DataFrame) -> pd.DataFrame:
    """Keep every row belonging to the latest submission for each NPI program."""

    if frame.empty or "Program" not in frame:
        return frame
    result = frame.copy()
    snapshot_time = (
        pd.to_datetime(result["Submission Date"], errors="coerce")
        if "Submission Date" in result
        else pd.Series(pd.NaT, index=result.index, dtype="datetime64[ns]")
    )
    if "Date" in result:
        snapshot_time = snapshot_time.fillna(
            pd.to_datetime(result["Date"], errors="coerce")
        )
    result["_snapshot_time"] = snapshot_time
    dated_rows = result["_snapshot_time"].notna()
    if dated_rows.any():
        latest_time = result.groupby("Program")["_snapshot_time"].transform("max")
        result = result[~dated_rows | result["_snapshot_time"].eq(latest_time)]
    return result.drop(columns="_snapshot_time")


def normalize_date_selection(
    selection: object,
    fallback_start: pd.Timestamp,
    fallback_end: pd.Timestamp,
) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Normalize Streamlit's scalar and transient tuple date-input states."""

    if isinstance(selection, (tuple, list)):
        selected_dates = [value for value in selection if value is not None]
        if len(selected_dates) >= 2:
            return pd.Timestamp(selected_dates[0]), pd.Timestamp(selected_dates[1])
        if len(selected_dates) == 1:
            selected = pd.Timestamp(selected_dates[0])
            return selected, selected
        return pd.Timestamp(fallback_start), pd.Timestamp(fallback_end)
    selected = pd.Timestamp(selection)
    return selected, selected


def previous_calendar_month(
    reference_date: object | None = None,
) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Return the full calendar month immediately before the reference date."""

    reference = (
        pd.Timestamp.today().normalize()
        if reference_date is None
        else pd.Timestamp(reference_date).normalize()
    )
    current_month_start = reference.to_period("M").start_time
    previous_month_start = current_month_start - pd.DateOffset(months=1)
    previous_month_end = current_month_start - pd.Timedelta(days=1)
    return previous_month_start, previous_month_end


def filter_injury_history(
    frame: pd.DataFrame,
    end_date: pd.Timestamp | None,
    months: int = 12,
) -> pd.DataFrame:
    """Keep trailing monthly injury history through the selected end month."""

    if frame.empty or end_date is None or "Date" not in frame:
        return frame
    if months < 1:
        raise ValueError("months must be at least 1")
    history_end = pd.Timestamp(end_date).to_period("M").end_time.normalize()
    history_start = (
        history_end.to_period("M").start_time - pd.DateOffset(months=months - 1)
    )
    return frame[
        (frame["Date"] >= history_start) & (frame["Date"] <= history_end)
    ]


def filter_through_end_month(
    frame: pd.DataFrame,
    end_date: pd.Timestamp | None,
) -> pd.DataFrame:
    """Keep historical rows needed for current and prior-year injury metrics."""

    if frame.empty or end_date is None or "Date" not in frame:
        return frame
    history_end = pd.Timestamp(end_date).to_period("M").end_time.normalize()
    return frame[frame["Date"] <= history_end]


def make_injury_chart(frame: pd.DataFrame) -> go.Figure:
    return make_production_trend_chart(frame)


def below_chart_legend(y: float = -0.24) -> dict[str, object]:
    """Return a consistent horizontal legend positioned below the plot area."""

    return {
        "orientation": "h",
        "yanchor": "top",
        "y": y,
        "xanchor": "center",
        "x": 0.5,
    }


def make_assessment_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    snapshot = frame.copy()
    if "Date" in snapshot and snapshot["Date"].notna().any():
        latest_by_factory = snapshot.groupby("Factory")["Date"].transform("max")
        snapshot = snapshot[snapshot["Date"].eq(latest_by_factory)]
    counts = snapshot.groupby("Risk Level")["Count"].sum().reset_index()
    fig = px.pie(
        counts,
        values="Count",
        names="Risk Level",
        color="Risk Level",
        color_discrete_map={
            "Critical": "#6B1E5B",
            "Red": "#E45756",
            "Yellow": "#F58518",
            "Green": "#54A24B",
        },
    )
    fig.update_layout(
        title="Ergo Assessment Breakdown",
        height=330,
        margin=dict(l=16, r=16, t=44, b=76),
        legend=below_chart_legend(-0.12),
    )
    return fig


def make_npi_headcount_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    summary = frame.groupby("Program", as_index=False)[["Headcount Actual", "Headcount Target"]].sum()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=summary["Program"], y=summary["Headcount Actual"], name="Actual", marker_color="#4C78A8"))
    fig.add_trace(go.Bar(x=summary["Program"], y=summary["Headcount Target"], name="Target", marker_color="#F58518"))
    fig.update_layout(
        title="Headcount vs Target",
        barmode="group",
        height=330,
        margin=dict(l=16, r=16, t=44, b=82),
        xaxis_title="Program",
        yaxis_title="Headcount",
        legend=below_chart_legend(),
    )
    return fig


def make_npi_reds_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    summary = frame.groupby(["Program", "Status"])["Count"].sum().reset_index()
    fig = px.bar(summary, x="Program", y="Count", color="Status", barmode="stack")
    fig.update_layout(
        title="Ergo Reds Status by Program",
        height=340,
        margin=dict(l=16, r=16, t=44, b=90),
        xaxis_title="Program",
        yaxis_title="Count",
        legend=below_chart_legend(),
    )
    return fig


def make_npi_comparison_chart(
    frame: pd.DataFrame,
    current_col: str,
    future_col: str,
    title: str,
    yaxis_title: str,
    current_name: str = "Current",
    future_name: str = "Future",
    current_color: str = "#B7C9AE",
    future_color: str = "#5E8F39",
) -> go.Figure:
    if frame.empty:
        return go.Figure()
    summary = frame.groupby("Program", as_index=False)[[current_col, future_col]].sum()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=summary["Program"], y=summary[current_col], name=current_name, marker_color=current_color))
    fig.add_trace(go.Bar(x=summary["Program"], y=summary[future_col], name=future_name, marker_color=future_color))
    fig.update_layout(
        title=title,
        barmode="group",
        height=330,
        margin=dict(l=16, r=16, t=52, b=82),
        xaxis_title="",
        yaxis_title=yaxis_title,
        legend=below_chart_legend(-0.28),
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E6EAF0", zeroline=False)
    return fig


def make_npi_detail_snapshot_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    order = ["Started", "Open", "Pending", "Validation", "Closed Administrative", "Closed Engineering"]
    display_labels = {
        "Started": "Started",
        "Open": "Open",
        "Pending": "Pending",
        "Validation": "Validation",
        "Closed Administrative": "Closed Admin",
        "Closed Engineering": "Closed Eng",
    }
    status_colors = {
        "Started": "#9CA3AF",
        "Open": "#D98F8F",
        "Pending": "#E6C26A",
        "Validation": "#3B82F6",
        "Closed Administrative": "#A7D7A9",
        "Closed Engineering": "#2E7D32",
    }
    summary = frame.groupby("Status", as_index=False)["Count"].sum()
    summary["Status"] = pd.Categorical(summary["Status"], categories=order, ordered=True)
    summary = summary.sort_values("Status")
    status_labels = summary["Status"].astype("object")
    fig = go.Figure(
        go.Bar(
            x=status_labels.map(display_labels).fillna(status_labels.astype(str)),
            y=summary["Count"],
            marker_color=status_labels.map(status_colors),
            text=summary["Count"],
            textposition="outside",
        )
    )
    fig.update_layout(
        title="Ergo Reds Snapshot",
        height=260,
        margin=dict(l=16, r=16, t=56, b=24),
        xaxis_title="",
        yaxis_title="Count",
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    fig.update_xaxes(tickangle=0, showgrid=False)
    max_count = float(summary["Count"].max()) if not summary.empty else 0
    fig.update_yaxes(range=[0, max_count * 1.2 if max_count else 1], gridcolor="#E6EAF0", zeroline=False)
    return fig


def make_production_trend_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    monthly = (
        frame.groupby(frame["Date"].dt.month)
        .agg(Injuries=("Injuries", "sum"), Recordable_Rate=("Recordable Rate", "mean"), Dart_Rate=("DART Rate", "mean"))
        .reindex(range(1, 13), fill_value=0)
        .reset_index()
        .rename(columns={"Date": "Month", "index": "Month"})
    )
    fig = go.Figure()
    fig.add_trace(go.Bar(x=monthly["Month"], y=monthly["Injuries"], name="2026 Ergo Recordables", marker_color="#2F6DB0"))
    fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Recordable_Rate"], name="RATE 2026 Recordables", mode="lines+markers", yaxis="y2", line=dict(color="#D91E18", width=2.5)))
    fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Dart_Rate"], name="RATE 2026 DART", mode="lines+markers", yaxis="y2", line=dict(color="#F4B400", width=2, dash="dash")))
    fig.update_layout(
        title="GE Appliances Ergo Trends",
        height=390,
        margin=dict(l=16, r=16, t=44, b=100),
        barmode="group",
        legend=below_chart_legend(),
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="",
        yaxis_title="Count",
        yaxis2=dict(title="Rate", overlaying="y", side="right", range=[0, max(monthly["Recordable_Rate"].max(), monthly["Dart_Rate"].max(), 1.2) * 1.2]),
    )
    fig.update_xaxes(tickmode="array", tickvals=list(range(1, 13)), ticktext=[str(i) for i in range(1, 13)], showgrid=True, gridcolor="#E6EAF0")
    fig.update_yaxes(gridcolor="#E6EAF0", zeroline=False)
    return fig


def get_goal_value(goals_frame: pd.DataFrame, year: int, metric_aliases: list[str], default: float = 0.0) -> float:
    if goals_frame.empty:
        return default
    normalized_aliases = {alias.strip().lower() for alias in metric_aliases}
    matched = goals_frame[
        (goals_frame["Year"] == year)
        & (goals_frame["Metric"].str.strip().str.lower().isin(normalized_aliases))
    ]
    if matched.empty:
        return default
    return float(matched.iloc[0]["Goal"])


def make_production_trend_chart(
    frame: pd.DataFrame,
    recordable_goal: float = 0.0,
    dart_goal: float = 0.0,
) -> go.Figure:
    if frame.empty:
        return go.Figure()
    monthly = frame.sort_values("Date").copy()
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=monthly["Date"],
            y=monthly["Ergo All Rate R12"],
            name="All Ergo Rate (R12)",
            mode="lines+markers",
            line=dict(color="#2F6DB0", width=2.5),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=monthly["Date"],
            y=monthly["Ergo Recordable Rate R12"],
            name="Ergo Recordable Rate (R12)",
            mode="lines+markers",
            line=dict(color="#D91E18", width=2.2),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=monthly["Date"],
            y=monthly["EHS Recordable Rate R12"],
            name="Total EHS Recordable Rate (R12)",
            mode="lines",
            line=dict(color="#5E8F39", width=2, dash="dash"),
        )
    )
    if recordable_goal:
        fig.add_trace(go.Scatter(x=monthly["Date"], y=[recordable_goal] * len(monthly), name="Recordable Goal", mode="lines", line=dict(color="#8B5CF6", width=2, dash="dot")))
    if dart_goal:
        fig.add_trace(go.Scatter(x=monthly["Date"], y=[dart_goal] * len(monthly), name="DART Goal", mode="lines", line=dict(color="#F4B400", width=2, dash="dot")))
    rate_max = pd.concat(
        [
            monthly["Ergo All Rate R12"],
            monthly["Ergo Recordable Rate R12"],
            monthly["EHS Recordable Rate R12"],
        ]
    ).max(skipna=True)
    if pd.isna(rate_max):
        rate_max = 1.0
    fig.update_layout(
        title="Rolling-12 Injury Trends",
        height=390,
        margin=dict(l=16, r=16, t=44, b=100),
        legend=below_chart_legend(),
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="",
        yaxis=dict(
            title="Cases per 200,000 hours",
            range=[0, max(float(rate_max), recordable_goal, dart_goal, 1.0) * 1.2],
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#E6EAF0")
    fig.update_yaxes(gridcolor="#E6EAF0", zeroline=False)
    return fig


def make_ergo_classification_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    monthly = frame.sort_values("Date")
    fig = go.Figure()
    series = [
        ("Ergo First Aid", "First Aid", "#9BB7D4"),
        ("Ergo DAFW", "DAFW", "#D91E18"),
        ("Ergo Restricted", "Restricted", "#F4B400"),
        ("Ergo Other Recordable", "Other Recordable", "#5E8F39"),
    ]
    for column, label, color in series:
        fig.add_trace(
            go.Bar(x=monthly["Date"], y=monthly[column], name=label, marker_color=color)
        )
    fig.update_layout(
        title="Monthly Ergo Cases by Classification",
        barmode="stack",
        height=360,
        margin=dict(l=16, r=16, t=44, b=90),
        xaxis_title="",
        yaxis_title="Cases",
        legend=below_chart_legend(),
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    fig.update_yaxes(gridcolor="#E6EAF0", zeroline=False)
    return fig


def make_production_assessment_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    snapshot = frame.copy()
    if "Date" in snapshot and snapshot["Date"].notna().any():
        latest_by_factory = snapshot.groupby("Factory")["Date"].transform("max")
        snapshot = snapshot[snapshot["Date"].eq(latest_by_factory)]
    counts = snapshot.groupby("Risk Level", as_index=False)["Count"].sum()
    fig = px.pie(
        counts,
        values="Count",
        names="Risk Level",
        color="Risk Level",
        color_discrete_map={
            "Critical": "#6B1E5B",
            "Red": "#D92D20",
            "Yellow": "#F4B400",
            "Green": "#69C466",
        },
    )
    fig.update_traces(textposition="inside", textinfo="percent")
    fig.update_layout(title="Ergo Assessments", height=255, margin=dict(l=0, r=0, t=32, b=0), showlegend=False, paper_bgcolor="white")
    return fig


def make_pdf_assessment_chart(frame: pd.DataFrame) -> go.Figure:
    """Create the compact assessment breakdown used on PDF factory pages."""

    fig = make_production_assessment_chart(frame)
    if not fig.data:
        return fig
    fig.update_traces(
        hole=0,
        textinfo="percent",
        textposition="inside",
        textfont_size=11,
    )
    fig.update_layout(
        title=None,
        height=300,
        margin=dict(l=2, r=2, t=2, b=2),
        showlegend=False,
    )
    return fig


def make_production_projects_chart(actual_points: float, goal_points: float, year: int) -> go.Figure:
    if goal_points <= 0 and actual_points <= 0:
        return go.Figure()
    capped_actual = min(actual_points, goal_points) if goal_points > 0 else actual_points
    remaining_points = max(goal_points - capped_actual, 0)
    yaxis_top = max(goal_points, actual_points, 1)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=[str(year)], y=[capped_actual], name="Actual YTD", marker_color="#4D74BE"))
    if remaining_points > 0:
        fig.add_trace(go.Bar(x=[str(year)], y=[remaining_points], name="Remaining to Goal", marker_color="#D8E1F0"))
    fig.update_layout(
        title="Risk Reduction Progress",
        height=255,
        margin=dict(l=16, r=16, t=42, b=70),
        paper_bgcolor="white",
        plot_bgcolor="white",
        barmode="stack",
        xaxis_title="",
        yaxis_title="Points",
        legend=below_chart_legend(-0.18),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(range=[0, yaxis_top], gridcolor="#E6EAF0", zeroline=False)
    return fig


def render_npi_style() -> None:
    st.markdown(
        """
        <style>
        .npi-header {
            padding: 0.2rem 0 0.75rem 0;
            border-bottom: 2px solid #111827;
            margin-bottom: 0.9rem;
        }
        .npi-title {
            font-size: 2rem;
            font-weight: 700;
            line-height: 1.05;
            color: #111827;
        }
        .npi-subtitle {
            font-size: 1rem;
            font-weight: 600;
            color: #1f2937;
            margin-top: 0.15rem;
        }
        .npi-meta {
            font-size: 0.9rem;
            color: #475569;
            margin-top: 0.15rem;
        }
        .summary-card {
            border-radius: 12px;
            padding: 0.8rem 0.9rem;
            min-height: 92px;
            box-shadow: inset 0 0 0 1px rgba(148, 163, 184, 0.22);
        }
        .summary-label {
            font-size: 0.9rem;
            font-weight: 700;
            color: #111827;
            margin-bottom: 0.15rem;
        }
        .summary-value {
            font-size: 1.95rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.05;
        }
        .summary-context {
            font-size: 0.82rem;
            color: #475569;
            margin-top: 0.2rem;
        }
        .summary-card.progress { background: #e7ebf1; }
        .summary-card.green { background: #ddf5df; }
        .summary-card.red { background: #f8dddd; }
        .summary-card.yellow { background: #fff0c9; }
        .summary-card.neutral { background: #eef2f7; }
        .status-title {
            text-align: center;
            font-size: 0.85rem;
            font-weight: 700;
            color: #475569;
            margin-bottom: 0.4rem;
        }
        .status-orb {
            width: 124px;
            height: 124px;
            border-radius: 999px;
            margin: 0 auto 0.45rem auto;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
            padding: 1rem;
            font-size: 0.86rem;
            font-weight: 700;
            color: #1f2937;
            border: 1px solid rgba(148, 163, 184, 0.28);
        }
        .status-finished { background: #9ba8ab; color: white; }
        .status-progress { background: linear-gradient(90deg, #e9edf2 50%, #9ba8ab 50%); }
        .status-notstarted { background: #eef2f6; }
        .wins-card {
            border-radius: 10px;
            padding: 0.8rem 0.9rem;
            background: #fbfcfe;
            border: 1px solid #d8dee8;
            min-height: 108px;
        }
        .wins-title {
            font-weight: 700;
            color: #111827;
            margin-bottom: 0.3rem;
        }
        .wins-body {
            color: #334155;
            font-size: 0.88rem;
            line-height: 1.4;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_production_style() -> None:
    st.markdown(
        """
        <style>
        .prod-panel-title {
            font-size: 1.1rem;
            font-weight: 700;
            color: #111827;
            margin-bottom: 0.35rem;
        }
        .prod-metric-box {
            border: 1px solid #d8dee8;
            background: #fbfcfe;
            border-radius: 10px;
            padding: 0.55rem 0.75rem;
            text-align: center;
        }
        .prod-metric-value {
            font-size: 1.9rem;
            font-weight: 800;
            line-height: 1.05;
        }
        .prod-metric-red { background: #f7d8da; }
        .prod-metric-critical { background: #ead7e7; color: #551947; }
        .prod-metric-green { background: #d8f1d8; }
        .prod-metric-yellow { background: #fff0c9; }
        .prod-metric-neutral { background: #eef2f7; }
        .prod-kicker {
            font-size: 0.78rem;
            color: #475569;
            margin-top: 0.18rem;
        }
        .goal-card, .prod-text-card {
            border-radius: 10px;
            border: 1px solid #d8dee8;
            background: #fbfcfe;
            padding: 0.75rem 0.85rem;
            min-height: 250px;
        }
        .goal-title, .prod-text-title {
            font-weight: 700;
            color: #111827;
            text-align: center;
            margin-bottom: 0.45rem;
        }
        .goal-year {
            font-size: 2.25rem;
            font-weight: 800;
            text-align: center;
            margin-bottom: 0.35rem;
        }
        .goal-row {
            display: flex;
            justify-content: space-between;
            gap: 1rem;
            padding: 0.42rem 0;
            border-top: 1px solid #e6eaf0;
            font-size: 0.92rem;
        }
        .goal-row:first-of-type {
            border-top: 0;
        }
        .goal-value-highlight {
            background: #fff34a;
            padding: 0.1rem 0.35rem;
            font-weight: 800;
        }
        .goal-value-bad {
            color: #b42318;
            font-weight: 700;
        }
        .goal-summary-card {
            border-radius: 10px;
            border: 1px solid #d8dee8;
            background: #fbfcfe;
            padding: 0.5rem 0.65rem;
            margin-top: 0.15rem;
        }
        .goal-summary-title {
            font-weight: 700;
            color: #111827;
            text-align: center;
            margin-bottom: 0.35rem;
        }
        .goal-summary-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.3rem;
        }
        .goal-summary-item {
            border: 1px solid #e6eaf0;
            border-radius: 7px;
            padding: 0.3rem 0.4rem;
            text-align: center;
        }
        .goal-summary-label {
            color: #64748b;
            font-size: 0.7rem;
        }
        .goal-summary-value {
            color: #111827;
            font-size: 1rem;
            font-weight: 800;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def format_change(current_value: float, future_value: float) -> str:
    if not current_value:
        return "0.0%"
    return f"{((future_value - current_value) / current_value) * 100:+.1f}%"


def injury_trend_status(
    current_value: float,
    prior_value: float,
    has_comparison: bool = True,
) -> str:
    """Classify injury trends where a lower current value is favorable."""

    if (
        not has_comparison
        or pd.isna(current_value)
        or pd.isna(prior_value)
    ):
        return "neutral"
    if current_value < prior_value:
        return "green"
    if current_value > prior_value:
        return "red"
    return "yellow"


REDUCTION_CHART_COLORS = {
    "green": "#5E8F39",
    "red": "#D64545",
    "yellow": "#D9A600",
    "neutral": "#9CA3AF",
}


def reduction_status(current_value: float, future_value: float) -> str:
    """Classify a metric where reducing the future value is favorable."""

    if pd.isna(current_value) or pd.isna(future_value):
        return "neutral"
    if future_value < current_value:
        return "green"
    if future_value > current_value:
        return "red"
    return "yellow"


def open_ergo_reds_status(open_reds: float, has_ergo_data: bool = True) -> str:
    """Classify open ergo reds: zero is favorable; any open item is unfavorable."""

    if not has_ergo_data or pd.isna(open_reds):
        return "neutral"
    return "green" if float(open_reds) == 0 else "red"


def summarize_npi_dashboard(metrics_frame: pd.DataFrame, reds_frame: pd.DataFrame) -> dict[str, object]:
    current_eau = float(metrics_frame["EAU Target"].sum())
    future_eau = float(metrics_frame["EAU Actual"].sum())
    current_part = float(metrics_frame["Part Count Target"].sum())
    future_part = float(metrics_frame["Part Count Actual"].sum())
    current_headcount = float(metrics_frame["Headcount Target"].sum())
    future_headcount = float(metrics_frame["Headcount Actual"].sum())
    current_hpu = (current_headcount * 2080 / current_eau) if current_eau else 0.0
    future_hpu = (future_headcount * 2080 / future_eau) if future_eau else 0.0

    status_scores = {"Finished": 100, "In Progress": 55, "Not Started": 10}
    tool_statuses = pd.concat(
        [
            metrics_frame["Fishbone Status"],
            metrics_frame["PAAG Status"],
            metrics_frame["Pin Map Status"],
        ],
        ignore_index=True,
    )
    tool_progress = tool_statuses.map(lambda value: status_scores.get(str(value).strip(), 35)).mean() if not tool_statuses.empty else 0
    open_reds = float(reds_frame.loc[reds_frame["Status"] == "Open", "Count"].sum())
    total_reds = float(reds_frame["Count"].sum())
    closure_score = ((total_reds - open_reds) / total_reds * 100) if total_reds else 100
    hpu_score = min(max(((current_hpu - future_hpu) / current_hpu * 100) if current_hpu else 0, -100), 100)
    part_score = min(max(((current_part - future_part) / current_part * 100) if current_part else 0, -100), 100)
    overall_progress = max(0.0, min(100.0, (tool_progress * 0.45) + (closure_score * 0.15) + ((hpu_score + 100) / 2 * 0.2) + ((part_score + 100) / 2 * 0.2)))

    return {
        "current_eau": current_eau,
        "future_eau": future_eau,
        "current_part": current_part,
        "future_part": future_part,
        "current_headcount": current_headcount,
        "future_headcount": future_headcount,
        "current_hpu": current_hpu,
        "future_hpu": future_hpu,
        "part_change": format_change(current_part, future_part),
        "headcount_change": format_change(current_headcount, future_headcount),
        "hpu_change": format_change(current_hpu, future_hpu),
        "overall_progress": overall_progress,
        "open_reds": int(open_reds),
    }


def make_summary_card(label: str, value: str, context: str, theme: str) -> str:
    return f"""
    <div class="summary-card {theme}">
        <div class="summary-label">{label}</div>
        <div class="summary-value">{value}</div>
        <div class="summary-context">{context}</div>
    </div>
    """


def render_status_orb(label: str, status: str) -> None:
    status = status.strip()
    if not status or status.lower() in {"nan", "none"}:
        status = "N/A"
    normalized = status.strip().lower()
    css_class = "status-progress"
    if normalized == "finished":
        css_class = "status-finished"
    elif normalized == "not started":
        css_class = "status-notstarted"
    st.markdown(
        f"""
        <div class="status-title">{label}</div>
        <div class="status-orb {css_class}">{status}</div>
        """,
        unsafe_allow_html=True,
    )


def latest_month_label(metrics_frame: pd.DataFrame, reds_frame: pd.DataFrame) -> str:
    date_series = pd.concat([metrics_frame.get("Date", pd.Series(dtype="datetime64[ns]")), reds_frame.get("Date", pd.Series(dtype="datetime64[ns]"))], ignore_index=True).dropna()
    if date_series.empty:
        return "No reporting month"
    latest_date = pd.to_datetime(date_series).max()
    return latest_date.strftime("%Y-%m")


def collect_note_text(series: pd.Series, fallback: str) -> str:
    notes = []
    seen = set()
    for value in series.fillna("").astype(str):
        cleaned = re.sub(r"\s+", " ", value).strip(" -")
        if not cleaned or cleaned.lower() == "nan" or cleaned in seen:
            continue
        seen.add(cleaned)
        notes.append(cleaned)
    if not notes:
        return fallback
    return "<br>".join(notes[:3])


def summarize_production_dashboard(
    prod_frame: pd.DataFrame,
    assessment_frame: pd.DataFrame,
    project_frame: pd.DataFrame,
    medical_frame: pd.DataFrame,
    workers_comp_frame: pd.DataFrame,
    goals_frame: pd.DataFrame,
    risk_reduction_goal: float | None = None,
) -> dict[str, object]:
    assessment_snapshot = assessment_frame.copy()
    if not assessment_snapshot.empty and assessment_snapshot["Date"].notna().any():
        latest_by_factory = assessment_snapshot.groupby("Factory")["Date"].transform("max")
        assessment_snapshot = assessment_snapshot[
            assessment_snapshot["Date"].eq(latest_by_factory)
        ]
    assessment_counts = assessment_snapshot.groupby("Risk Level")["Count"].sum().to_dict()
    critical_count = int(assessment_counts.get("Critical", 0))
    red_count = int(assessment_counts.get("Red", 0))
    yellow_count = int(assessment_counts.get("Yellow", 0))
    green_count = int(assessment_counts.get("Green", 0))
    total_assessments = critical_count + red_count + yellow_count + green_count

    medical_snapshot = medical_frame.copy()
    if not medical_snapshot.empty and medical_snapshot["Date"].notna().any():
        latest_by_factory = medical_snapshot.groupby("Factory")["Date"].transform("max")
        medical_snapshot = medical_snapshot[
            medical_snapshot["Date"].eq(latest_by_factory)
        ]
    med_counts = medical_snapshot.groupby("Status")["Count"].sum().to_dict()
    closed_requests = int(med_counts.get("Closed", 0))
    past_due_requests = int(med_counts.get("Past Due", 0))
    open_requests = int(med_counts.get("Open", 0))
    total_medical = closed_requests + past_due_requests + open_requests

    latest_date = prod_frame["Date"].max() if not prod_frame.empty else pd.Timestamp.today().normalize()
    latest_year = latest_date.year if pd.notna(latest_date) else pd.Timestamp.today().year
    total_reduction = float(project_frame["Reduction Points"].sum())
    sheet_target_points = (
        float(project_frame.groupby("Factory")["Target Points"].max().sum())
        if not project_frame.empty and "Factory" in project_frame
        else 0.0
    )
    target_points = (
        float(risk_reduction_goal)
        if risk_reduction_goal is not None
        else get_goal_value(
            goals_frame,
            latest_year,
            ["risk_reduction_projects", "risk reduction projects", "risk_reduction_goal", "annual risk reduction goal"],
            default=sheet_target_points,
        )
    )
    progress_fraction = (latest_date.month / 12) if pd.notna(latest_date) else (pd.Timestamp.today().month / 12)
    expected_progress = target_points * progress_fraction
    completion_pct = (total_reduction / target_points * 100) if target_points else 0.0
    remaining_calendar_pct = max(0.0, (1 - progress_fraction) * 100)

    valid_rates = prod_frame.dropna(subset=["Ergo All Rate R12"]).sort_values("Date")
    latest_prod = valid_rates.iloc[-1] if not valid_rates.empty else None
    prior_prod = None
    if latest_prod is not None:
        prior_date = latest_prod["Date"] - pd.DateOffset(years=1)
        prior_rows = valid_rates[valid_rates["Date"] == prior_date]
        if not prior_rows.empty:
            prior_prod = prior_rows.iloc[-1]

    def metric(row, column: str) -> float:
        if row is None or pd.isna(row.get(column)):
            return 0.0
        return float(row[column])

    def has_metric(row, column: str) -> bool:
        return row is not None and pd.notna(row.get(column))

    current_recordable = metric(latest_prod, "Ergo Recordable Rate R12")
    current_dart = metric(latest_prod, "Ergo DART Rate R12")
    total_case_rate = metric(latest_prod, "Ergo All Rate R12")
    current_dafw = metric(latest_prod, "Ergo DAFW Rate R12")
    current_ehs_recordable = metric(latest_prod, "EHS Recordable Rate R12")
    current_ehs_dart = metric(latest_prod, "EHS DART Rate R12")
    ergo_share = metric(latest_prod, "Ergo Share of EHS Recordables")
    rolling_hours = metric(latest_prod, "Hours R12")
    prior_recordable = metric(prior_prod, "Ergo Recordable Rate R12")
    prior_dart = metric(prior_prod, "Ergo DART Rate R12")
    prior_total_case_rate = metric(prior_prod, "Ergo All Rate R12")
    prior_dafw = metric(prior_prod, "Ergo DAFW Rate R12")
    prior_ehs_recordable = metric(prior_prod, "EHS Recordable Rate R12")
    prior_ergo_share = metric(prior_prod, "Ergo Share of EHS Recordables")
    recordable_goal = get_goal_value(goals_frame, latest_year, ["recordable_rate", "recordable rate", "injury_trends_recordable_rate_goal"], default=0.0)
    dart_goal = get_goal_value(goals_frame, latest_year, ["dart_rate", "dart rate", "injury_trends_dart_rate_goal"], default=0.0)

    workers_by_year = pd.DataFrame()
    current_workers_comp = 0.0
    prior_workers_comp = 0.0
    workers_change = 0.0
    if not workers_comp_frame.empty:
        workers_by_year = workers_comp_frame.assign(Year=workers_comp_frame["Date"].dt.year).groupby("Year", as_index=False)["Amount"].sum().sort_values("Year")
        current_workers_comp = float(workers_by_year.iloc[-1]["Amount"])
        if len(workers_by_year) > 1:
            prior_workers_comp = float(workers_by_year.iloc[-2]["Amount"])
        workers_change = ((current_workers_comp - prior_workers_comp) / prior_workers_comp * 100) if prior_workers_comp else 0.0

    return {
        "critical_count": critical_count,
        "red_count": red_count,
        "yellow_count": yellow_count,
        "green_count": green_count,
        "total_assessments": total_assessments,
        "closed_requests": closed_requests,
        "past_due_requests": past_due_requests,
        "open_requests": open_requests,
        "total_medical": total_medical,
        "total_reduction": total_reduction,
        "target_points": target_points,
        "expected_progress": expected_progress,
        "completion_pct": completion_pct,
        "remaining_calendar_pct": remaining_calendar_pct,
        "current_recordable": current_recordable,
        "prior_recordable": prior_recordable,
        "recordable_change": format_change(prior_recordable, current_recordable),
        "recordable_goal": recordable_goal,
        "current_dart": current_dart,
        "prior_dart": prior_dart,
        "dart_change": format_change(prior_dart, current_dart),
        "dart_goal": dart_goal,
        "total_case_rate": total_case_rate,
        "prior_total_case_rate": prior_total_case_rate,
        "total_case_change": format_change(prior_total_case_rate, total_case_rate),
        "has_total_case_comparison": has_metric(latest_prod, "Ergo All Rate R12")
        and has_metric(prior_prod, "Ergo All Rate R12"),
        "current_dafw": current_dafw,
        "prior_dafw": prior_dafw,
        "dafw_change": format_change(prior_dafw, current_dafw),
        "current_ehs_recordable": current_ehs_recordable,
        "prior_ehs_recordable": prior_ehs_recordable,
        "ehs_recordable_change": format_change(
            prior_ehs_recordable,
            current_ehs_recordable,
        ),
        "has_recordable_comparison": has_metric(
            latest_prod, "Ergo Recordable Rate R12"
        ) and has_metric(prior_prod, "Ergo Recordable Rate R12"),
        "has_dart_comparison": has_metric(latest_prod, "Ergo DART Rate R12")
        and has_metric(prior_prod, "Ergo DART Rate R12"),
        "has_ehs_recordable_comparison": has_metric(
            latest_prod, "EHS Recordable Rate R12"
        ) and has_metric(prior_prod, "EHS Recordable Rate R12"),
        "current_ehs_dart": current_ehs_dart,
        "ergo_share": ergo_share,
        "prior_ergo_share": prior_ergo_share,
        "ergo_share_change": format_change(prior_ergo_share, ergo_share),
        "has_ergo_share_comparison": has_metric(
            latest_prod, "Ergo Share of EHS Recordables"
        ) and has_metric(prior_prod, "Ergo Share of EHS Recordables"),
        "rolling_hours": rolling_hours,
        "ergo_all_cases_r12": metric(latest_prod, "Ergo All Cases R12"),
        "ergo_first_aid_r12": metric(latest_prod, "Ergo First Aid R12"),
        "ergo_recordable_r12": metric(latest_prod, "Ergo Recordable R12"),
        "ergo_dafw_r12": metric(latest_prod, "Ergo DAFW R12"),
        "ergo_restricted_r12": metric(latest_prod, "Ergo Restricted R12"),
        "ergo_dart_r12": metric(latest_prod, "Ergo DART R12"),
        "ehs_recordables_r12": metric(latest_prod, "EHS Recordables R12"),
        "workers_by_year": workers_by_year,
        "current_workers_comp": current_workers_comp,
        "prior_workers_comp": prior_workers_comp,
        "workers_change": workers_change,
        "latest_date": latest_date,
        "latest_year": latest_year,
    }


def ensure_plotly_image_export_ready() -> None:
    importlib.invalidate_caches()
    try:
        import kaleido
        import plotly.io._kaleido as plotly_kaleido
    except ImportError as exc:
        raise RuntimeError(
            "PDF export requires Kaleido. Install it into this app's virtual environment with: "
            ".\\.venv\\Scripts\\python.exe -m pip install --upgrade kaleido"
        ) from exc

    plotly_kaleido.kaleido = kaleido
    plotly_kaleido._KALEIDO_AVAILABLE = None
    plotly_kaleido._KALEIDO_MAJOR = None
    plotly_kaleido.kaleido_available()


def render_pdf_report(
    data: Dict[str, pd.DataFrame],
    scope: str = "GE Appliances",
    start_date: pd.Timestamp | None = None,
    end_date: pd.Timestamp | None = None,
    site_risk_reduction_goals: dict[str, float] | None = None,
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Ergonomics and Industrial Engineering Dashboard", styles["Title"]))
    story.append(Paragraph(f"Production report — {scope}", styles["Heading2"]))
    story.append(
        Paragraph(
            f"Completed months through {end_date.strftime('%B %Y') if end_date is not None else 'latest available month'}",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 10))

    production_history = aggregate_scope(
        data.get("production_injuries", pd.DataFrame()), scope
    )
    production = filter_injury_history(production_history, end_date)
    production_summary = filter_through_end_month(production_history, end_date)
    assessments = apply_scope_date_filter(
        data.get("ergo_assessments", pd.DataFrame()), scope, start_date, end_date
    )
    risk_projects = apply_scope_ytd_filter(
        data.get("risk_projects", pd.DataFrame()), scope, end_date
    )
    medical_requests = apply_scope_date_filter(
        data.get("medical_requests", pd.DataFrame()), scope, start_date, end_date
    )
    workers_comp = apply_scope_date_filter(
        data.get("workers_comp", pd.DataFrame()), scope, start_date, end_date
    )
    summary = summarize_production_dashboard(
        production_summary,
        assessments,
        risk_projects,
        medical_requests,
        workers_comp,
        data.get("dashboard_goals", pd.DataFrame()),
        risk_reduction_goal=(
            risk_reduction_goal_for_scope(site_risk_reduction_goals, scope)
            if site_risk_reduction_goals is not None
            else None
        ),
    )

    def add_chart(fig, title: str):
        ensure_plotly_image_export_ready()
        img_buffer = io.BytesIO()
        fig.write_image(img_buffer, format="png", scale=2)
        img_buffer.seek(0)
        story.append(Paragraph(title, styles["Heading3"]))
        story.append(Spacer(1, 8))
        story.append(Image(img_buffer, width=500, height=280))
        story.append(Spacer(1, 12))

    if not production.empty:
        kpi_rows = [
            ["Rolling-12 metric", "Value"],
            ["Hours worked", f"{summary['rolling_hours']:,.0f}"],
            ["All ergo rate", f"{summary['total_case_rate']:.2f}"],
            ["Ergo recordable rate", f"{summary['current_recordable']:.2f}"],
            ["Ergo DART rate", f"{summary['current_dart']:.2f}"],
            ["Ergo DAFW rate", f"{summary['current_dafw']:.2f}"],
            ["Total EHS recordable rate", f"{summary['current_ehs_recordable']:.2f}"],
            ["Ergo share of EHS recordables", f"{summary['ergo_share']:.1f}%"],
        ]
        kpi_table = Table(kpi_rows, colWidths=[330, 140])
        kpi_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F6DB0")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D8DEE8")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")]),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(kpi_table)
        story.append(Spacer(1, 12))
        add_chart(
            make_production_trend_chart(
                production,
                recordable_goal=summary["recordable_goal"],
                dart_goal=summary["dart_goal"],
            ),
            "Rolling-12 Injury Trends",
        )
        ensure_plotly_image_export_ready()
        classification_chart = _snapshot_chart(
            make_ergo_classification_chart(production),
            326,
            170,
            820,
            420,
        )
        risk_reduction = _snapshot_risk_reduction_block(summary)
        classification_and_risk = Table(
            [[classification_chart, risk_reduction]],
            colWidths=[330, 210],
        )
        classification_and_risk.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )
        story.append(classification_and_risk)
        story.append(Spacer(1, 12))
    if not assessments.empty:
        add_chart(make_assessment_chart(assessments), "Production - Ergo Assessment Breakdown")
    if not medical_requests.empty:
        story.append(Paragraph("Production - Medical Requests Status", styles["Heading3"]))
        story.append(Spacer(1, 6))
        pivot = pd.pivot_table(medical_requests, values="Count", index="Factory", columns="Status", aggfunc="sum", fill_value=0)
        table_data = [["Factory", *pivot.columns.tolist()]]
        table_data.extend([[index, *row.tolist()] for index, row in pivot.iterrows()])
        medical_table = Table(table_data, repeatRows=1)
        medical_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E6EAF0")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
                    ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ]
            )
        )
        story.append(medical_table)
        story.append(Spacer(1, 12))
    story.append(
        Paragraph(
            "Rate methodology: rolling-12 cases × 200,000 ÷ rolling-12 hours. "
            "DART counts each case once when DAFW or restricted. All Ergo counts "
            "each unique ergonomic case once, including first aid. The ergo share "
            "uses ergo recordables ÷ total EHS recordables.",
            styles["BodyText"],
        )
    )
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def _snapshot_chart(
    figure: go.Figure,
    width: float,
    height: float,
    pixel_width: int = 1000,
    pixel_height: int = 420,
) -> Image:
    image_buffer = io.BytesIO()
    figure.write_image(
        image_buffer,
        format="png",
        width=pixel_width,
        height=pixel_height,
        scale=2,
    )
    image_buffer.seek(0)
    return Image(image_buffer, width=width, height=height)


PDF_CARD_BORDER = colors.HexColor("#D8DEE8")
PDF_PAGE_BACKGROUND = colors.HexColor("#F4F7FA")
PDF_TEXT = colors.HexColor("#111827")
PDF_MUTED_TEXT = colors.HexColor("#475569")
PDF_BLUE = colors.HexColor("#315B7D")
PDF_GREEN = colors.HexColor("#DDF5DF")
PDF_RED = colors.HexColor("#F8DDDD")
PDF_YELLOW = colors.HexColor("#FFF0C9")
PDF_NEUTRAL = colors.HexColor("#EEF2F7")


def _snapshot_card(
    content: object,
    width: float,
    *,
    background: colors.Color = colors.white,
    padding: float = 5,
) -> Table:
    table = Table([[content]], colWidths=[width])
    table.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.65, PDF_CARD_BORDER),
                ("BACKGROUND", (0, 0), (-1, -1), background),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), padding),
                ("RIGHTPADDING", (0, 0), (-1, -1), padding),
                ("TOPPADDING", (0, 0), (-1, -1), padding),
                ("BOTTOMPADDING", (0, 0), (-1, -1), padding),
            ]
        )
    )
    return table


def _snapshot_card_row(
    cards: list[object],
    widths: list[float],
    *,
    gap: float = 6,
) -> Table:
    row: list[object] = []
    column_widths: list[float] = []
    for index, (card, width) in enumerate(zip(cards, widths)):
        if index:
            row.append("")
            column_widths.append(gap)
        row.append(card)
        column_widths.append(width)
    table = Table([row], colWidths=column_widths)
    table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return table


def _snapshot_table(
    rows: list[list[object]],
    widths: list[float],
    font_size: float = 7,
) -> Table:
    table = Table(rows, colWidths=widths)
    table.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.55, PDF_CARD_BORDER),
                ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#E3E8EF")),
                ("BACKGROUND", (0, 0), (-1, 0), PDF_BLUE),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 1), (-1, -1), PDF_TEXT),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("LEADING", (0, 0), (-1, -1), font_size + 1.5),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return table


def _snapshot_kpis(
    items: list[tuple[str, str, str]],
    backgrounds: list[colors.Color] | None = None,
) -> Table:
    gap = 5
    card_width = (744 - (gap * (len(items) - 1))) / len(items)
    card_backgrounds = backgrounds or [PDF_NEUTRAL] * len(items)
    cards: list[Table] = []
    for index, (label, value, context) in enumerate(items):
        card = Table(
            [[label], [value], [context]],
            colWidths=[card_width],
        )
        card.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 0.55, PDF_CARD_BORDER),
                    ("BACKGROUND", (0, 0), (-1, -1), card_backgrounds[index]),
                    ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 6.4),
                    ("FONTSIZE", (0, 1), (-1, 1), 14),
                    ("FONTSIZE", (0, 2), (-1, 2), 5.5),
                    ("TEXTCOLOR", (0, 0), (-1, 1), PDF_TEXT),
                    ("TEXTCOLOR", (0, 2), (-1, 2), PDF_MUTED_TEXT),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        cards.append(card)
    return _snapshot_card_row(cards, [card_width] * len(cards), gap=gap)


def _snapshot_risk_reduction_block(
    summary: dict[str, object],
    width: float = 210,
) -> Table:
    chart = _snapshot_chart(
        make_production_projects_chart(
            float(summary["total_reduction"]),
            float(summary["target_points"]),
            int(summary["latest_year"]),
        ),
        width,
        105,
        540,
        280,
    )
    on_track = (
        "Yes"
        if float(summary["total_reduction"]) >= float(summary["expected_progress"])
        else "No"
    )
    goal_summary = _snapshot_table(
        [
            [f"{summary['latest_year']} Annual Goal", "Value"],
            ["Actual YTD / Goal", f"{summary['total_reduction']:,.0f} / {summary['target_points']:,.0f}"],
            ["Completed", f"{summary['completion_pct']:.1f}%"],
            ["Calendar remaining", f"{summary['remaining_calendar_pct']:.1f}%"],
            ["On track", on_track],
        ],
        [width * 0.6, width * 0.4],
        font_size=6.2,
    )
    block = Table([[chart], [goal_summary]], colWidths=[width])
    block.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return block


def _snapshot_assessment_block(
    frame: pd.DataFrame,
    summary: dict[str, object],
    styles: dict,
    width: float = 184,
) -> Table:
    """Place the assessment pie and its color-coded counts in one PDF card."""

    chart_width = width - 62
    chart = _snapshot_chart(
        make_pdf_assessment_chart(frame),
        chart_width,
        142,
        500,
        500,
    )
    breakdown_rows = [
        ["Green", f"{summary['green_count']:,}"],
        ["Yellow", f"{summary['yellow_count']:,}"],
        ["Red", f"{summary['red_count']:,}"],
        ["Critical", f"{summary['critical_count']:,}"],
    ]
    breakdown = Table(
        breakdown_rows,
        colWidths=[36, 26],
        rowHeights=[35.5] * len(breakdown_rows),
    )
    breakdown.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.45, PDF_CARD_BORDER),
                ("INNERGRID", (0, 0), (-1, -1), 1.5, colors.white),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#69C466")),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F4B400")),
                ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#D92D20")),
                ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#6B1E5B")),
                ("TEXTCOLOR", (0, 0), (-1, 1), PDF_TEXT),
                ("TEXTCOLOR", (0, 2), (-1, 3), colors.white),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (0, -1), 5.8),
                ("FONTSIZE", (1, 0), (1, -1), 12),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    body = Table([[chart, breakdown]], colWidths=[chart_width, 62])
    body.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    title = Paragraph("<b>Ergo Assessment Breakdown</b>", styles["Small"])
    block = Table([[title], [body]], colWidths=[width])
    block.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    return block


def _snapshot_operational_items(
    summary: dict[str, object],
) -> list[tuple[str, str, str]]:
    return [
        (
            "Closed Medical Requests",
            f"{summary['closed_requests']:,}",
            "Selected period",
        ),
        (
            "Past-Due Medical Requests",
            f"{summary['past_due_requests']:,}",
            "Selected period",
        ),
        (
            f"{summary['latest_year']} Workers Comp",
            f"${summary['current_workers_comp']:,.0f}",
            "Paid out",
        ),
    ]


def _snapshot_factory_kpi_items(
    summary: dict[str, object],
) -> list[tuple[str, str, str]]:
    """Build the factory PDF KPI row with prior-year comparisons."""

    def comparison(change_key: str, available_key: str) -> str:
        if not bool(summary.get(available_key, True)):
            return "Prior-year data unavailable"
        return f"{summary[change_key]} vs prior year"

    return [
        (
            "All Ergo Rate",
            f"{summary['total_case_rate']:.2f}",
            comparison("total_case_change", "has_total_case_comparison"),
        ),
        (
            "Ergo Recordable",
            f"{summary['current_recordable']:.2f}",
            comparison("recordable_change", "has_recordable_comparison"),
        ),
        (
            "Ergo DART",
            f"{summary['current_dart']:.2f}",
            comparison("dart_change", "has_dart_comparison"),
        ),
        (
            "EHS Recordable",
            f"{summary['current_ehs_recordable']:.2f}",
            comparison(
                "ehs_recordable_change",
                "has_ehs_recordable_comparison",
            ),
        ),
        (
            "Ergo % of EHS",
            f"{summary['ergo_share']:.1f}%",
            comparison("ergo_share_change", "has_ergo_share_comparison"),
        ),
    ]


def _snapshot_factory_kpi_statuses(summary: dict[str, object]) -> list[str]:
    """Return status colors for the five factory PDF KPI cards."""

    return [
        injury_trend_status(
            summary["total_case_rate"],
            summary["prior_total_case_rate"],
            bool(summary["has_total_case_comparison"]),
        ),
        injury_trend_status(
            summary["current_recordable"],
            summary["prior_recordable"],
            bool(summary["has_recordable_comparison"]),
        ),
        injury_trend_status(
            summary["current_dart"],
            summary["prior_dart"],
            bool(summary["has_dart_comparison"]),
        ),
        injury_trend_status(
            summary["current_ehs_recordable"],
            summary["prior_ehs_recordable"],
            bool(summary["has_ehs_recordable_comparison"]),
        ),
        injury_trend_status(
            summary["ergo_share"],
            summary["prior_ergo_share"],
            bool(summary["has_ergo_share_comparison"]),
        ),
    ]


def _snapshot_case_table_rows(summary: dict[str, object]) -> list[list[str]]:
    """Build rolling-12 case counts and rates for the factory PDF table."""

    hours = float(summary["rolling_hours"])

    def row(label: str, count_key: str) -> list[str]:
        count = float(summary[count_key])
        rate = count * 200_000 / hours if hours > 0 else None
        return [
            label,
            f"{count:,.0f}",
            f"{rate:.2f}" if rate is not None else "N/A",
        ]

    return [
        ["Rolling-12 cases", "Count", "Rate"],
        row("All ergonomic cases", "ergo_all_cases_r12"),
        row("First aid", "ergo_first_aid_r12"),
        row("Recordable", "ergo_recordable_r12"),
        row("DAFW", "ergo_dafw_r12"),
        row("Restricted", "ergo_restricted_r12"),
        row("DART", "ergo_dart_r12"),
        row("All EHS recordables", "ehs_recordables_r12"),
    ]


def _snapshot_header(
    title: str,
    subtitle: str,
    right_text: str,
    styles: dict,
) -> Table:
    table = Table(
        [
            [
                Paragraph(
                    f"<b>{html.escape(title)}</b><br/><font size='8'>{html.escape(subtitle)}</font>",
                    styles["Heading2"],
                ),
                Paragraph(right_text, styles["BodyText"]),
            ]
        ],
        colWidths=[510, 234],
    )
    table.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.65, PDF_CARD_BORDER),
                ("BACKGROUND", (0, 0), (0, 0), colors.white),
                ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#EEF4F8")),
                ("LINEBEFORE", (0, 0), (0, 0), 4, PDF_BLUE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ("TEXTCOLOR", (0, 0), (-1, -1), PDF_TEXT),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return table


def _snapshot_note(series: pd.Series, fallback: str) -> str:
    notes: list[str] = []
    seen: set[str] = set()
    for value in series.fillna("").astype(str):
        cleaned = re.sub(r"\s+", " ", value).strip(" -")
        if not cleaned or cleaned.lower() == "nan" or cleaned in seen:
            continue
        seen.add(cleaned)
        notes.append(html.escape(cleaned))
    return "<br/>".join(notes[:3]) if notes else html.escape(fallback)


def _factory_snapshot_page(
    data: Dict[str, pd.DataFrame],
    scope: str,
    start_date: pd.Timestamp | None,
    end_date: pd.Timestamp | None,
    styles: dict,
    site_risk_reduction_goals: dict[str, float] | None = None,
) -> list[object]:
    production_history = aggregate_scope(
        data.get("production_injuries", pd.DataFrame()), scope
    )
    production = filter_injury_history(production_history, end_date)
    production_summary = filter_through_end_month(production_history, end_date)
    if production.empty:
        return []
    assessments = apply_scope_date_filter(
        data.get("ergo_assessments", pd.DataFrame()), scope, start_date, end_date
    )
    projects = apply_scope_ytd_filter(
        data.get("risk_projects", pd.DataFrame()), scope, end_date
    )
    medical = apply_scope_date_filter(
        data.get("medical_requests", pd.DataFrame()), scope, start_date, end_date
    )
    workers = apply_scope_date_filter(
        data.get("workers_comp", pd.DataFrame()), scope, start_date, end_date
    )
    summary = summarize_production_dashboard(
        production_summary,
        assessments,
        projects,
        medical,
        workers,
        data.get("dashboard_goals", pd.DataFrame()),
        risk_reduction_goal=(
            risk_reduction_goal_for_scope(site_risk_reduction_goals, scope)
            if site_risk_reduction_goals is not None
            else None
        ),
    )
    period_start = start_date.strftime("%b %Y") if start_date is not None else "Earliest"
    period_end = end_date.strftime("%b %Y") if end_date is not None else "Latest"
    header = _snapshot_header(
        f"{scope} Ergo Metrics",
        "AIE and Ergo Snapshot Report",
        f"<b>Selected period</b><br/>{period_start} - {period_end}<br/>"
        f"Latest completed: {summary['latest_date'].strftime('%B %Y')}",
        styles,
    )
    kpi_backgrounds = {
        "green": PDF_GREEN,
        "red": PDF_RED,
        "yellow": PDF_YELLOW,
        "neutral": PDF_NEUTRAL,
    }
    kpis = _snapshot_kpis(
        _snapshot_factory_kpi_items(summary),
        [
            kpi_backgrounds[status]
            for status in _snapshot_factory_kpi_statuses(summary)
        ],
    )
    trend = _snapshot_chart(
        make_production_trend_chart(
            production,
            recordable_goal=summary["recordable_goal"],
            dart_goal=summary["dart_goal"],
        ),
        476,
        184,
        1100,
        430,
    )
    case_table = _snapshot_table(
        _snapshot_case_table_rows(summary),
        [134, 52, 52],
    )
    top_row = _snapshot_card_row(
        [
            _snapshot_card(trend, 488),
            _snapshot_card(case_table, 250),
        ],
        [488, 250],
    )

    classifications = _snapshot_chart(
        make_ergo_classification_chart(production), 310, 166, 820, 420
    )
    risk_reduction = _snapshot_risk_reduction_block(summary, width=202)
    assessment_breakdown = _snapshot_assessment_block(
        assessments,
        summary,
        styles,
        width=184,
    )
    lower_row = _snapshot_card_row(
        [
            _snapshot_card(classifications, 322),
            _snapshot_card(risk_reduction, 214),
            _snapshot_card(assessment_breakdown, 196),
        ],
        [322, 214, 196],
    )
    operational_strip = _snapshot_kpis(
        _snapshot_operational_items(summary),
        [PDF_GREEN, PDF_RED, colors.HexColor("#E3EDF7")],
    )
    return [
        header,
        Spacer(1, 6),
        kpis,
        Spacer(1, 6),
        top_row,
        Spacer(1, 4),
        lower_row,
        Spacer(1, 4),
        operational_strip,
    ]


def _npi_snapshot_page(
    data: Dict[str, pd.DataFrame],
    program: str,
    start_date: pd.Timestamp | None,
    end_date: pd.Timestamp | None,
    styles: dict,
) -> list[object]:
    metrics = apply_factory_date_filter(
        data.get("npi_metrics", pd.DataFrame()).copy(),
        [],
        start_date,
        end_date,
        include_undated=True,
    )
    reds = apply_factory_date_filter(
        data.get("npi_ergo_reds", pd.DataFrame()).copy(),
        [],
        start_date,
        end_date,
        include_undated=True,
    )
    metrics = metrics[metrics["Program"] == program]
    reds = reds[reds["Program"] == program]
    metrics = latest_npi_snapshots(metrics)
    reds = latest_npi_snapshots(reds)
    if metrics.empty and reds.empty:
        return []

    snapshot = metrics.head(1).copy()

    has_ergo_data = not reds.empty
    has_eau_data = has_npi_values(snapshot, "EAU Target", "EAU Actual")
    has_part_data = has_npi_values(snapshot, "Part Count Target", "Part Count Actual")
    has_headcount_data = has_npi_values(snapshot, "Headcount Target", "Headcount Actual")
    has_hpu_data = has_eau_data and has_headcount_data
    has_complete_npi_data = (
        has_eau_data and has_part_data and has_headcount_data and has_ergo_data
    )
    summary = summarize_npi_dashboard(snapshot, reds)
    part_reduction_status = (
        reduction_status(summary["current_part"], summary["future_part"])
        if has_part_data
        else "neutral"
    )
    headcount_reduction_status = (
        reduction_status(summary["current_headcount"], summary["future_headcount"])
        if has_headcount_data
        else "neutral"
    )
    open_reds_status = open_ergo_reds_status(summary["open_reds"], has_ergo_data)
    pdf_status_backgrounds = {
        "green": PDF_GREEN,
        "red": PDF_RED,
        "yellow": PDF_YELLOW,
        "neutral": PDF_NEUTRAL,
    }
    status = snapshot.iloc[0] if not snapshot.empty else pd.Series(dtype=object)
    unavailable = "N/A"

    def status_value(column: str) -> str:
        value = str(status.get(column, "")).strip()
        return value if value and value.lower() not in {"nan", "none"} else unavailable

    header = _snapshot_header(
        program,
        "NPI Progress Report",
        f"<b>Snapshot month</b><br/>{latest_month_label(snapshot, reds)}<br/>"
        f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}",
        styles,
    )
    kpis = _snapshot_kpis(
        [
            ("Overall Progress", f"{summary['overall_progress']:.1f}/100" if has_complete_npi_data else unavailable, "Progress snapshot" if has_complete_npi_data else "Awaiting complete data"),
            ("Part Count", summary["part_change"] if has_part_data else unavailable, "Current to future" if has_part_data else "AIE data not available"),
            ("Headcount", summary["headcount_change"] if has_headcount_data else unavailable, "Current to future" if has_headcount_data else "AIE data not available"),
            ("Open Ergo Reds", f"{summary['open_reds']}" if has_ergo_data else unavailable, ("Open issues" if summary["open_reds"] else "No open issues") if has_ergo_data else "Ergo data not available"),
        ],
        [
            PDF_NEUTRAL,
            pdf_status_backgrounds[part_reduction_status],
            pdf_status_backgrounds[headcount_reduction_status],
            pdf_status_backgrounds[open_reds_status],
        ],
    )
    part_content: object = (
        _snapshot_chart(
            make_npi_comparison_chart(
                snapshot,
                "Part Count Target",
                "Part Count Actual",
                "Part Count Reduction",
                "Part Count",
                future_color=REDUCTION_CHART_COLORS[part_reduction_status],
            ),
            357,
            154,
            900,
            400,
        )
        if has_part_data
        else [
            Paragraph(
                "<b>Part Count Reduction</b><br/><font color='#64748B'>No AIE part-count data is available for this program.</font>",
                styles["Small"],
            ),
            Spacer(1, 125),
        ]
    )
    headcount_content: object = (
        _snapshot_chart(
            make_npi_comparison_chart(
                snapshot,
                "Headcount Target",
                "Headcount Actual",
                "Headcount Reduction",
                "Headcount",
                current_color="#B5CAE3",
                future_color=REDUCTION_CHART_COLORS[headcount_reduction_status],
            ),
            357,
            154,
            900,
            400,
        )
        if has_headcount_data
        else [
            Paragraph(
                "<b>Headcount Reduction</b><br/><font color='#64748B'>No AIE headcount data is available for this program.</font>",
                styles["Small"],
            ),
            Spacer(1, 125),
        ]
    )
    charts = _snapshot_card_row(
        [
            _snapshot_card(part_content, 369),
            _snapshot_card(headcount_content, 369),
        ],
        [369, 369],
    )
    metrics_table = _snapshot_table(
        [
            ["Metric", "Current", "Future", "% Difference"],
            ["EAU", f"{summary['current_eau']:,.0f}" if has_eau_data else unavailable, f"{summary['future_eau']:,.0f}" if has_eau_data else unavailable, format_change(summary["current_eau"], summary["future_eau"]) if has_eau_data else unavailable],
            ["Part Count", f"{summary['current_part']:,.0f}" if has_part_data else unavailable, f"{summary['future_part']:,.0f}" if has_part_data else unavailable, summary["part_change"] if has_part_data else unavailable],
            ["Headcount", f"{summary['current_headcount']:,.0f}" if has_headcount_data else unavailable, f"{summary['future_headcount']:,.0f}" if has_headcount_data else unavailable, summary["headcount_change"] if has_headcount_data else unavailable],
            ["Hours Per Unit", f"{summary['current_hpu']:.2f}" if has_hpu_data else unavailable, f"{summary['future_hpu']:.2f}" if has_hpu_data else unavailable, summary["hpu_change"] if has_hpu_data else unavailable],
        ],
        [96, 72, 72, 84],
    )
    metrics_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (3, 2), (3, 2), pdf_status_backgrounds[part_reduction_status]),
                ("BACKGROUND", (3, 3), (3, 3), pdf_status_backgrounds[headcount_reduction_status]),
            ]
        )
    )
    process_table = _snapshot_table(
        [
            ["Process planning tool", "Status"],
            ["Fishbone", status_value("Fishbone Status")],
            ["At-a-Glance / PAAG", status_value("PAAG Status")],
            ["Pin Map", status_value("Pin Map Status")],
        ],
        [180, 144],
    )
    left_detail = Table([[metrics_table], [Spacer(1, 4)], [process_table]], colWidths=[324])
    left_detail.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    reds_content: object = (
        _snapshot_chart(
            make_npi_detail_snapshot_chart(reds), 390, 150, 950, 365
        )
        if has_ergo_data
        else [
            Paragraph(
                "<b>Ergo Reds Snapshot</b><br/><font color='#64748B'>No ergo assessment data is available for this program.</font>",
                styles["Small"],
            ),
            Spacer(1, 115),
        ]
    )
    details = _snapshot_card_row(
        [
            _snapshot_card(left_detail, 336),
            _snapshot_card(reds_content, 402),
        ],
        [336, 402],
    )
    aie_notes = Paragraph(
        "<b>AIE Wins and Opportunities</b><br/>"
        + _snapshot_note(
            snapshot.get("AIE Wins Opportunities", pd.Series(dtype=str)),
            "No AIE notes provided.",
        ),
        styles["Small"],
    )
    ergo_notes = Paragraph(
        "<b>Ergo Wins and Opportunities</b><br/>"
        + _snapshot_note(
            reds.get("Ergo Wins Opportunities", pd.Series(dtype=str)),
            "No ergo notes provided.",
        ),
        styles["Small"],
    )
    notes = _snapshot_card_row(
        [
            _snapshot_card(aie_notes, 369, background=colors.HexColor("#FBFCFE"), padding=7),
            _snapshot_card(ergo_notes, 369, background=colors.HexColor("#FBFCFE"), padding=7),
        ],
        [369, 369],
    )
    return [header, Spacer(1, 6), kpis, Spacer(1, 5), charts, Spacer(1, 4), details, Spacer(1, 4), notes]


def render_combined_pdf_report(
    data: Dict[str, pd.DataFrame],
    scopes: list[str],
    programs: list[str],
    start_date: pd.Timestamp | None = None,
    end_date: pd.Timestamp | None = None,
    site_risk_reduction_goals: dict[str, float] | None = None,
) -> bytes:
    if not scopes and not programs:
        raise ValueError("Select at least one factory/scope or NPI program.")
    ensure_plotly_image_export_ready()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=24,
        leftMargin=24,
        topMargin=20,
        bottomMargin=18,
        title="AIE and Ergo Combined Snapshot Report",
        author="GE Appliances AIE and Ergonomics Dashboard",
    )
    styles = getSampleStyleSheet()
    styles["Heading2"].fontSize = 16
    styles["Heading2"].leading = 17
    styles["Heading2"].textColor = PDF_TEXT
    styles["BodyText"].fontSize = 7
    styles["BodyText"].leading = 9
    styles["BodyText"].textColor = PDF_MUTED_TEXT
    styles.add(
        type(styles["BodyText"])(
            "Small",
            parent=styles["BodyText"],
            fontSize=6.2,
            leading=7.4,
            textColor=colors.HexColor("#475569"),
        )
    )
    pages: list[list[object]] = []
    for scope in scopes:
        page = _factory_snapshot_page(
            data,
            scope,
            start_date,
            end_date,
            styles,
            site_risk_reduction_goals,
        )
        if page:
            pages.append(page)
    for program in programs:
        page = _npi_snapshot_page(data, program, start_date, end_date, styles)
        if page:
            pages.append(page)
    if not pages:
        raise ValueError("No report data is available for the selected pages and dates.")
    story: list[object] = []
    for index, page in enumerate(pages):
        if index:
            story.append(PageBreak())
        story.append(KeepTogether(page))

    def footer(canvas, document) -> None:
        canvas.saveState()
        canvas.setFillColor(PDF_PAGE_BACKGROUND)
        canvas.rect(0, 0, landscape(letter)[0], landscape(letter)[1], fill=1, stroke=0)
        canvas.setStrokeColor(PDF_CARD_BORDER)
        canvas.setLineWidth(0.5)
        canvas.line(24, 15, 768, 15)
        canvas.setFont("Helvetica", 5.5)
        canvas.setFillColor(PDF_MUTED_TEXT)
        canvas.drawString(
            24,
            6,
            "Rates: rolling-12 cases x 200,000 / hours | DART: DAFW or restricted | All Ergo includes first aid",
        )
        canvas.drawRightString(768, 6, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer.getvalue()


def main() -> None:
    st.title("Ergonomics and Industrial Engineering Dashboard")
    st.caption("Upload the attached Excel template or use it automatically if it is present in the workspace.")

    template_paths = [
        Path(__file__).with_name("AIE_Ergo_Dashboard_Template NICOLES VERSION.xlsx"),
        Path.cwd() / "AIE_Ergo_Dashboard_Template NICOLES VERSION.xlsx",
    ]
    template_path = next((path for path in template_paths if path.exists()), None)
    if template_path is not None:
        st.sidebar.download_button(
            "Download Excel template",
            data=template_path.read_bytes(),
            file_name=template_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    uploaded_file = st.sidebar.file_uploader("Upload Excel workbook", type=["xlsx", "xlsm"])
    page = st.sidebar.radio("Page", ["Production Dashboard", "NPI Dashboard"])
    st.sidebar.markdown("---")

    file_bytes = uploaded_file.getvalue() if uploaded_file is not None else None
    data, error_message = load_workbook_data(file_bytes, uploaded_file.name if uploaded_file is not None else None)

    if error_message:
        st.error(error_message)
        return

    if not data:
        st.info("No usable data is available yet. Upload a workbook with the expected sheets.")
        return

    prod_df = data.get("production_injuries", pd.DataFrame()).copy()
    assessment_df = data.get("ergo_assessments", pd.DataFrame()).copy()
    risk_project_df = data.get("risk_projects", pd.DataFrame()).copy()
    risk_reduction_goals_df = data.get("risk_reduction_goals", pd.DataFrame()).copy()
    medical_df = data.get("medical_requests", pd.DataFrame()).copy()
    workers_comp_df = data.get("workers_comp", pd.DataFrame()).copy()
    dashboard_goals_df = data.get("dashboard_goals", pd.DataFrame()).copy()
    npi_metrics_df = data.get("npi_metrics", pd.DataFrame()).copy()
    npi_reds_df = data.get("npi_ergo_reds", pd.DataFrame()).copy()

    factories = sorted({str(v) for v in pd.concat([assessment_df["Factory"], risk_project_df["Factory"], medical_df["Factory"], workers_comp_df["Factory"], npi_metrics_df["Factory"]], ignore_index=True) if str(v) != "nan" and str(v) != "None"})
    scopes = available_scopes(prod_df)
    programs = available_npi_programs(npi_metrics_df, npi_reds_df)
    site_risk_reduction_goals: dict[str, float] = {}

    with st.sidebar:
        if page == "Production Dashboard":
            selected_scope = st.selectbox("Reporting scope", scopes, index=0)
            selected_factories = []
            selected_programs = []
            scoped_for_dates = aggregate_scope(prod_df, selected_scope)
        else:
            selected_scope = "GE Appliances"
            selected_factories = st.multiselect("Factory", factories, default=[])
            selected_programs = st.multiselect("Program", programs, default=[])
            scoped_for_dates = aggregate_scope(prod_df, selected_scope)
        valid_months = (
            scoped_for_dates.dropna(subset=["Ergo All Rate R12"])
            if not scoped_for_dates.empty
            else pd.DataFrame()
        )
        min_date = pd.Timestamp(valid_months["Date"].min()) if not valid_months.empty else None
        max_date = pd.Timestamp(valid_months["Date"].max()) if not valid_months.empty else None
        if min_date is not None and max_date is not None:
            default_start, default_end = previous_calendar_month()
            date_range = st.date_input(
                "Date range",
                value=(
                    default_start.to_pydatetime().date(),
                    default_end.to_pydatetime().date(),
                ),
                key="date_range",
                help=(
                    "Defaults to the previous calendar month. Injury charts use "
                    "the trailing 12 months through the selected end month."
                ),
            )
            start_date, end_date = normalize_date_selection(
                date_range,
                fallback_start=default_start,
                fallback_end=default_end,
            )
        else:
            st.info("Date filter is unavailable for the current workbook.")
            start_date = None
            end_date = None
        goal_year = (
            end_date.year
            if end_date is not None
            else previous_calendar_month()[1].year
        )
        goal_sites = [
            scope
            for scope in scopes
            if scope not in {"GE Appliances", "Supply Chain", "Other"}
        ]
        site_risk_reduction_goals = default_site_risk_reduction_goals(
            risk_reduction_goals_df,
            goal_sites,
            goal_year,
        )
        st.markdown("---")
        st.markdown(f"**{goal_year} Annual Risk Reduction Goals**")
        goal_rows = pd.DataFrame(
            [
                {
                    "Site": site,
                    "Annual Goal": site_risk_reduction_goals.get(site, 0.0),
                }
                for site in goal_sites
            ],
            columns=["Site", "Annual Goal"],
        )
        st.dataframe(
            goal_rows,
            hide_index=True,
            width="stretch",
            height=min(38 + 35 * max(len(goal_rows), 1), 330),
            column_config={
                "Site": st.column_config.TextColumn("Site"),
                "Annual Goal": st.column_config.NumberColumn(
                    "Annual Goal",
                    min_value=0.0,
                    step=1.0,
                    format="%.0f",
                ),
            },
        )
        goals_for_selected_year = (
            risk_reduction_goals_df[
                risk_reduction_goals_df["Year"] == goal_year
            ]
            if not risk_reduction_goals_df.empty
            else pd.DataFrame()
        )
        if risk_reduction_goals_df.empty:
            st.warning(
                "No Ergo_Risk_Reduction sheet was found. Add Site, Year, and "
                "Risk_Reduction_Goals columns to the uploaded workbook."
            )
        elif goals_for_selected_year.empty:
            st.warning(
                f"Ergo_Risk_Reduction has no {goal_year} rows. Goals will be 0 "
                "until that year is added to the workbook."
            )
        else:
            st.caption("Loaded from the Ergo_Risk_Reduction workbook tab.")
        selected_scope_goal = risk_reduction_goal_for_scope(
            site_risk_reduction_goals,
            selected_scope,
        )
        st.caption(
            f"{selected_scope} annual goal: {selected_scope_goal:,.0f} points. "
            "Actual points are tracked year to date through the selected end month."
        )
        st.markdown("---")
        st.markdown("**Combined PDF report**")
        default_pdf_sites = default_combined_pdf_sites(scopes)
        pdf_scopes = st.multiselect(
            "Factory/scope pages",
            scopes,
            default=default_pdf_sites,
            help=(
                "All sites and aggregate scopes except Other are selected by default. Each "
                "selected factory or aggregate scope becomes one landscape page."
            ),
            key="combined_pdf_scopes_all_except_other",
        )
        pdf_programs = st.multiselect(
            "NPI program pages",
            programs,
            default=programs,
            help="Each selected NPI program becomes one landscape page in the same PDF.",
            key="combined_pdf_programs",
        )
        st.caption(
            f"Planned report: {len(pdf_scopes)} factory page(s) + "
            f"{len(pdf_programs)} NPI page(s)."
        )

    if page == "Production Dashboard":
        production_history = aggregate_scope(prod_df, selected_scope)
        prod_filtered = filter_injury_history(production_history, end_date)
        production_summary = filter_through_end_month(production_history, end_date)
        assessment_filtered = apply_scope_date_filter(assessment_df, selected_scope, start_date, end_date)
        risk_project_filtered = apply_scope_ytd_filter(
            risk_project_df,
            selected_scope,
            end_date,
        )
        medical_filtered = apply_scope_date_filter(medical_df, selected_scope, start_date, end_date)
        workers_comp_filtered = apply_scope_date_filter(workers_comp_df, selected_scope, start_date, end_date)

        st.subheader(f"Production Dashboard — {selected_scope}")
        if prod_filtered.empty:
            st.info("No production data is available for the selected filters.")
        else:
            render_production_style()
            summary = summarize_production_dashboard(
                production_summary,
                assessment_filtered,
                risk_project_filtered,
                medical_filtered,
                workers_comp_filtered,
                dashboard_goals_df,
                risk_reduction_goal=selected_scope_goal,
            )
            latest_year = summary["latest_year"]
            st.caption(
                f"Latest completed month: {summary['latest_date'].strftime('%B %Y')} · "
                "Rates = rolling-12 cases × 200,000 ÷ rolling-12 hours"
            )
            metric_columns = st.columns([1, 1, 1.3, 1, 1])
            metric_columns[0].metric(
                "All Ergo Rate (R12)",
                f"{summary['total_case_rate']:.2f}",
                delta=f"{summary['total_case_change']} vs prior year",
                delta_color="inverse",
            )
            metric_columns[1].metric(
                "Ergo Recordable Rate (R12)",
                f"{summary['current_recordable']:.2f}",
                delta=f"{summary['recordable_change']} vs prior year",
                delta_color="inverse",
            )
            metric_columns[2].metric(
                "Ergo DART Rate (R12)",
                f"{summary['current_dart']:.2f}",
                delta=f"{summary['dart_change']} vs prior year",
                delta_color="inverse",
            )
            metric_columns[3].metric(
                "Total EHS Recordable Rate",
                f"{summary['current_ehs_recordable']:.2f}",
                delta=f"{summary['ehs_recordable_change']} vs prior year",
                delta_color="inverse",
            )
            metric_columns[4].metric(
                "Ergo % of EHS Recordables",
                f"{summary['ergo_share']:.1f}%",
                delta=f"{summary['ergo_share_change']} vs prior year",
                delta_color="inverse",
            )
            medical_table = pd.DataFrame(
                [
                    {"Metric": "Closed Medical Requests (YTD)", "Count": summary["closed_requests"], "% of total": f"{(summary['closed_requests'] / summary['total_medical'] * 100):.0f}%" if summary["total_medical"] else "0%"},
                    {"Metric": "Past Due Medical Requests", "Count": summary["past_due_requests"], "% of total": f"{(summary['past_due_requests'] / summary['total_medical'] * 100):.0f}%" if summary["total_medical"] else "0%"},
                    {"Metric": "Open Medical Requests", "Count": summary["open_requests"], "% of total": f"{(summary['open_requests'] / summary['total_medical'] * 100):.0f}%" if summary["total_medical"] else "0%"},
                ]
            )
            injury_table = pd.DataFrame(
                [
                    {"Metric": "All Ergo Rate (R12)", "Current": f"{summary['total_case_rate']:.2f}", "Prior year": f"{summary['prior_total_case_rate']:.2f}", "Change": summary["total_case_change"]},
                    {"Metric": "Ergo Recordable Rate (R12)", "Current": f"{summary['current_recordable']:.2f}", "Prior year": f"{summary['prior_recordable']:.2f}", "Change": summary["recordable_change"]},
                    {"Metric": "Ergo DART Rate (R12)", "Current": f"{summary['current_dart']:.2f}", "Prior year": f"{summary['prior_dart']:.2f}", "Change": summary["dart_change"]},
                    {"Metric": "Ergo Recordables / EHS Recordables", "Current": f"{summary['ergo_share']:.1f}%", "Prior year": "—", "Change": "—"},
                ]
            )
            workers_table = summary["workers_by_year"].copy()
            if not workers_table.empty:
                workers_table["Paid Out"] = workers_table["Amount"].map(lambda value: f"$ {value:,.0f}")
                workers_table = workers_table[["Year", "Paid Out"]]

            top_left, top_right = st.columns([2.4, 1.0])
            with top_left:
                with st.container(border=True):
                    st.plotly_chart(
                        make_production_trend_chart(
                            prod_filtered,
                            recordable_goal=summary["recordable_goal"],
                            dart_goal=summary["dart_goal"],
                        ),
                        width="stretch",
                    )
            with top_right:
                with st.container(border=True):
                    assess_top = st.columns([2.0, 1.0])
                    with assess_top[0]:
                        st.plotly_chart(make_production_assessment_chart(assessment_filtered), width="stretch")
                    with assess_top[1]:
                        st.markdown(f'<div class="prod-metric-box prod-metric-green"><div class="prod-metric-value">{summary["green_count"]}</div><div class="prod-kicker">Green</div></div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="prod-metric-box prod-metric-yellow"><div class="prod-metric-value">{summary["yellow_count"]}</div><div class="prod-kicker">Yellow</div></div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="prod-metric-box prod-metric-red"><div class="prod-metric-value">{summary["red_count"]}</div><div class="prod-kicker">Red</div></div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="prod-metric-box prod-metric-critical"><div class="prod-metric-value">{summary["critical_count"]}</div><div class="prod-kicker">Critical</div></div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="prod-metric-box prod-metric-neutral"><div class="prod-metric-value">{summary["total_assessments"]}</div><div class="prod-kicker">Total</div></div>', unsafe_allow_html=True)

            cases_column, risk_reduction_column = st.columns([2.4, 1.0])
            with cases_column:
                with st.container(border=True):
                    st.plotly_chart(
                        make_ergo_classification_chart(prod_filtered),
                        width="stretch",
                    )
            with risk_reduction_column:
                with st.container(border=True):
                    on_track = "Yes" if summary["total_reduction"] >= summary["expected_progress"] else "No"
                    st.plotly_chart(
                        make_production_projects_chart(
                            summary["total_reduction"],
                            summary["target_points"],
                            latest_year,
                        ),
                        width="stretch",
                    )
                    st.markdown(
                        f"""
                        <div class="goal-summary-card">
                            <div class="goal-summary-title">{latest_year} Annual Risk Reduction Goal</div>
                            <div class="goal-summary-grid">
                                <div class="goal-summary-item"><div class="goal-summary-label">Actual YTD</div><div class="goal-summary-value">{summary['total_reduction']:,.0f}</div></div>
                                <div class="goal-summary-item"><div class="goal-summary-label">Annual goal</div><div class="goal-summary-value">{summary['target_points']:,.0f}</div></div>
                                <div class="goal-summary-item"><div class="goal-summary-label">Completed</div><div class="goal-summary-value">{summary['completion_pct']:.1f}%</div></div>
                                <div class="goal-summary-item"><div class="goal-summary-label">Calendar remaining</div><div class="goal-summary-value">{summary['remaining_calendar_pct']:.1f}%</div></div>
                                <div class="goal-summary-item"><div class="goal-summary-label">On track?</div><div class="goal-summary-value">{on_track}</div></div>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            medical_column, injury_column, workers_column = st.columns([1.0, 1.35, 1.0])
            with medical_column:
                with st.container(border=True):
                    st.markdown("**Medical Requests**")
                    st.dataframe(medical_table, width="stretch", hide_index=True)
            with injury_column:
                with st.container(border=True):
                    st.markdown("**Injury Trend Analysis**")
                    st.dataframe(injury_table, width="stretch", hide_index=True)
            with workers_column:
                with st.container(border=True):
                    st.markdown("**Workers Compensation Ergo Claims Paid Out (YTD)**")
                    if not workers_table.empty:
                        st.dataframe(workers_table, width="stretch", hide_index=True)
                        if summary["prior_workers_comp"]:
                            st.caption(f"Change vs prior year: {summary['workers_change']:+.1f}%")
                    else:
                        st.caption("No workers compensation data is available for the selected filters.")
            st.caption(
                "Data note: the EHS source provides a field named EHS_DAFW_Cases but "
                "labels its supplied rate as EHS DART. Until an all-EHS restricted-case "
                "count is available, the dashboard carries that source field as the EHS "
                "DART input. Other-location ergonomic cases are included in GE Appliances "
                "and Supply Chain counts; no separate hours are available for the Other site."
            )
    else:
        npi_metrics_filtered = apply_factory_date_filter(
            npi_metrics_df,
            selected_factories,
            start_date,
            end_date,
            include_undated=True,
        )
        npi_metrics_filtered = apply_program_filter(npi_metrics_filtered, selected_programs)
        npi_reds_filtered = apply_factory_date_filter(
            npi_reds_df,
            selected_factories,
            start_date,
            end_date,
            include_undated=True,
        )
        npi_reds_filtered = apply_program_filter(npi_reds_filtered, selected_programs)
        npi_metrics_filtered = latest_npi_snapshots(npi_metrics_filtered)
        npi_reds_filtered = latest_npi_snapshots(npi_reds_filtered)

        st.subheader("NPI Dashboard")
        if npi_metrics_filtered.empty and npi_reds_filtered.empty:
            st.info("No NPI data is available for the selected filters.")
        else:
            render_npi_style()
            has_aie_data = not npi_metrics_filtered.empty
            has_ergo_data = not npi_reds_filtered.empty
            has_eau_data = has_npi_values(npi_metrics_filtered, "EAU Target", "EAU Actual")
            has_part_data = has_npi_values(npi_metrics_filtered, "Part Count Target", "Part Count Actual")
            has_headcount_data = has_npi_values(npi_metrics_filtered, "Headcount Target", "Headcount Actual")
            has_hpu_data = has_eau_data and has_headcount_data
            has_complete_npi_data = has_eau_data and has_part_data and has_headcount_data and has_ergo_data
            summary = summarize_npi_dashboard(npi_metrics_filtered, npi_reds_filtered)
            part_reduction_status = (
                reduction_status(summary["current_part"], summary["future_part"])
                if has_part_data
                else "neutral"
            )
            headcount_reduction_status = (
                reduction_status(summary["current_headcount"], summary["future_headcount"])
                if has_headcount_data
                else "neutral"
            )
            open_reds_status = open_ergo_reds_status(
                summary["open_reds"],
                has_ergo_data,
            )
            program_names = available_npi_programs(npi_metrics_filtered, npi_reds_filtered)
            primary_title = program_names[0] if len(program_names) == 1 else f"{len(program_names)} NPI Programs"
            status_row = (
                npi_metrics_filtered.iloc[0]
                if has_aie_data
                else pd.Series(
                    {
                        "Fishbone Status": "N/A",
                        "PAAG Status": "N/A",
                        "Pin Map Status": "N/A",
                    }
                )
            )
            unavailable = "N/A"
            current_parts_per_eau = (
                summary["current_part"] / summary["current_eau"]
                if has_part_data and has_eau_data and summary["current_eau"]
                else None
            )
            future_parts_per_eau = (
                summary["future_part"] / summary["future_eau"]
                if has_part_data and has_eau_data and summary["future_eau"]
                else None
            )
            part_metrics_table = pd.DataFrame(
                [
                    {"Metric": "EAU (Annual Demand)", "Current": f"{summary['current_eau']:,.0f}" if has_eau_data else unavailable, "Future": f"{summary['future_eau']:,.0f}" if has_eau_data else unavailable, "% Difference": format_change(summary["current_eau"], summary["future_eau"]) if has_eau_data else unavailable},
                    {"Metric": "Part Count", "Current": f"{summary['current_part']:,.0f}" if has_part_data else unavailable, "Future": f"{summary['future_part']:,.0f}" if has_part_data else unavailable, "% Difference": summary["part_change"] if has_part_data else unavailable},
                    {"Metric": "Parts per EAU", "Current": f"{current_parts_per_eau:.4f}" if current_parts_per_eau is not None else unavailable, "Future": f"{future_parts_per_eau:.4f}" if future_parts_per_eau is not None else unavailable, "% Difference": format_change(current_parts_per_eau, future_parts_per_eau) if current_parts_per_eau is not None and future_parts_per_eau is not None else unavailable},
                ]
            )
            headcount_metrics_table = pd.DataFrame(
                [
                    {"Metric": "EAU (Annual Demand)", "Current": f"{summary['current_eau']:,.0f}" if has_eau_data else unavailable, "Future": f"{summary['future_eau']:,.0f}" if has_eau_data else unavailable, "% Difference": format_change(summary["current_eau"], summary["future_eau"]) if has_eau_data else unavailable},
                    {"Metric": "Headcount", "Current": f"{summary['current_headcount']:,.0f}" if has_headcount_data else unavailable, "Future": f"{summary['future_headcount']:,.0f}" if has_headcount_data else unavailable, "% Difference": summary["headcount_change"] if has_headcount_data else unavailable},
                    {"Metric": "HPU (Hours Per Unit)", "Current": f"{summary['current_hpu']:.2f}" if has_hpu_data else unavailable, "Future": f"{summary['future_hpu']:.2f}" if has_hpu_data else unavailable, "% Difference": summary["hpu_change"] if has_hpu_data else unavailable},
                ]
            )

            st.markdown(
                f"""
                <div class="npi-header">
                    <div class="npi-title">{primary_title}</div>
                    <div class="npi-subtitle">NPI Progress Report</div>
                    <div class="npi-meta">Month: {latest_month_label(npi_metrics_filtered, npi_reds_filtered)} | Generated: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            card_cols = st.columns(4)
            card_cols[0].markdown(make_summary_card("Overall Progress", f"{summary['overall_progress']:.1f}/100" if has_complete_npi_data else unavailable, "Progress snapshot" if has_complete_npi_data else "Awaiting complete AIE and ergo data", "progress"), unsafe_allow_html=True)
            card_cols[1].markdown(make_summary_card("Part Count Reduction", summary["part_change"] if has_part_data else unavailable, "Change vs current count" if has_part_data else "AIE data not available", part_reduction_status), unsafe_allow_html=True)
            card_cols[2].markdown(make_summary_card("Headcount Reduction", summary["headcount_change"] if has_headcount_data else unavailable, "Change vs current headcount" if has_headcount_data else "AIE data not available", headcount_reduction_status), unsafe_allow_html=True)
            card_cols[3].markdown(make_summary_card("Open Ergo Reds", f"{summary['open_reds']}" if has_ergo_data else unavailable, ("No open issues" if summary["open_reds"] == 0 else "Open issues requiring action") if has_ergo_data else "Ergo data not available", open_reds_status), unsafe_allow_html=True)

            upper_left, upper_right = st.columns(2)
            with upper_left:
                with st.container(border=True):
                    st.dataframe(part_metrics_table, width="stretch", hide_index=True)
                    if has_part_data:
                        st.plotly_chart(
                            make_npi_comparison_chart(
                                npi_metrics_filtered,
                                "Part Count Target",
                                "Part Count Actual",
                                "Part Count Reduction",
                                "Part Count",
                                future_color=REDUCTION_CHART_COLORS[part_reduction_status],
                            ),
                            width="stretch",
                        )
                    else:
                        st.caption("No AIE part-count data is available for this program.")
            with upper_right:
                with st.container(border=True):
                    st.dataframe(headcount_metrics_table, width="stretch", hide_index=True)
                    if has_headcount_data:
                        st.plotly_chart(
                            make_npi_comparison_chart(
                                npi_metrics_filtered,
                                "Headcount Target",
                                "Headcount Actual",
                                "Headcount Reduction",
                                "Headcount",
                                current_color="#B5CAE3",
                                future_color=REDUCTION_CHART_COLORS[headcount_reduction_status],
                            ),
                            width="stretch",
                        )
                    else:
                        st.caption("No AIE headcount data is available for this program.")

            lower_left, lower_right = st.columns(2)
            with lower_left:
                with st.container(border=True):
                    st.markdown("**Process Planning Tools**")
                    status_cols = st.columns(3)
                    with status_cols[0]:
                        render_status_orb("Fishbone", str(status_row["Fishbone Status"]))
                    with status_cols[1]:
                        render_status_orb("At-a-Glance", str(status_row["PAAG Status"]))
                    with status_cols[2]:
                        render_status_orb("Pin Map", str(status_row["Pin Map Status"]))
            with lower_right:
                with st.container(border=True):
                    if has_ergo_data:
                        st.plotly_chart(make_npi_detail_snapshot_chart(npi_reds_filtered), width="stretch")
                    else:
                        st.caption("No ergo assessment data is available for this program.")
                    reds_summary = (
                        npi_reds_filtered.groupby("Status", as_index=False)["Count"].sum()
                        .sort_values("Status")
                    )
                    snapshot_text = ", ".join(f"{row.Status}: {int(row.Count)}" for row in reds_summary.itertuples()) if not reds_summary.empty else "No ergo red details available."
                    st.caption(snapshot_text)

            wins_cols = st.columns(2)
            with wins_cols[0]:
                st.markdown(
                    f"""
                    <div class="wins-card">
                        <div class="wins-title">AIE Wins and Opportunities</div>
                        <div class="wins-body">{collect_note_text(npi_metrics_filtered["AIE Wins Opportunities"], "No AIE wins or opportunities were provided in the workbook.")}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with wins_cols[1]:
                st.markdown(
                    f"""
                    <div class="wins-card">
                        <div class="wins-title">Ergo Wins and Opportunities</div>
                        <div class="wins-body">{collect_note_text(npi_reds_filtered["Ergo Wins Opportunities"], "No ergo wins or opportunities were provided in the workbook.")}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.sidebar.markdown("---")
    if st.sidebar.button("Build Combined PDF"):
        try:
            with st.sidebar:
                with st.spinner(
                    "Building combined PDF. This may take a moment...",
                    show_time=True,
                    width="stretch",
                ):
                    pdf_bytes = render_combined_pdf_report(
                        data,
                        scopes=pdf_scopes,
                        programs=pdf_programs,
                        start_date=start_date,
                        end_date=end_date,
                        site_risk_reduction_goals=site_risk_reduction_goals,
                    )
        except (RuntimeError, ValueError) as exc:
            st.sidebar.error(str(exc))
        else:
            report_end = (
                end_date.strftime("%Y-%m") if end_date is not None else "latest"
            )
            st.sidebar.success("Combined PDF is ready to download.")
            st.sidebar.download_button(
                "Download combined snapshot PDF",
                data=pdf_bytes,
                file_name=f"aie_ergo_combined_snapshot_{report_end}.pdf",
                mime="application/pdf",
            )


if __name__ == "__main__":
    main()
