import io
import importlib
import re
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Image, PageBreak

st.set_page_config(page_title="Ergonomics & Industrial Engineering Dashboard", layout="wide")

SHEET_ALIASES = {
    "production_injuries": ["production_injuries", "Ergo_Injuries"],
    "ergo_assessments": ["ergo_assessments", "Ergo_Assessments_Risk_Reduction"],
    "risk_projects": ["risk_projects", "Ergo_Assessments_Risk_Reduction"],
    "medical_requests": ["medical_requests", "Medical_Requests"],
    "npi_metrics": ["npi_metrics", "AIE_NPI"],
    "npi_ergo_reds": ["npi_ergo_reds", "Ergo_NPI"],
}

REQUIRED_COLUMNS = {
    "production_injuries": ["Date", "Factory", "Injuries", "Recordable Rate", "DART Rate"],
    "ergo_assessments": ["Factory", "Risk Level", "Count"],
    "risk_projects": ["Factory", "Date", "Reduction Points", "Target Points"],
    "medical_requests": ["Factory", "Status", "Count"],
    "npi_metrics": ["Program", "Factory", "Headcount Actual", "Headcount Target"],
    "npi_ergo_reds": ["Program", "Status", "Count"],
}


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

    for dataset_name, aliases in SHEET_ALIASES.items():
        matched_sheet = None
        for alias in aliases:
            if alias in excel_file.sheet_names:
                matched_sheet = alias
                break
        if matched_sheet is None:
            errors.append(f"Missing required sheet: {dataset_name}")
            continue

        frame = pd.read_excel(io.BytesIO(bytes_data), sheet_name=matched_sheet, engine="openpyxl")

        if dataset_name == "production_injuries":
            normalized[dataset_name] = build_production_df(frame)
        elif dataset_name == "ergo_assessments":
            normalized[dataset_name] = build_assessment_df(frame)
        elif dataset_name == "risk_projects":
            normalized[dataset_name] = build_risk_project_df(frame)
        elif dataset_name == "medical_requests":
            normalized[dataset_name] = build_medical_requests_df(frame)
        elif dataset_name == "npi_metrics":
            normalized[dataset_name] = build_npi_metrics_df(frame)
        elif dataset_name == "npi_ergo_reds":
            normalized[dataset_name] = build_npi_ergo_reds_df(frame)

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
    df["Date"] = pd.to_datetime(df["Year"].astype(str) + "-" + df["Month"].astype(str).str.zfill(2) + "-01")
    df["Injuries"] = pd.to_numeric(df["Injuries"], errors="coerce").fillna(0).astype(int)
    df["Recordable Rate"] = pd.to_numeric(df["Recordable Rate"], errors="coerce")
    df["DART Rate"] = pd.to_numeric(df["DART Rate"], errors="coerce")
    return df[["Date", "Factory", "Injuries", "Recordable Rate", "DART Rate"]]


def build_assessment_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    critical_red_col = find_column(frame, ["Critical_Red", "Critical Red"])
    red_col = find_column(frame, ["Red"])
    yellow_col = find_column(frame, ["Yellow"])
    green_col = find_column(frame, ["Green"])
    if not all([factory_col, red_col, yellow_col, green_col]):
        raise ValueError("Ergo assessments sheet is missing one or more required columns.")
    df = frame[[factory_col, critical_red_col or red_col, red_col, yellow_col, green_col]].copy()
    df.columns = ["Factory", "Critical_Red", "Red", "Yellow", "Green"]
    df["Factory"] = df["Factory"].astype(str)
    df["Critical_Red"] = pd.to_numeric(df["Critical_Red"], errors="coerce").fillna(0)
    df["Red"] = pd.to_numeric(df["Red"], errors="coerce").fillna(0)
    df["Yellow"] = pd.to_numeric(df["Yellow"], errors="coerce").fillna(0)
    df["Green"] = pd.to_numeric(df["Green"], errors="coerce").fillna(0)
    red_total = df["Red"] + df["Critical_Red"]
    rows = []
    for _, row in df.iterrows():
        rows.append({"Factory": row["Factory"], "Risk Level": "Red", "Count": int(red_total.loc[_])})
        rows.append({"Factory": row["Factory"], "Risk Level": "Yellow", "Count": int(row["Yellow"])})
        rows.append({"Factory": row["Factory"], "Risk Level": "Green", "Count": int(row["Green"])})
    return pd.DataFrame(rows)


def build_risk_project_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    date_col = find_column(frame, ["Date", "Month", "Year"])
    reduction_col = find_column(frame, ["Risk_Reduction_Points_This_Month", "Reduction Points"])
    target_col = find_column(frame, ["Target Points", "Target_Points", "Target"])
    if not all([factory_col, reduction_col]):
        raise ValueError("Risk projects sheet is missing required columns.")
    df = frame[[factory_col, reduction_col]].copy()
    df.columns = ["Factory", "Reduction Points"]
    df["Factory"] = df["Factory"].astype(str)
    df["Reduction Points"] = pd.to_numeric(df["Reduction Points"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(frame[date_col] if date_col else pd.Series([pd.Timestamp.today()] * len(frame)))
    if target_col:
        df["Target Points"] = pd.to_numeric(frame[target_col], errors="coerce").fillna(0)
    else:
        monthly_average = df["Reduction Points"].mean() if not df.empty else 0
        df["Target Points"] = monthly_average * 12
    return df[["Factory", "Date", "Reduction Points", "Target Points"]]


def build_medical_requests_df(frame: pd.DataFrame) -> pd.DataFrame:
    factory_col = find_column(frame, ["Site", "Factory"])
    completed_col = find_column(frame, ["Completed", "Closed"])
    past_due_col = find_column(frame, ["Completed_Past_Due", "Past Due"])
    open_col = find_column(frame, ["Open"])
    if not all([factory_col, completed_col, past_due_col, open_col]):
        raise ValueError("Medical requests sheet is missing required columns.")
    df = frame[[factory_col, completed_col, past_due_col, open_col]].copy()
    df.columns = ["Factory", "Closed", "Past Due", "Open"]
    df["Factory"] = df["Factory"].astype(str)
    df["Closed"] = pd.to_numeric(df["Closed"], errors="coerce").fillna(0).astype(int)
    df["Past Due"] = pd.to_numeric(df["Past Due"], errors="coerce").fillna(0).astype(int)
    df["Open"] = pd.to_numeric(df["Open"], errors="coerce").fillna(0).astype(int)
    rows = []
    for _, row in df.iterrows():
        rows.append({"Factory": row["Factory"], "Status": "Closed", "Count": int(row["Closed"])})
        rows.append({"Factory": row["Factory"], "Status": "Past Due", "Count": int(row["Past Due"])})
        rows.append({"Factory": row["Factory"], "Status": "Open", "Count": int(row["Open"])})
    return pd.DataFrame(rows)


def build_npi_metrics_df(frame: pd.DataFrame) -> pd.DataFrame:
    program_col = find_column(frame, ["NPI", "Program"])
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    actual_eau_col = find_column(frame, ["Future_EAU_Volume", "EAU Actual", "EAU Current"])
    target_eau_col = find_column(frame, ["Baseline_EAU-Volume", "EAU Target", "EAU Baseline"])
    actual_headcount_col = find_column(frame, ["Future_Headcount", "Headcount Actual"])
    target_headcount_col = find_column(frame, ["Baseline_Headcount", "Headcount Target"])
    actual_part_col = find_column(frame, ["Future_Part_Count", "Part Count Actual"])
    target_part_col = find_column(frame, ["Baseline_Part_Count", "Part Count Target"])
    fishbone_col = find_column(frame, ["Fishbone_Status", "Fishbone Status"])
    paag_col = find_column(frame, ["PAAG_Status", "At-a-Glance_Status", "At-a-Glance Status", "PAAG Status"])
    pin_map_col = find_column(frame, ["Pin_Map_Status", "Pin Map Status"])
    wins_col = find_column(frame, ["AIE_Wins_Opportunities", "AIE Wins Opportunities"])
    if not all([program_col, factory_col, actual_headcount_col, target_headcount_col]):
        raise ValueError("NPI metrics sheet is missing required columns.")
    source_columns = [
        program_col,
        factory_col,
        year_col or program_col,
        month_col or program_col,
        actual_eau_col or actual_headcount_col,
        target_eau_col or target_headcount_col,
        actual_headcount_col,
        target_headcount_col,
        actual_part_col or target_headcount_col,
        target_part_col or target_headcount_col,
        fishbone_col or program_col,
        paag_col or program_col,
        pin_map_col or program_col,
        wins_col or program_col,
    ]
    df = frame[source_columns].copy()
    df.columns = [
        "Program",
        "Factory",
        "Year",
        "Month",
        "EAU Actual",
        "EAU Target",
        "Headcount Actual",
        "Headcount Target",
        "Part Count Actual",
        "Part Count Target",
        "Fishbone Status",
        "PAAG Status",
        "Pin Map Status",
        "AIE Wins Opportunities",
    ]
    df["Program"] = df["Program"].astype(str)
    df["Factory"] = df["Factory"].astype(str)
    df["Date"] = pd.to_datetime(df["Year"].astype(str) + "-" + df["Month"].astype(str).str.zfill(2) + "-01", errors="coerce")
    for col in ["EAU Actual", "EAU Target", "Headcount Actual", "Headcount Target", "Part Count Actual", "Part Count Target"]:
        df[col] = coerce_numeric(df[col]).fillna(0)
    for col in ["Fishbone Status", "PAAG Status", "Pin Map Status", "AIE Wins Opportunities"]:
        df[col] = df[col].fillna("").astype(str)
    return df


def build_npi_ergo_reds_df(frame: pd.DataFrame) -> pd.DataFrame:
    program_col = find_column(frame, ["NPI", "Program"])
    factory_col = find_column(frame, ["Site", "Factory"])
    year_col = find_column(frame, ["Year"])
    month_col = find_column(frame, ["Month"])
    started_col = find_column(frame, ["Started_Ergo_Assessments", "Started"])
    open_col = find_column(frame, ["Open_Ergo_Assessments", "Open"])
    pending_col = find_column(frame, ["Pending_Ergo_Assessments", "Pending"])
    validation_col = find_column(frame, ["Validation_Ergo_Assessments", "Validation"])
    closed_admin_col = find_column(frame, ["Closed_Ergo_Assessments_Admin", "Closed Administrative"])
    closed_engineering_col = find_column(frame, ["Closed_Ergo_Assessments_Engineering", "Closed Engineering"])
    wins_col = find_column(frame, ["Ergo_Wins_Opportunities", "Ergo Wins Opportunities"])
    if not all([program_col, started_col, open_col, pending_col, validation_col, closed_admin_col, closed_engineering_col]):
        raise ValueError("NPI ergo reds sheet is missing required columns.")
    source_columns = [
        program_col,
        factory_col or program_col,
        year_col or program_col,
        month_col or program_col,
        started_col,
        open_col,
        pending_col,
        validation_col,
        closed_admin_col,
        closed_engineering_col,
        wins_col or program_col,
    ]
    df = frame[source_columns].copy()
    df.columns = ["Program", "Factory", "Year", "Month", "Started", "Open", "Pending", "Validation", "Closed Administrative", "Closed Engineering", "Ergo Wins Opportunities"]
    df["Program"] = df["Program"].astype(str)
    df["Factory"] = df["Factory"].astype(str)
    df["Date"] = pd.to_datetime(df["Year"].astype(str) + "-" + df["Month"].astype(str).str.zfill(2) + "-01", errors="coerce")
    for col in ["Started", "Open", "Pending", "Validation", "Closed Administrative", "Closed Engineering"]:
        df[col] = coerce_numeric(df[col]).fillna(0).astype(int)
    df["Ergo Wins Opportunities"] = df["Ergo Wins Opportunities"].fillna("").astype(str)
    rows = []
    for _, row in df.iterrows():
        for status_col in ["Started", "Open", "Pending", "Validation", "Closed Administrative", "Closed Engineering"]:
            rows.append(
                {
                    "Program": row["Program"],
                    "Factory": row["Factory"],
                    "Date": row["Date"],
                    "Status": status_col,
                    "Count": int(row[status_col]),
                    "Ergo Wins Opportunities": row["Ergo Wins Opportunities"],
                }
            )
    return pd.DataFrame(rows)


def apply_factory_date_filter(frame: pd.DataFrame, selected_factories: list[str], start_date: pd.Timestamp | None, end_date: pd.Timestamp | None) -> pd.DataFrame:
    if frame.empty:
        return frame
    result = frame.copy()
    if selected_factories:
        result = result[result["Factory"].isin(selected_factories)]
    if "Date" in result.columns and start_date is not None and end_date is not None:
        result = result[(result["Date"] >= start_date) & (result["Date"] <= end_date)]
    return result


def apply_program_filter(frame: pd.DataFrame, selected_programs: list[str]) -> pd.DataFrame:
    if frame.empty or not selected_programs:
        return frame
    return frame[frame["Program"].isin(selected_programs)]


def make_injury_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    monthly = (
        frame.groupby([frame["Date"].dt.to_period("M")])
        .agg(Injuries=("Injuries", "sum"), Recordable_Rate=("Recordable Rate", "mean"), Dart_Rate=("DART Rate", "mean"))
        .reset_index()
    )
    monthly["Month"] = monthly["Date"].dt.to_timestamp()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=monthly["Month"], y=monthly["Injuries"], name="Injuries", marker_color="#4C78A8"))
    fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Recordable_Rate"], name="Recordable Rate", mode="lines+markers", yaxis="y2", line=dict(color="#F58518")))
    fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Dart_Rate"], name="DART Rate", mode="lines+markers", yaxis="y2", line=dict(color="#54A24B")))
    fig.update_layout(title="Ergo Injury Trends", xaxis_title="Month", yaxis_title="Injuries", yaxis2=dict(title="Rates", overlaying="y", side="right"), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    return fig


def make_assessment_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    counts = frame.groupby("Risk Level")["Count"].sum().reset_index()
    fig = px.pie(counts, values="Count", names="Risk Level", color="Risk Level", color_discrete_map={"Red": "#E45756", "Yellow": "#F58518", "Green": "#54A24B"})
    fig.update_layout(title="Ergo Assessment Breakdown")
    return fig


def make_npi_headcount_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    summary = frame.groupby("Program", as_index=False)[["Headcount Actual", "Headcount Target"]].sum()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=summary["Program"], y=summary["Headcount Actual"], name="Actual", marker_color="#4C78A8"))
    fig.add_trace(go.Bar(x=summary["Program"], y=summary["Headcount Target"], name="Target", marker_color="#F58518"))
    fig.update_layout(title="Headcount vs Target", barmode="group", xaxis_title="Program", yaxis_title="Headcount")
    return fig


def make_npi_reds_chart(frame: pd.DataFrame) -> go.Figure:
    if frame.empty:
        return go.Figure()
    summary = frame.groupby(["Program", "Status"])["Count"].sum().reset_index()
    fig = px.bar(summary, x="Program", y="Count", color="Status", barmode="stack")
    fig.update_layout(title="Ergo Reds Status by Program", xaxis_title="Program", yaxis_title="Count")
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
        height=290,
        margin=dict(l=16, r=16, t=52, b=42),
        xaxis_title="",
        yaxis_title=yaxis_title,
        legend=dict(orientation="h", yanchor="top", y=-0.28, xanchor="center", x=0.5),
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
    summary = frame.groupby("Status", as_index=False)["Count"].sum()
    summary["Status"] = pd.Categorical(summary["Status"], categories=order, ordered=True)
    summary = summary.sort_values("Status")
    fig = go.Figure(
        go.Bar(
            x=summary["Status"].map(display_labels).fillna(summary["Status"].astype(str)),
            y=summary["Count"],
            marker_color=["#B9C9AE", "#D98F8F", "#E6C26A", "#A8C596", "#D7DCE3", "#8BA1B8"],
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


def format_change(current_value: float, future_value: float) -> str:
    if not current_value:
        return "0.0%"
    return f"{((future_value - current_value) / current_value) * 100:+.1f}%"


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


def render_pdf_report(data: Dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Ergonomics and Industrial Engineering Dashboard", styles["Title"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Full Report (All Factories and Programs)", styles["Heading2"]))
    story.append(Spacer(1, 12))

    production = data.get("production_injuries", pd.DataFrame())
    assessments = data.get("ergo_assessments", pd.DataFrame())
    risk_projects = data.get("risk_projects", pd.DataFrame())
    medical_requests = data.get("medical_requests", pd.DataFrame())
    npi_metrics = data.get("npi_metrics", pd.DataFrame())
    npi_reds = data.get("npi_ergo_reds", pd.DataFrame())

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
        add_chart(make_injury_chart(production), "Production - Ergo Injury Trends")
    if not assessments.empty:
        add_chart(make_assessment_chart(assessments), "Production - Ergo Assessment Breakdown")
    if not risk_projects.empty:
        add_chart(make_assessment_chart(assessments), "Production - Risk Reduction Projects")
    if not medical_requests.empty:
        story.append(Paragraph("Production - Medical Requests Status", styles["Heading3"]))
        story.append(Spacer(1, 6))
        pivot = pd.pivot_table(medical_requests, values="Count", index="Factory", columns="Status", aggfunc="sum", fill_value=0)
        story.append(Paragraph(pivot.to_string(index=True), styles["BodyText"]))
        story.append(Spacer(1, 12))
    if not npi_metrics.empty:
        add_chart(make_npi_headcount_chart(npi_metrics), "NPI - Headcount vs Target")
    if not npi_reds.empty:
        add_chart(make_npi_reds_chart(npi_reds), "NPI - Ergo Reds Status")
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def main() -> None:
    st.title("Ergonomics and Industrial Engineering Dashboard")
    st.caption("Upload the attached Excel template or use it automatically if it is present in the workspace.")

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
    medical_df = data.get("medical_requests", pd.DataFrame()).copy()
    npi_metrics_df = data.get("npi_metrics", pd.DataFrame()).copy()
    npi_reds_df = data.get("npi_ergo_reds", pd.DataFrame()).copy()

    factories = sorted({str(v) for v in pd.concat([prod_df["Factory"], assessment_df["Factory"], risk_project_df["Factory"], medical_df["Factory"], npi_metrics_df["Factory"]], ignore_index=True) if str(v) != "nan" and str(v) != "None"})
    programs = sorted({str(v) for v in npi_metrics_df["Program"].dropna() if str(v) != "nan" and str(v) != "None"})

    with st.sidebar:
        selected_factories = st.multiselect("Factory", factories, default=[])
        selected_programs = st.multiselect("Program", programs, default=[]) if page == "NPI Dashboard" else []
        min_date = pd.Timestamp(prod_df["Date"].min()) if not prod_df.empty else None
        max_date = pd.Timestamp(prod_df["Date"].max()) if not prod_df.empty else None
        if min_date is not None and max_date is not None:
            date_range = st.date_input("Date range", value=(min_date.to_pydatetime().date(), max_date.to_pydatetime().date()), key="date_range")
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_date = pd.Timestamp(date_range[0])
                end_date = pd.Timestamp(date_range[1])
            else:
                start_date = pd.Timestamp(date_range)
                end_date = pd.Timestamp(date_range)
        else:
            st.info("Date filter is unavailable for the current workbook.")
            start_date = None
            end_date = None

    if page == "Production Dashboard":
        prod_filtered = apply_factory_date_filter(prod_df, selected_factories, start_date, end_date)
        assessment_filtered = apply_factory_date_filter(assessment_df, selected_factories, start_date, end_date)
        risk_project_filtered = apply_factory_date_filter(risk_project_df, selected_factories, start_date, end_date)
        medical_filtered = apply_factory_date_filter(medical_df, selected_factories, start_date, end_date)

        st.subheader("Production Dashboard")
        st.markdown("### Ergo Injury Trends")
        st.plotly_chart(make_injury_chart(prod_filtered), use_container_width=True)

        st.markdown("### Ergo Assessment Breakdown")
        st.plotly_chart(make_assessment_chart(assessment_filtered), use_container_width=True)

        st.markdown("### Risk Reduction Projects")
        if not risk_project_filtered.empty:
            total_reduction = risk_project_filtered["Reduction Points"].sum()
            target_points = risk_project_filtered["Target Points"].sum()
            current_month = pd.Timestamp.today().month if not risk_project_filtered.empty else 1
            expected_progress = (current_month / 12) * target_points
            completion_pct = (total_reduction / target_points * 100) if target_points else 0
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Reduction Points", f"{total_reduction:,.0f}")
            col2.metric("Target Points", f"{target_points:,.0f}")
            col3.metric("% Complete", f"{completion_pct:,.1f}%")
            col4.metric("On Track", "Yes" if total_reduction >= expected_progress else "No")
        else:
            st.info("No risk-project data is available for the selected filters.")

        st.markdown("### Medical Requests Status")
        if not medical_filtered.empty:
            pivot = pd.pivot_table(medical_filtered, values="Count", index="Factory", columns="Status", aggfunc="sum", fill_value=0)
            st.dataframe(pivot, use_container_width=True)
        else:
            st.info("No medical requests data is available for the selected filters.")
    else:
        npi_metrics_filtered = apply_factory_date_filter(npi_metrics_df, selected_factories, start_date, end_date)
        npi_metrics_filtered = apply_program_filter(npi_metrics_filtered, selected_programs)
        npi_reds_filtered = apply_program_filter(npi_reds_df, selected_programs)

        st.subheader("NPI Dashboard")
        if npi_metrics_filtered.empty:
            st.info("No NPI data is available for the selected filters.")
        else:
            render_npi_style()
            summary = summarize_npi_dashboard(npi_metrics_filtered, npi_reds_filtered)
            program_names = sorted(npi_metrics_filtered["Program"].dropna().unique().tolist())
            primary_title = program_names[0] if len(program_names) == 1 else f"{len(program_names)} NPI Programs"
            status_row = npi_metrics_filtered.iloc[0]
            part_metrics_table = pd.DataFrame(
                [
                    {"Metric": "EAU (Annual Demand)", "Current": f"{summary['current_eau']:,.0f}", "Future": f"{summary['future_eau']:,.0f}", "% Difference": format_change(summary["current_eau"], summary["future_eau"])},
                    {"Metric": "Part Count", "Current": f"{summary['current_part']:,.0f}", "Future": f"{summary['future_part']:,.0f}", "% Difference": summary["part_change"]},
                    {"Metric": "Parts per EAU", "Current": f"{(summary['current_part'] / summary['current_eau']):.4f}" if summary["current_eau"] else "0.0000", "Future": f"{(summary['future_part'] / summary['future_eau']):.4f}" if summary["future_eau"] else "0.0000", "% Difference": format_change((summary["current_part"] / summary["current_eau"]) if summary["current_eau"] else 0, (summary["future_part"] / summary["future_eau"]) if summary["future_eau"] else 0)},
                ]
            )
            headcount_metrics_table = pd.DataFrame(
                [
                    {"Metric": "EAU (Annual Demand)", "Current": f"{summary['current_eau']:,.0f}", "Future": f"{summary['future_eau']:,.0f}", "% Difference": format_change(summary["current_eau"], summary["future_eau"])},
                    {"Metric": "Headcount", "Current": f"{summary['current_headcount']:,.0f}", "Future": f"{summary['future_headcount']:,.0f}", "% Difference": summary["headcount_change"]},
                    {"Metric": "HPU (Hours Per Unit)", "Current": f"{summary['current_hpu']:.2f}", "Future": f"{summary['future_hpu']:.2f}", "% Difference": summary["hpu_change"]},
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
            card_cols[0].markdown(make_summary_card("Overall Progress", f"{summary['overall_progress']:.1f}/100", "Progress snapshot", "progress"), unsafe_allow_html=True)
            card_cols[1].markdown(make_summary_card("Part Count Reduction", summary["part_change"], "Change vs current count", "green"), unsafe_allow_html=True)
            card_cols[2].markdown(make_summary_card("Hours Per Unit", summary["hpu_change"], f"Headcount {summary['headcount_change']}", "green"), unsafe_allow_html=True)
            card_cols[3].markdown(make_summary_card("Open Ergo Reds", f"{summary['open_reds']}", "Open issues requiring action", "red"), unsafe_allow_html=True)

            upper_left, upper_right = st.columns(2)
            with upper_left:
                with st.container(border=True):
                    st.dataframe(part_metrics_table, use_container_width=True, hide_index=True)
                    st.plotly_chart(
                        make_npi_comparison_chart(
                            npi_metrics_filtered,
                            "Part Count Target",
                            "Part Count Actual",
                            "Part Count Reduction",
                            "Part Count",
                        ),
                        use_container_width=True,
                    )
            with upper_right:
                with st.container(border=True):
                    st.dataframe(headcount_metrics_table, use_container_width=True, hide_index=True)
                    st.plotly_chart(
                        make_npi_comparison_chart(
                            npi_metrics_filtered,
                            "Headcount Target",
                            "Headcount Actual",
                            "Headcount Reduction",
                            "Headcount",
                            current_color="#B5CAE3",
                            future_color="#5B82B1",
                        ),
                        use_container_width=True,
                    )

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
                    st.plotly_chart(make_npi_detail_snapshot_chart(npi_reds_filtered), use_container_width=True)
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
    if st.sidebar.button("Export Full Report (PDF)"):
        try:
            pdf_bytes = render_pdf_report(data)
        except RuntimeError as exc:
            st.sidebar.error(str(exc))
        else:
            st.sidebar.download_button("Download PDF", data=pdf_bytes, file_name="ergonomics_dashboard_report.pdf", mime="application/pdf")


if __name__ == "__main__":
    main()
