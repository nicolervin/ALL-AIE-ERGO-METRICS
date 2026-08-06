from __future__ import annotations

import math
import re
from typing import Iterable

import pandas as pd


RATE_BASE = 200_000.0

RAW_SITE_MAP = {
    "Louisville, AP1, KY": "AP1",
    "Louisville, AP3, KY": "AP3",
    "Louisville, AP4, KY": "AP4",
    "Louisville, AP5, KY": "AP5",
    "Decatur, AL": "DPO",
    "Selmer, TN - Monogram": "MRO",
    "Lafayette, GA": "RPO",
    "Camden, SC": "CPO",
    "Factory Service": "BODE",
}

KNOWN_SITE_CODES = {
    "AP1",
    "AP2",
    "AP3",
    "AP4",
    "AP5",
    "BOD",
    "BODE",
    "CPO",
    "DPO",
    "MRO",
    "RPO",
}

COUNT_COLUMNS = [
    "EHS Recordables",
    "EHS DART Cases",
    "Ergo All Cases",
    "Ergo First Aid",
    "Ergo Recordable",
    "Ergo DAFW",
    "Ergo Restricted",
    "Ergo Other Recordable",
    "Ergo DART",
]

R12_COUNT_COLUMNS = {
    "Hours": "Hours R12",
    "EHS Recordables": "EHS Recordables R12",
    "EHS DART Cases": "EHS DART Cases R12",
    "Ergo All Cases": "Ergo All Cases R12",
    "Ergo First Aid": "Ergo First Aid R12",
    "Ergo Recordable": "Ergo Recordable R12",
    "Ergo DAFW": "Ergo DAFW R12",
    "Ergo Restricted": "Ergo Restricted R12",
    "Ergo Other Recordable": "Ergo Other Recordable R12",
    "Ergo DART": "Ergo DART R12",
}


def _clean_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_site(value: object, *, raw_site: bool = False) -> str:
    site = _clean_header(value)
    if not site or site.lower() == "nan":
        return "Other"
    if raw_site:
        return RAW_SITE_MAP.get(site, "Other")
    upper = site.upper()
    if upper in {"BOD", "BODE", "BODEWELL", "FS", "FACTORY SERVICE"}:
        return "BODE"
    if upper in KNOWN_SITE_CODES:
        return upper
    if upper in {"GEA", "GE APPLIANCES"}:
        return "GEA"
    return "Other"


def normalize_sites(frame: pd.DataFrame, column: str = "Factory") -> pd.DataFrame:
    result = frame.copy()
    if column in result.columns:
        result[column] = result[column].map(normalize_site)
    return result


def fiscal_year_start(year: int) -> pd.Timestamp:
    """Return the Sunday starting fiscal week 1.

    The supplied calendar anchors fiscal 2026 at 2025-12-28. The available
    history uses 52 weeks for 2023-2025; fiscal 2026 contains 53 weeks.
    """

    anchor = pd.Timestamp("2025-12-28")
    if year <= 2026:
        return anchor - pd.Timedelta(weeks=52 * (2026 - year))
    if year == 2027:
        return anchor + pd.Timedelta(weeks=53)
    return anchor + pd.Timedelta(weeks=53 + (year - 2027) * 52)


def _required_columns(frame: pd.DataFrame, names: Iterable[str], sheet: str) -> None:
    missing = [name for name in names if name not in frame.columns]
    if missing:
        raise ValueError(f"{sheet} is missing required columns: {', '.join(missing)}")


def _latest_majority_week(frame: pd.DataFrame) -> tuple[int, int]:
    latest_year = int(frame["Year"].max())
    current = frame[frame["Year"] == latest_year]
    site_count = max(int(current["Factory"].nunique()), 1)
    required_coverage = max(math.ceil(site_count / 2), 1)
    coverage = current.groupby("FW")["Factory"].nunique()
    eligible = coverage[coverage >= required_coverage]
    if eligible.empty:
        return latest_year, int(current["FW"].max())
    return latest_year, int(eligible.index.max())


def build_ehs_monthly(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp]:
    required = [
        "Site",
        "FW",
        "Year",
        "Hours_Worked_This_Month",
        "EHS_recordable_cases",
        "EHS_DAFW_Cases",
    ]
    _required_columns(frame, required, "EHS_Hours_Injuries")

    data = frame[required].copy()
    data.columns = [
        "Factory",
        "FW",
        "Year",
        "Hours YTD",
        "EHS Recordables YTD",
        "EHS DART Cases YTD",
    ]
    data["Factory"] = data["Factory"].map(normalize_site)
    data = data[data["Factory"] != "GEA"]
    data["FW"] = pd.to_numeric(
        data["FW"].astype(str).str.extract(r"(\d+)")[0], errors="coerce"
    )
    data["Year"] = pd.to_numeric(data["Year"], errors="coerce")
    for column in ["Hours YTD", "EHS Recordables YTD", "EHS DART Cases YTD"]:
        data[column] = pd.to_numeric(data[column], errors="coerce")
    data = data.dropna(subset=["Factory", "FW", "Year", "Hours YTD"])
    data[["FW", "Year"]] = data[["FW", "Year"]].astype(int)

    latest_year, latest_week = _latest_majority_week(data)
    data = data[
        (data["Year"] < latest_year)
        | ((data["Year"] == latest_year) & (data["FW"] <= latest_week))
    ].copy()
    data = data.sort_values(["Factory", "Year", "FW"])

    ytd_columns = ["Hours YTD", "EHS Recordables YTD", "EHS DART Cases YTD"]
    increment_columns = ["Hours", "EHS Recordables", "EHS DART Cases"]
    grouped = data.groupby(["Factory", "Year"], sort=False)
    for ytd, increment in zip(ytd_columns, increment_columns):
        data[increment] = grouped[ytd].diff()
        first_in_group = grouped.cumcount() == 0
        data.loc[first_in_group, increment] = data.loc[first_in_group, ytd]

    data["Week Start"] = [
        fiscal_year_start(year) + pd.Timedelta(weeks=week - 1)
        for year, week in zip(data["Year"], data["FW"])
    ]
    data["Date"] = data["Week Start"].dt.to_period("M").dt.to_timestamp()

    latest_week_start = data.loc[data["Year"] == latest_year, "Week Start"].max()
    available_through = latest_week_start + pd.Timedelta(days=6)
    latest_complete_month = (
        available_through.to_period("M") - 1
        if available_through < available_through.to_period("M").end_time.normalize()
        else available_through.to_period("M")
    ).to_timestamp()

    monthly = (
        data.groupby(["Date", "Factory"], as_index=False)[increment_columns]
        .sum()
        .sort_values(["Date", "Factory"])
    )
    monthly = monthly[monthly["Date"] <= latest_complete_month]
    return monthly, latest_complete_month


def _flag(frame: pd.DataFrame, columns: list[str]) -> pd.Series:
    result = pd.Series(False, index=frame.index)
    for column in columns:
        values = frame[column]
        numeric = pd.to_numeric(values, errors="coerce").eq(1)
        text = values.astype(str).str.strip().str.lower().isin({"yes", "true", "1"})
        result = result | numeric | text
    return result


def build_ergo_monthly(frame: pd.DataFrame) -> pd.DataFrame:
    required = [
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
    ]
    cleaned_columns = {_clean_header(column): column for column in frame.columns}
    missing = [name for name in required if name not in cleaned_columns]
    if missing:
        raise ValueError(f"Ergo_Injuries is missing required columns: {', '.join(missing)}")

    data = frame[[cleaned_columns[name] for name in required]].copy()
    data.columns = required
    data["Case Date"] = pd.to_datetime(data["Case Date"], errors="coerce")
    data["Factory"] = data["Site"].map(lambda value: normalize_site(value, raw_site=True))
    data = data.dropna(subset=["System Case ID", "Case Date"])
    data = data.drop_duplicates(
        subset=["Factory", "System Case ID"],
        keep="last",
    )

    data["Ergo Recordable"] = (
        data["U.S. OSHA Recordable?"].astype(str).str.strip().str.lower().eq("yes")
    )
    data["Ergo DAFW"] = _flag(data, ["Inj. DAFW?", "Ill. DAFW?"])
    data["Ergo Restricted"] = _flag(data, ["Inj. RWA?", "Ill. RWA?"])
    data["Ergo Other Recordable"] = _flag(
        data, ["Inj. Other Recordable Case?", "Ill. Other Recordable Case?"]
    )
    data["Ergo DART"] = data["Ergo DAFW"] | data["Ergo Restricted"]
    data["Ergo First Aid"] = ~data["Ergo Recordable"]
    data["Ergo All Cases"] = True
    data["Date"] = data["Case Date"].dt.to_period("M").dt.to_timestamp()

    metric_columns = [
        "Ergo All Cases",
        "Ergo First Aid",
        "Ergo Recordable",
        "Ergo DAFW",
        "Ergo Restricted",
        "Ergo Other Recordable",
        "Ergo DART",
    ]
    monthly = (
        data.groupby(["Date", "Factory"], as_index=False)[metric_columns]
        .sum()
        .sort_values(["Date", "Factory"])
    )
    for column in metric_columns:
        monthly[column] = monthly[column].astype(int)
    return monthly


def build_monthly_detail(
    ehs_frame: pd.DataFrame, ergo_frame: pd.DataFrame
) -> tuple[pd.DataFrame, pd.Timestamp]:
    ehs_monthly, latest_complete_month = build_ehs_monthly(ehs_frame)
    ergo_monthly = build_ergo_monthly(ergo_frame)
    detail = pd.merge(
        ehs_monthly,
        ergo_monthly,
        on=["Date", "Factory"],
        how="outer",
    )
    detail = detail[detail["Date"] <= latest_complete_month].copy()
    numeric_columns = ["Hours", *COUNT_COLUMNS]
    for column in numeric_columns:
        if column not in detail:
            detail[column] = 0.0
        detail[column] = pd.to_numeric(detail[column], errors="coerce").fillna(0.0)
    return detail.sort_values(["Date", "Factory"]), latest_complete_month


def available_scopes(detail: pd.DataFrame) -> list[str]:
    preferred = ["GE Appliances", "Supply Chain"]
    sites = sorted(
        site
        for site in detail["Factory"].dropna().astype(str).unique()
        if site not in {"GEA", "Other"}
    )
    return preferred + sites


def aggregate_scope(detail: pd.DataFrame, scope: str) -> pd.DataFrame:
    data = detail.copy()
    if scope == "GE Appliances":
        # Corporate rates must keep the numerator and denominator on the same
        # population. "Other" has injury cases but no matching EHS hours in
        # the supplied workbook, so including it would inflate every rate.
        data = data[~data["Factory"].isin({"GEA", "Other"})]
    elif scope == "Supply Chain":
        data = data[~data["Factory"].isin({"GEA", "BODE", "Other"})]
    else:
        data = data[data["Factory"] == scope]

    if data.empty:
        return pd.DataFrame()

    monthly = data.groupby("Date", as_index=False)[["Hours", *COUNT_COLUMNS]].sum()
    date_index = pd.date_range(monthly["Date"].min(), monthly["Date"].max(), freq="MS")
    monthly = (
        monthly.set_index("Date")
        .reindex(date_index, fill_value=0)
        .rename_axis("Date")
        .reset_index()
    )

    for source, target in R12_COUNT_COLUMNS.items():
        monthly[target] = monthly[source].rolling(12, min_periods=12).sum()

    hours = monthly["Hours R12"].where(monthly["Hours R12"] > 0)
    rate_sources = {
        "EHS Recordable Rate R12": "EHS Recordables R12",
        "EHS DART Rate R12": "EHS DART Cases R12",
        "Ergo All Rate R12": "Ergo All Cases R12",
        "Ergo Recordable Rate R12": "Ergo Recordable R12",
        "Ergo DAFW Rate R12": "Ergo DAFW R12",
        "Ergo DART Rate R12": "Ergo DART R12",
    }
    for rate_column, count_column in rate_sources.items():
        monthly[rate_column] = monthly[count_column] * RATE_BASE / hours

    denominator = monthly["EHS Recordables R12"].where(
        monthly["EHS Recordables R12"] > 0
    )
    monthly["Ergo Share of EHS Recordables"] = (
        monthly["Ergo Recordable R12"] / denominator * 100
    )
    monthly["Scope"] = scope
    return monthly
