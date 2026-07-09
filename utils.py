import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(layout="wide")

st.title("Ergonomics & Industrial Engineering Dashboard")

# Sidebar
st.sidebar.header("Controls")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx"])

page = st.sidebar.selectbox("Select Page", ["Production Dashboard", "NPI Dashboard"])

factory_filter = st.sidebar.multiselect("Factory")
program_filter = st.sidebar.multiselect("Program")

# Load data
@st.cache_data
def load_data(file):
    data = {
        "injuries": pd.read_excel(file, sheet_name="production_injuries"),
        "assessments": pd.read_excel(file, sheet_name="ergo_assessments"),
        "projects": pd.read_excel(file, sheet_name="risk_projects"),
        "medical": pd.read_excel(file, sheet_name="medical_requests"),
        "npi": pd.read_excel(file, sheet_name="npi_metrics"),
        "npi_reds": pd.read_excel(file, sheet_name="npi_ergo_reds"),
    }
    return data

if uploaded_file:
    data = load_data(uploaded_file)

    # ------------------------
    # PRODUCTION DASHBOARD
    # ------------------------
    if page == "Production Dashboard":
        st.header("Production Metrics")

        injuries = data["injuries"]

        # Ergo Injury Trends
        st.subheader("Ergo Injury Trends")

        fig = go.Figure()

        fig.add_bar(
            x=injuries["Date"],
            y=injuries["Injuries"],
            name="Injuries"
        )

        fig.add_trace(go.Scatter(
            x=injuries["Date"],
            y=injuries["Recordable Rate"],
            mode='lines+markers',
            name="Recordable Rate"
        ))

        fig.add_trace(go.Scatter(
            x=injuries["Date"],
            y=injuries["DART Rate"],
            mode='lines+markers',
            name="DART Rate"
        ))

        st.plotly_chart(fig, use_container_width=True)

        # Ergo Assessment Breakdown
        st.subheader("Ergo Assessment Breakdown")

        assess = data["assessments"]
        pie = assess.groupby("Risk Level")["Count"].sum().reset_index()

        fig2 = go.Figure(data=[go.Pie(
            labels=pie["Risk Level"],
            values=pie["Count"]
        )])

        st.plotly_chart(fig2)

        # Risk Reduction Projects
        st.subheader("Risk Reduction Projects")

        proj = data["projects"]

        total_points = proj["Reduction Points"].sum()
        target_points = proj["Target Points"].sum()
        percent = (total_points / target_points) * 100 if target_points else 0

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Reduction Points", int(total_points))
        col2.metric("Target Points", int(target_points))
        col3.metric("% Complete", f"{percent:.1f}%")

        # Medical Requests
        st.subheader("Medical Requests Status")

        med = data["medical"]
        table = med.pivot_table(
            index="Factory",
            columns="Status",
            values="Count",
            aggfunc="sum"
        ).fillna(0)

        st.dataframe(table)

    # ------------------------
    # NPI DASHBOARD
    # ------------------------
    elif page == "NPI Dashboard":
        st.header("NPI Metrics")

        npi = data["npi"]

        # Headcount vs Target
        st.subheader("Headcount vs Target")

        fig3 = go.Figure()

        fig3.add_bar(
            x=npi["Program"],
            y=npi["Headcount Actual"],
            name="Actual"
        )

        fig3.add_bar(
            x=npi["Program"],
            y=npi["Headcount Target"],
            name="Target"
        )

        st.plotly_chart(fig3, use_container_width=True)

        # Ergo Reds Status
        st.subheader("Ergo Reds Status")

        reds = data["npi_reds"]

        fig4 = go.Figure()

        for status in reds["Status"].unique():
            df = reds[reds["Status"] == status]
            fig4.add_bar(
                x=df["Program"],
                y=df["Count"],
                name=status
            )

        st.plotly_chart(fig4, use_container_width=True)

else:
    st.info("Upload an Excel file to begin.")
    