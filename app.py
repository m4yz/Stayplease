import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="StayPlease Dashboard", page_icon="📊", layout="wide")

DEFAULT_FILE = "Task Report_639243879690744740.xlsx"
EXPECTED_COLUMNS = [
    "Location","Task","Quantity","Req Department","Requestor","Department",
    "Assignee","Status","To Do","Doing","Pause Time","Resume Time","Done",
    "Time Usage","Pause Duration","Pause Reason","Priority","Message","Extra"
]

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    frames = []

    for sheet in xls.sheet_names:
        if sheet == "Deleted Tasks":
            continue

        raw = pd.read_excel(file, sheet_name=sheet, header=2)
        raw = raw.dropna(how="all")

        # Remove accidental repeated header rows inside exported reports
        raw = raw[raw.iloc[:, 0].astype(str).str.strip().ne("Location")]

        raw.columns = EXPECTED_COLUMNS[:len(raw.columns)]
        for col in EXPECTED_COLUMNS:
            if col not in raw.columns:
                raw[col] = pd.NA

        raw["Team"] = sheet.replace("1_", "").replace("2_", "").replace("3_", "").replace("4_", "").replace("5_", "")
        frames.append(raw[EXPECTED_COLUMNS + ["Team"]])

    df = pd.concat(frames, ignore_index=True)

    # Normalize text
    for col in ["Location","Task","Req Department","Requestor","Department","Assignee",
                "Status","Priority","Pause Reason","Message","Team"]:
        df[col] = df[col].astype("string").str.strip()

    # Normalize status values
    status_map = {"ToDo": "To Do", "Todo": "To Do", "Done": "Done",
                  "Doing": "Doing", "Pause": "Paused"}
    df["Status"] = df["Status"].replace(status_map)

    # Convert dates
    for col in ["To Do", "Doing", "Pause Time", "Resume Time", "Done"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # Main reporting date: Done first, otherwise To Do
    df["Report Date"] = df["Done"].fillna(df["To Do"])
    df["Report Date"] = pd.to_datetime(df["Report Date"], errors="coerce")

    # Calculate completion duration where possible
    df["Duration Hours"] = (df["Done"] - df["To Do"]).dt.total_seconds() / 3600
    df.loc[df["Duration Hours"] < 0, "Duration Hours"] = pd.NA

    # Numeric quantity
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")

    return df


st.title("📊 StayPlease Task Dashboard")
st.caption("Operational task performance dashboard")

with st.sidebar:
    st.header("Data Source")
    uploaded = st.file_uploader("Upload Task Report (.xlsx)", type=["xlsx"])

    if uploaded:
        source = uploaded
    else:
        source = DEFAULT_FILE

    st.divider()
    st.header("Filters")

try:
    df = load_data(source)
except FileNotFoundError:
    st.warning("Upload the Task Report Excel file using the sidebar.")
    st.stop()
except Exception as e:
    st.error(f"Unable to read the Excel report: {e}")
    st.stop()

# Sidebar filters
with st.sidebar:
    teams = sorted(df["Team"].dropna().unique().tolist())
    selected_teams = st.multiselect("Department / Team", teams, default=teams)

    statuses = sorted(df["Status"].dropna().unique().tolist())
    selected_statuses = st.multiselect("Status", statuses, default=statuses)

    priorities = sorted(df["Priority"].dropna().unique().tolist())
    selected_priorities = st.multiselect("Priority", priorities, default=priorities)

    valid_dates = df["Report Date"].dropna()
    date_range = None
    if not valid_dates.empty:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        date_range = st.date_input("Report Date Range", value=(min_date, max_date))

filtered = df[
    df["Team"].isin(selected_teams) &
    df["Status"].isin(selected_statuses)
].copy()

if selected_priorities:
    filtered = filtered[
        filtered["Priority"].isna() | filtered["Priority"].isin(selected_priorities)
    ]

if date_range and len(date_range) == 2:
    start_date, end_date = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
    filtered = filtered[
        filtered["Report Date"].isna() |
        ((filtered["Report Date"] >= start_date) & (filtered["Report Date"] < end_date))
    ]

# KPIs
total = len(filtered)
done = int((filtered["Status"] == "Done").sum())
todo = int((filtered["Status"] == "To Do").sum())
doing = int((filtered["Status"] == "Doing").sum())
paused = int((filtered["Status"] == "Paused").sum())
completion_rate = (done / total * 100) if total else 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Tasks", f"{total:,}")
k2.metric("Completed", f"{done:,}")
k3.metric("To Do", f"{todo:,}")
k4.metric("In Progress", f"{doing:,}")
k5.metric("Completion Rate", f"{completion_rate:.1f}%")

st.divider()

tab1, tab2, tab3 = st.tabs(["📈 Overview", "👥 Team Performance", "📋 Task Explorer"])

with tab1:
    c1, c2 = st.columns(2)

    status_counts = filtered["Status"].fillna("Unknown").value_counts().reset_index()
    status_counts.columns = ["Status", "Tasks"]
    fig_status = px.pie(status_counts, names="Status", values="Tasks",
                        title="Task Status Distribution", hole=0.55)
    c1.plotly_chart(fig_status, use_container_width=True)

    team_counts = filtered.groupby("Team").size().reset_index(name="Tasks").sort_values("Tasks", ascending=True)
    fig_team = px.bar(team_counts, x="Tasks", y="Team", orientation="h",
                      title="Tasks by Team", text="Tasks")
    c2.plotly_chart(fig_team, use_container_width=True)

    trend = filtered.dropna(subset=["Report Date"]).copy()
    if not trend.empty:
        trend["Date"] = trend["Report Date"].dt.date
        daily = trend.groupby(["Date", "Status"]).size().reset_index(name="Tasks")
        fig_trend = px.bar(daily, x="Date", y="Tasks", color="Status",
                           title="Daily Task Activity", barmode="stack")
        st.plotly_chart(fig_trend, use_container_width=True)

    urgent = filtered[
        filtered["Priority"].fillna("").str.contains("Urgent", case=False, na=False)
    ]
    st.subheader("🚨 Urgent Tasks")
    if urgent.empty:
        st.success("No urgent tasks in the current selection.")
    else:
        st.dataframe(
            urgent[["Team","Location","Task","Assignee","Status","Priority","To Do","Done"]]
            .sort_values("To Do", ascending=False),
            use_container_width=True,
            hide_index=True
        )

with tab2:
    team_summary = (
        filtered.groupby("Team")
        .agg(
            Total_Tasks=("Task", "size"),
            Completed=("Status", lambda x: (x == "Done").sum()),
            To_Do=("Status", lambda x: (x == "To Do").sum()),
            Doing=("Status", lambda x: (x == "Doing").sum()),
            Paused=("Status", lambda x: (x == "Paused").sum()),
            Avg_Duration_Hours=("Duration Hours", "mean")
        )
        .reset_index()
    )
    team_summary["Completion Rate %"] = (
        team_summary["Completed"] / team_summary["Total_Tasks"] * 100
    ).round(1)
    team_summary["Avg_Duration_Hours"] = team_summary["Avg_Duration_Hours"].round(2)

    st.dataframe(team_summary, use_container_width=True, hide_index=True)

    assignee = filtered.dropna(subset=["Assignee"])
    if not assignee.empty:
        assignee_summary = (
            assignee.groupby(["Team","Assignee"])
            .size().reset_index(name="Tasks")
            .sort_values("Tasks", ascending=False)
            .head(20)
        )
        fig_assignee = px.bar(
            assignee_summary, x="Tasks", y="Assignee", color="Team",
            orientation="h", title="Top 20 Assignees by Task Volume"
        )
        st.plotly_chart(fig_assignee, use_container_width=True)

with tab3:
    search = st.text_input("🔎 Search task, location, requestor, assignee, or message")

    display = filtered.copy()
    if search:
        cols = ["Task","Location","Requestor","Assignee","Message","Req Department"]
        mask = pd.Series(False, index=display.index)
        for col in cols:
            mask |= display[col].fillna("").astype(str).str.contains(search, case=False, na=False)
        display = display[mask]

    display_cols = [
        "Team","Location","Task","Quantity","Req Department","Requestor",
        "Department","Assignee","Status","Priority","To Do","Doing","Done",
        "Duration Hours","Pause Reason","Message"
    ]
    st.dataframe(
        display[display_cols].sort_values("To Do", ascending=False),
        use_container_width=True,
        hide_index=True,
        height=600
    )

    csv = display[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download Filtered Data (CSV)",
        data=csv,
        file_name="stayplease_filtered_tasks.csv",
        mime="text/csv"
    )

st.caption(f"Showing {len(filtered):,} tasks across {filtered['Team'].nunique()} team(s).")
