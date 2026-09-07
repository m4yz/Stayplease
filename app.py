import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="StayPlease Operational Intelligence",
    page_icon="🏨",
    layout="wide"
)

EXPECTED_COLUMNS = [
    "Location", "Task", "Quantity", "Req Department", "Requestor",
    "Department", "Assignee", "Status", "To Do", "Doing",
    "Pause Time", "Resume Time", "Done", "Time Usage",
    "Pause Duration", "Pause Reason", "Priority", "Message", "Extra"
]


def clean_team_name(sheet):
    name = str(sheet).strip()
    if "_" in name:
        first, rest = name.split("_", 1)
        if first.isdigit():
            return rest
    return name


def read_single_excel(uploaded_file):
    frames = []

    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception:
        return pd.DataFrame()

    for sheet in xls.sheet_names:
        if str(sheet).strip().lower() == "deleted tasks":
            continue

        try:
            # StayPlease report format: row 3 contains column headers
            raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=2)
            raw = raw.dropna(how="all")

            if raw.empty:
                continue

            # Remove repeated headers and empty-looking rows
            first_col = raw.iloc[:, 0].astype(str).str.strip()
            raw = raw[first_col.ne("Location")]

            # Map available columns safely
            raw = raw.iloc[:, :len(EXPECTED_COLUMNS)].copy()
            raw.columns = EXPECTED_COLUMNS[:len(raw.columns)]

            for col in EXPECTED_COLUMNS:
                if col not in raw.columns:
                    raw[col] = pd.NA

            raw = raw[EXPECTED_COLUMNS]
            raw["Team"] = clean_team_name(sheet)
            raw["Source File"] = uploaded_file.name

            frames.append(raw)

        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def load_all_files(uploaded_files):
    frames = []

    for file in uploaded_files:
        df_file = read_single_excel(file)
        if not df_file.empty:
            frames.append(df_file)

    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    # Text normalization
    text_cols = [
        "Location", "Task", "Req Department", "Requestor", "Department",
        "Assignee", "Status", "Priority", "Pause Reason", "Message",
        "Team", "Source File"
    ]

    for col in text_cols:
        df[col] = df[col].astype("string").str.strip()

    # Status normalization
    status_map = {
        "todo": "To Do",
        "to do": "To Do",
        "doing": "Doing",
        "in progress": "Doing",
        "pause": "Paused",
        "paused": "Paused",
        "done": "Done",
        "completed": "Done"
    }

    df["Status"] = (
        df["Status"]
        .astype("string")
        .str.strip()
        .str.lower()
        .map(status_map)
        .fillna(df["Status"])
    )

    # Date conversion
    date_cols = ["To Do", "Doing", "Pause Time", "Resume Time", "Done"]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")

    # Resolution time = Done - To Do
    df["Resolution Hours"] = (
        (df["Done"] - df["To Do"]).dt.total_seconds() / 3600
    )

    # Remove invalid negative durations and extreme accidental values
    df.loc[df["Resolution Hours"] < 0, "Resolution Hours"] = pd.NA

    # Reporting date
    df["Report Date"] = df["Done"].fillna(df["To Do"])

    # Open task age
    now = pd.Timestamp.now()
    df["Open Age Hours"] = (
        (now - df["To Do"]).dt.total_seconds() / 3600
    )
    df.loc[df["Status"] == "Done", "Open Age Hours"] = pd.NA
    df.loc[df["Open Age Hours"] < 0, "Open Age Hours"] = pd.NA

    # Hour and weekday analytics
    df["Request Hour"] = df["To Do"].dt.hour
    df["Request Day"] = df["To Do"].dt.day_name()

    return df


def format_duration(hours):
    if pd.isna(hours):
        return "-"
    hours = float(hours)
    if hours < 1:
        return f"{hours * 60:.0f} min"
    if hours < 24:
        return f"{hours:.1f} hrs"
    return f"{hours / 24:.1f} days"


def top_n_chart(data, group_col, title, n=10):
    if data.empty or group_col not in data.columns:
        st.info("No data available.")
        return

    chart_data = (
        data.dropna(subset=[group_col])
        .groupby(group_col)
        .size()
        .reset_index(name="Tasks")
        .sort_values("Tasks", ascending=False)
        .head(n)
        .sort_values("Tasks", ascending=True)
    )

    if chart_data.empty:
        st.info("No data available.")
        return

    fig = px.bar(
        chart_data,
        x="Tasks",
        y=group_col,
        orientation="h",
        text="Tasks",
        title=title
    )
    fig.update_layout(showlegend=False, yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)


# =========================================================
# HEADER
# =========================================================

st.title("🏨 StayPlease Operational Intelligence")
st.caption("Task • Request • Defect • Resolution Performance Dashboard")


# =========================================================
# DATA UPLOAD + GLOBAL FILTERS
# =========================================================

with st.sidebar:
    st.header("📤 Data Upload")

    uploaded_files = st.file_uploader(
        "Upload Task Report Excel Files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Multiple Excel reports can be uploaded and consolidated automatically."
    )

if not uploaded_files:
    st.info("👈 Upload one or multiple StayPlease Task Report Excel files to start.")
    st.stop()

with st.spinner("Reading and consolidating Excel files..."):
    df = load_all_files(uploaded_files)

if df.empty:
    st.error("No valid task data was found. Please check the Excel report format.")
    st.stop()


with st.sidebar:
    st.divider()
    st.header("🎛️ Global Filters")

    source_options = sorted(df["Source File"].dropna().unique().tolist())
    selected_sources = st.multiselect(
        "📁 Source File",
        source_options,
        default=source_options
    )

    team_options = sorted(df["Team"].dropna().unique().tolist())
    selected_teams = st.multiselect(
        "🏢 Team",
        team_options,
        default=team_options
    )

    status_options = sorted(df["Status"].dropna().unique().tolist())
    selected_statuses = st.multiselect(
        "📌 Status",
        status_options,
        default=status_options
    )

    priority_options = sorted(df["Priority"].dropna().unique().tolist())
    selected_priorities = st.multiselect(
        "🔥 Priority",
        priority_options,
        default=priority_options
    )

    valid_dates = df["Report Date"].dropna()
    selected_dates = None

    if not valid_dates.empty:
        selected_dates = st.date_input(
            "📅 Report Date Range",
            value=(valid_dates.min().date(), valid_dates.max().date())
        )


# =========================================================
# APPLY GLOBAL FILTERS
# =========================================================

filtered = df.copy()

if selected_sources:
    filtered = filtered[filtered["Source File"].isin(selected_sources)]

if selected_teams:
    filtered = filtered[filtered["Team"].isin(selected_teams)]

if selected_statuses:
    filtered = filtered[filtered["Status"].isin(selected_statuses)]

if selected_priorities:
    filtered = filtered[
        filtered["Priority"].isna()
        | filtered["Priority"].isin(selected_priorities)
    ]

if selected_dates and len(selected_dates) == 2:
    start_date = pd.Timestamp(selected_dates[0])
    end_date = pd.Timestamp(selected_dates[1]) + pd.Timedelta(days=1)

    filtered = filtered[
        filtered["Report Date"].isna()
        | (
            (filtered["Report Date"] >= start_date)
            & (filtered["Report Date"] < end_date)
        )
    ]


# =========================================================
# MAIN NAVIGATION
# =========================================================

overview_tab, operations_tab, defects_tab, team_tab, explorer_tab = st.tabs([
    "🏠 Executive Overview",
    "📊 Operational Analytics",
    "🔧 Defect Analytics",
    "👥 Team Performance",
    "📋 Task Explorer"
])


# =========================================================
# 1. EXECUTIVE OVERVIEW
# =========================================================

with overview_tab:

    total_tasks = len(filtered)
    completed = int((filtered["Status"] == "Done").sum())
    open_tasks = total_tasks - completed
    urgent_open = int(
        (
            (filtered["Priority"].fillna("").str.contains("urgent", case=False))
            & (filtered["Status"] != "Done")
        ).sum()
    )

    completion_rate = (completed / total_tasks * 100) if total_tasks else 0

    completed_data = filtered[
        (filtered["Status"] == "Done")
        & filtered["Resolution Hours"].notna()
    ]

    avg_resolution = completed_data["Resolution Hours"].mean()

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    c1.metric("📋 Total Tasks", f"{total_tasks:,}")
    c2.metric("✅ Completed", f"{completed:,}")
    c3.metric("🔄 Open Tasks", f"{open_tasks:,}")
    c4.metric("🎯 Completion Rate", f"{completion_rate:.1f}%")
    c5.metric("⏱️ Avg Resolution", format_duration(avg_resolution))
    c6.metric("🚨 Urgent Open", f"{urgent_open:,}")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        status_counts = (
            filtered["Status"]
            .fillna("Unknown")
            .value_counts()
            .reset_index()
        )
        status_counts.columns = ["Status", "Tasks"]

        fig_status = px.pie(
            status_counts,
            names="Status",
            values="Tasks",
            hole=0.55,
            title="Task Status Distribution"
        )
        st.plotly_chart(fig_status, use_container_width=True)

    with col2:
        team_counts = (
            filtered.groupby("Team")
            .size()
            .reset_index(name="Tasks")
            .sort_values("Tasks", ascending=True)
        )

        fig_team = px.bar(
            team_counts,
            x="Tasks",
            y="Team",
            orientation="h",
            text="Tasks",
            title="Tasks by Team"
        )
        fig_team.update_layout(yaxis_title="")
        st.plotly_chart(fig_team, use_container_width=True)

    trend = filtered.dropna(subset=["Report Date"]).copy()

    if not trend.empty:
        trend["Date"] = trend["Report Date"].dt.date

        daily = (
            trend.groupby(["Date", "Status"])
            .size()
            .reset_index(name="Tasks")
        )

        fig_trend = px.bar(
            daily,
            x="Date",
            y="Tasks",
            color="Status",
            barmode="stack",
            title="Daily Task Activity"
        )
        st.plotly_chart(fig_trend, use_container_width=True)


# =========================================================
# 2. OPERATIONAL ANALYTICS
# =========================================================

with operations_tab:

    st.subheader("📊 Operational Analytics")

    col1, col2 = st.columns(2)

    with col1:
        top_n_chart(
            filtered,
            "Task",
            "🏆 Top 10 Most Requested Tasks"
        )

    with col2:
        top_n_chart(
            filtered,
            "Location",
            "📍 Top 10 Locations / Rooms by Task Volume"
        )

    st.divider()

    col3, col4 = st.columns(2)

    with col3:
        req_dept = (
            filtered.dropna(subset=["Req Department"])
            .groupby("Req Department")
            .size()
            .reset_index(name="Tasks")
            .sort_values("Tasks", ascending=False)
        )

        if not req_dept.empty:
            fig_req = px.bar(
                req_dept,
                x="Req Department",
                y="Tasks",
                text="Tasks",
                title="Requests by Requesting Department"
            )
            st.plotly_chart(fig_req, use_container_width=True)
        else:
            st.info("No requesting department data available.")

    with col4:
        hour_data = filtered.dropna(subset=["Request Hour"])

        if not hour_data.empty:
            hourly = (
                hour_data.groupby("Request Hour")
                .size()
                .reset_index(name="Tasks")
            )

            fig_hour = px.bar(
                hourly,
                x="Request Hour",
                y="Tasks",
                title="🕒 Peak Request Hour"
            )
            fig_hour.update_xaxes(dtick=1)
            st.plotly_chart(fig_hour, use_container_width=True)
        else:
            st.info("No request time data available.")


# =========================================================
# 3. DEFECT ANALYTICS
# =========================================================

with defects_tab:

    st.subheader("🔧 Defect Analytics")

    # User-selectable definition instead of hardcoding assumptions
    defect_mode = st.radio(
        "Defect data source",
        ["Engineering Team", "Engineering Department", "All Tasks"],
        horizontal=True,
        help="Choose how defects should be identified from your current Excel structure."
    )

    if defect_mode == "Engineering Team":
        defect_df = filtered[
            filtered["Team"]
            .fillna("")
            .str.contains("engineering", case=False, na=False)
        ].copy()

    elif defect_mode == "Engineering Department":
        defect_df = filtered[
            filtered["Department"]
            .fillna("")
            .str.contains("engineering", case=False, na=False)
        ].copy()

    else:
        defect_df = filtered.copy()

    d1, d2, d3 = st.columns(3)

    defect_total = len(defect_df)
    defect_done = int((defect_df["Status"] == "Done").sum())
    defect_avg = defect_df.loc[
        defect_df["Status"] == "Done", "Resolution Hours"
    ].mean()

    d1.metric("🔧 Total Defect Tasks", f"{defect_total:,}")
    d2.metric("✅ Resolved Defects", f"{defect_done:,}")
    d3.metric("⏱️ Avg Resolution", format_duration(defect_avg))

    col1, col2 = st.columns(2)

    with col1:
        top_n_chart(
            defect_df,
            "Task",
            "🔧 Top 10 Defects"
        )

    with col2:
        top_n_chart(
            defect_df,
            "Location",
            "🚨 Top 10 Problematic Rooms / Locations"
        )

    st.divider()

    st.subheader("🐢 Top 10 Longest Resolution Tasks")

    longest = (
        defect_df[
            (defect_df["Status"] == "Done")
            & defect_df["Resolution Hours"].notna()
        ]
        .sort_values("Resolution Hours", ascending=False)
        .head(10)
    )

    if longest.empty:
        st.info("No completed defect tasks with valid resolution time.")
    else:
        longest_display = longest[
            ["Location", "Task", "Assignee", "Priority",
             "To Do", "Done", "Resolution Hours"]
        ].copy()

        longest_display["Resolution Time"] = (
            longest_display["Resolution Hours"].apply(format_duration)
        )

        st.dataframe(
            longest_display.drop(columns=["Resolution Hours"]),
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# 4. TEAM PERFORMANCE
# =========================================================

with team_tab:

    st.subheader("👥 Team Performance")

    if filtered.empty:
        st.info("No data available.")
    else:
        team_summary = (
            filtered.groupby("Team")
            .agg(
                Total_Tasks=("Task", "size"),
                Completed=("Status", lambda x: (x == "Done").sum()),
                To_Do=("Status", lambda x: (x == "To Do").sum()),
                Doing=("Status", lambda x: (x == "Doing").sum()),
                Paused=("Status", lambda x: (x == "Paused").sum()),
                Avg_Resolution_Hours=("Resolution Hours", "mean")
            )
            .reset_index()
        )

        team_summary["Completion Rate %"] = (
            team_summary["Completed"]
            / team_summary["Total_Tasks"]
            * 100
        ).round(1)

        team_summary["Avg Resolution"] = (
            team_summary["Avg_Resolution_Hours"].apply(format_duration)
        )

        st.dataframe(
            team_summary[
                ["Team", "Total_Tasks", "Completed", "To_Do", "Doing",
                 "Paused", "Completion Rate %", "Avg Resolution"]
            ],
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        assignee_summary = (
            filtered.dropna(subset=["Assignee"])
            .groupby(["Team", "Assignee"])
            .agg(
                Tasks=("Task", "size"),
                Completed=("Status", lambda x: (x == "Done").sum()),
                Avg_Resolution_Hours=("Resolution Hours", "mean")
            )
            .reset_index()
            .sort_values("Tasks", ascending=False)
        )

        col1, col2 = st.columns(2)

        with col1:
            top_assignees = assignee_summary.head(15).sort_values("Tasks")

            fig_assignee = px.bar(
                top_assignees,
                x="Tasks",
                y="Assignee",
                color="Team",
                orientation="h",
                title="Top Assignees by Workload"
            )
            st.plotly_chart(fig_assignee, use_container_width=True)

        with col2:
            assignee_completed = (
                assignee_summary.sort_values(
                    ["Completed", "Tasks"],
                    ascending=False
                )
                .head(15)
                .sort_values("Completed")
            )

            fig_completed = px.bar(
                assignee_completed,
                x="Completed",
                y="Assignee",
                color="Team",
                orientation="h",
                title="Top Assignees by Completed Tasks"
            )
            st.plotly_chart(fig_completed, use_container_width=True)


# =========================================================
# 5. TASK EXPLORER
# =========================================================

with explorer_tab:

    st.subheader("📋 Task Explorer")

    # Explorer starts from globally filtered data
    e1, e2, e3 = st.columns(3)

    with e1:
        explorer_teams = sorted(filtered["Team"].dropna().unique().tolist())
        selected_explorer_teams = st.multiselect(
            "🏢 Team",
            explorer_teams,
            default=explorer_teams,
            key="explorer_team"
        )

    with e2:
        explorer_locations = sorted(filtered["Location"].dropna().unique().tolist())
        selected_locations = st.multiselect(
            "📍 Location",
            explorer_locations,
            default=explorer_locations,
            key="explorer_location"
        )

    with e3:
        explorer_assignees = sorted(filtered["Assignee"].dropna().unique().tolist())
        selected_assignees = st.multiselect(
            "👤 Assignee",
            explorer_assignees,
            default=explorer_assignees,
            key="explorer_assignee"
        )

    e4, e5, e6 = st.columns(3)

    with e4:
        explorer_statuses = sorted(filtered["Status"].dropna().unique().tolist())
        selected_explorer_statuses = st.multiselect(
            "📌 Status",
            explorer_statuses,
            default=explorer_statuses,
            key="explorer_status"
        )

    with e5:
        explorer_priorities = sorted(filtered["Priority"].dropna().unique().tolist())
        selected_explorer_priorities = st.multiselect(
            "🔥 Priority",
            explorer_priorities,
            default=explorer_priorities,
            key="explorer_priority"
        )

    with e6:
        search = st.text_input(
            "🔎 Search",
            placeholder="Task, location, requestor, assignee...",
            key="explorer_search"
        )

    display = filtered.copy()

    if selected_explorer_teams:
        display = display[display["Team"].isin(selected_explorer_teams)]

    if selected_locations:
        display = display[display["Location"].isin(selected_locations)]

    if selected_assignees:
        display = display[display["Assignee"].isin(selected_assignees)]

    if selected_explorer_statuses:
        display = display[display["Status"].isin(selected_explorer_statuses)]

    if selected_explorer_priorities:
        display = display[
            display["Priority"].isna()
            | display["Priority"].isin(selected_explorer_priorities)
        ]

    if search:
        search_columns = [
            "Task", "Location", "Requestor", "Assignee",
            "Message", "Req Department", "Department"
        ]

        mask = pd.Series(False, index=display.index)

        for col in search_columns:
            mask |= (
                display[col]
                .fillna("")
                .astype(str)
                .str.contains(search, case=False, na=False)
            )

        display = display[mask]

    display_columns = [
        "Team", "Location", "Task", "Quantity",
        "Req Department", "Requestor", "Department",
        "Assignee", "Status", "Priority",
        "To Do", "Doing", "Done",
        "Resolution Hours", "Pause Reason", "Message"
    ]

    explorer_table = display[display_columns].copy()
    explorer_table["Resolution Time"] = (
        explorer_table["Resolution Hours"].apply(format_duration)
    )
    explorer_table = explorer_table.drop(columns=["Resolution Hours"])

    st.caption(f"Showing {len(explorer_table):,} task(s)")

    st.dataframe(
        explorer_table.sort_values("To Do", ascending=False),
        use_container_width=True,
        hide_index=True,
        height=600
    )

    csv = explorer_table.to_csv(index=False).encode("utf-8")

    st.download_button(
        "⬇️ Download Filtered Data (CSV)",
        data=csv,
        file_name="stayplease_filtered_tasks.csv",
        mime="text/csv"
    )


# =========================================================
# FOOTER
# =========================================================

st.divider()
st.caption(
    f"🏨 StayPlease Operational Intelligence | "
    f"📊 {len(filtered):,} filtered tasks | "
    f"🏢 {filtered['Team'].nunique()} team(s) | "
    f"📁 {filtered['Source File'].nunique()} source file(s)"
)
