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


def clean_team_name(sheet_name):
    name = str(sheet_name).strip()
    if "_" in name:
        first, rest = name.split("_", 1)
        if first.isdigit():
            return rest.strip()
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
            raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=2)
            raw = raw.dropna(how="all")

            if raw.empty:
                continue

            first_col = raw.iloc[:, 0].astype(str).str.strip()
            raw = raw[first_col.ne("Location")]

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

    text_cols = [
        "Location", "Task", "Req Department", "Requestor", "Department",
        "Assignee", "Status", "Priority", "Pause Reason", "Message",
        "Team", "Source File"
    ]

    for col in text_cols:
        # Critical: Location must remain categorical text, never numeric axis data
        df[col] = df[col].astype("string").str.strip()

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

    normalized = df["Status"].astype("string").str.strip().str.lower()
    df["Status"] = normalized.map(status_map).fillna(df["Status"])

    for col in ["To Do", "Doing", "Pause Time", "Resume Time", "Done"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")

    df["Resolution Hours"] = (
        (df["Done"] - df["To Do"]).dt.total_seconds() / 3600
    )
    df.loc[df["Resolution Hours"] < 0, "Resolution Hours"] = pd.NA

    df["Report Date"] = df["Done"].fillna(df["To Do"])
    df["Request Hour"] = df["To Do"].dt.hour
    df["Request Month"] = df["To Do"].dt.to_period("M").astype("string")

    now = pd.Timestamp.now()
    df["Open Age Hours"] = (now - df["To Do"]).dt.total_seconds() / 3600
    df.loc[df["Status"] == "Done", "Open Age Hours"] = pd.NA
    df.loc[df["Open Age Hours"] < 0, "Open Age Hours"] = pd.NA

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


def top_n_counts(data, group_col, n=10):
    return (
        data.dropna(subset=[group_col])
        .assign(**{group_col: lambda x: x[group_col].astype(str)})
        .groupby(group_col)
        .size()
        .reset_index(name="Tasks")
        .sort_values("Tasks", ascending=False)
        .head(n)
    )


def top_n_chart(data, group_col, title, n=10):
    chart_data = top_n_counts(data, group_col, n)

    if chart_data.empty:
        st.info("No data available.")
        return

    # Reverse order for readable horizontal ranking
    plot_data = chart_data.sort_values("Tasks", ascending=True)

    fig = px.bar(
        plot_data,
        x="Tasks",
        y=group_col,
        orientation="h",
        text="Tasks",
        title=title
    )
    fig.update_layout(
        showlegend=False,
        yaxis_title="",
        xaxis_title="Tasks",
        margin=dict(l=20, r=20, t=55, b=30)
    )
    fig.update_yaxes(type="category")
    st.plotly_chart(fig, use_container_width=True)


# =========================================================
# HEADER + UPLOAD
# =========================================================

st.title("🏨 StayPlease Operational Intelligence")
st.caption("Task • Request • Defect • Resolution Performance Dashboard")

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


# =========================================================
# GLOBAL FILTERS
# =========================================================

with st.sidebar:
    st.divider()
    st.header("🎛️ Global Filters")

    source_options = sorted(df["Source File"].dropna().unique().tolist())
    selected_sources = st.multiselect(
        "📁 Source File", source_options, default=source_options
    )

    team_options = sorted(df["Team"].dropna().unique().tolist())
    selected_teams = st.multiselect(
        "🏢 Team", team_options, default=team_options
    )

    status_options = sorted(df["Status"].dropna().unique().tolist())
    selected_statuses = st.multiselect(
        "📌 Status", status_options, default=status_options
    )

    priority_options = sorted(df["Priority"].dropna().unique().tolist())
    selected_priorities = st.multiselect(
        "🔥 Priority", priority_options, default=priority_options
    )

    valid_dates = df["Report Date"].dropna()
    selected_dates = None

    if not valid_dates.empty:
        selected_dates = st.date_input(
            "📅 Report Date Range",
            value=(valid_dates.min().date(), valid_dates.max().date())
        )


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
# TABS
# =========================================================

overview_tab, operations_tab, defects_tab, team_tab, explorer_tab = st.tabs([
    "🏠 Executive Overview",
    "📊 Operational Analytics",
    "🔧 Defect Analytics",
    "👥 Team Performance",
    "📋 Task Explorer"
])


# =========================================================
# EXECUTIVE OVERVIEW
# =========================================================

with overview_tab:
    total_tasks = len(filtered)
    completed = int((filtered["Status"] == "Done").sum())
    open_tasks = total_tasks - completed
    completion_rate = completed / total_tasks * 100 if total_tasks else 0

    urgent_open = int(
        (
            filtered["Priority"].fillna("")
            .str.contains("urgent", case=False, na=False)
            & (filtered["Status"] != "Done")
        ).sum()
    )

    completed_data = filtered[
        (filtered["Status"] == "Done")
        & filtered["Resolution Hours"].notna()
    ]

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("📋 Total Tasks", f"{total_tasks:,}")
    c2.metric("✅ Completed", f"{completed:,}")
    c3.metric("🔄 Open Tasks", f"{open_tasks:,}")
    c4.metric("🎯 Completion Rate", f"{completion_rate:.1f}%")
    c5.metric("⏱️ Avg Resolution", format_duration(completed_data["Resolution Hours"].mean()))
    c6.metric("🚨 Urgent Open", f"{urgent_open:,}")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        status_counts = (
            filtered["Status"].fillna("Unknown")
            .value_counts()
            .reset_index()
        )
        status_counts.columns = ["Status", "Tasks"]

        fig = px.pie(
            status_counts, names="Status", values="Tasks",
            hole=0.55, title="Task Status Distribution"
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        team_counts = (
            filtered.groupby("Team")
            .size()
            .reset_index(name="Tasks")
            .sort_values("Tasks", ascending=True)
        )

        fig = px.bar(
            team_counts, x="Tasks", y="Team",
            orientation="h", text="Tasks",
            title="Tasks by Team"
        )
        fig.update_layout(yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)


# =========================================================
# OPERATIONAL ANALYTICS
# =========================================================

with operations_tab:
    st.subheader("📊 Operational Analytics")

    col1, col2 = st.columns(2)

    with col1:
        top_n_chart(filtered, "Task", "🏆 Top 10 Most Requested Tasks")

    with col2:
        top_n_chart(filtered, "Location", "📍 Top 10 Locations / Rooms by Task Volume")

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
            fig = px.bar(
                req_dept, x="Req Department", y="Tasks",
                text="Tasks", title="Requests by Requesting Department"
            )
            st.plotly_chart(fig, use_container_width=True)

    with col4:
        hour_data = filtered.dropna(subset=["Request Hour"])

        if not hour_data.empty:
            hourly = (
                hour_data.groupby("Request Hour")
                .size()
                .reset_index(name="Tasks")
            )

            fig = px.bar(
                hourly, x="Request Hour", y="Tasks",
                title="🕒 Peak Request Hour"
            )
            fig.update_xaxes(dtick=1)
            st.plotly_chart(fig, use_container_width=True)


# =========================================================
# DEFECT ANALYTICS
# =========================================================

with defects_tab:
    st.subheader("🔧 Defect Analytics")

    # Defects are automatically based on Engineering team
    defect_df = filtered[
        filtered["Team"]
        .fillna("")
        .str.contains("engineering", case=False, na=False)
    ].copy()

    if defect_df.empty:
        st.warning("No Engineering team data is available with the current filters.")
        st.stop()

    d1, d2, d3 = st.columns(3)

    defect_total = len(defect_df)
    defect_done = int((defect_df["Status"] == "Done").sum())
    defect_avg = defect_df.loc[
        defect_df["Status"] == "Done", "Resolution Hours"
    ].mean()

    d1.metric("🔧 Total Defect Tasks", f"{defect_total:,}")
    d2.metric("✅ Resolved Defects", f"{defect_done:,}")
    d3.metric("⏱️ Avg Resolution", format_duration(defect_avg))

    st.divider()

    # ---------------- MAIN TOP 10 ----------------
    col1, col2 = st.columns(2)

    with col1:
        top_n_chart(defect_df, "Task", "🔧 Top 10 Defects")

    with col2:
        top_n_chart(
            defect_df,
            "Location",
            "🚨 Top 10 Problematic Rooms / Locations"
        )

    st.divider()

    # =====================================================
    # DEFECT-CENTRIC BREAKDOWN
    # =====================================================

    st.subheader("🔍 Defect Breakdown")

    top_defects = top_n_counts(defect_df, "Task", 10)["Task"].tolist()

    if top_defects:
        selected_defect = st.selectbox(
            "Select a defect to analyze",
            top_defects,
            key="selected_defect"
        )

        selected_defect_df = defect_df[
            defect_df["Task"].astype(str) == str(selected_defect)
        ].copy()

        b1, b2 = st.columns(2)

        with b1:
            recurring_locations = (
                selected_defect_df.dropna(subset=["Location"])
                .assign(Location=lambda x: x["Location"].astype(str))
                .groupby("Location")
                .size()
                .reset_index(name="Cases")
                .sort_values("Cases", ascending=False)
                .head(10)
                .sort_values("Cases", ascending=True)
            )

            if recurring_locations.empty:
                st.info("No location data available for this defect.")
            else:
                fig = px.bar(
                    recurring_locations,
                    x="Cases",
                    y="Location",
                    orientation="h",
                    text="Cases",
                    title=f"📍 Top Recurring Locations — {selected_defect}"
                )
                fig.update_yaxes(type="category")
                fig.update_layout(
                    showlegend=False,
                    yaxis_title="",
                    xaxis_title="Cases"
                )
                st.plotly_chart(fig, use_container_width=True)

        with b2:
            trend = selected_defect_df.dropna(subset=["To Do"]).copy()

            if trend.empty:
                st.info("No date/time data available for this defect.")
            else:
                trend["Month"] = trend["To Do"].dt.to_period("M").astype(str)

                trend_data = (
                    trend.groupby("Month")
                    .size()
                    .reset_index(name="Cases")
                )

                # Chronological sort using real timestamps
                trend_data["SortDate"] = pd.to_datetime(
                    trend_data["Month"] + "-01",
                    errors="coerce"
                )
                trend_data = trend_data.sort_values("SortDate")

                fig = px.line(
                    trend_data,
                    x="Month",
                    y="Cases",
                    markers=True,
                    title=f"📈 Trend Over Time — {selected_defect}"
                )
                fig.update_layout(
                    xaxis_title="Month",
                    yaxis_title="Cases"
                )
                st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # =====================================================
    # LOCATION-CENTRIC BREAKDOWN
    # =====================================================

    st.subheader("🏨 Problematic Room / Location Breakdown")

    top_locations = top_n_counts(defect_df, "Location", 10)["Location"].astype(str).tolist()

    if top_locations:
        selected_location = st.selectbox(
            "Select a problematic room / location",
            top_locations,
            key="selected_problem_location"
        )

        selected_location_df = defect_df[
            defect_df["Location"].astype(str) == str(selected_location)
        ].copy()

        issue_types = (
            selected_location_df.dropna(subset=["Task"])
            .groupby("Task")
            .size()
            .reset_index(name="Cases")
            .sort_values("Cases", ascending=False)
            .head(10)
            .sort_values("Cases", ascending=True)
        )

        if issue_types.empty:
            st.info("No defect/issue type data available for this location.")
        else:
            fig = px.bar(
                issue_types,
                x="Cases",
                y="Task",
                orientation="h",
                text="Cases",
                title=f"🔧 Types of Defects / Issues — Location {selected_location}"
            )
            fig.update_layout(
                showlegend=False,
                yaxis_title="",
                xaxis_title="Cases"
            )
            fig.update_yaxes(type="category")
            st.plotly_chart(fig, use_container_width=True)


# =========================================================
# TEAM PERFORMANCE
# =========================================================

with team_tab:
    st.subheader("👥 Team Performance")

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
        team_summary["Completed"] / team_summary["Total_Tasks"] * 100
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


# =========================================================
# TASK EXPLORER
# =========================================================

with explorer_tab:
    st.subheader("📋 Task Explorer")

    e1, e2, e3 = st.columns(3)

    with e1:
        options = sorted(filtered["Team"].dropna().unique().tolist())
        selected_explorer_teams = st.multiselect(
            "🏢 Team", options, default=options, key="explorer_team"
        )

    with e2:
        options = sorted(filtered["Location"].dropna().astype(str).unique().tolist())
        selected_locations = st.multiselect(
            "📍 Location", options, default=options, key="explorer_location"
        )

    with e3:
        options = sorted(filtered["Assignee"].dropna().unique().tolist())
        selected_assignees = st.multiselect(
            "👤 Assignee", options, default=options, key="explorer_assignee"
        )

    e4, e5, e6 = st.columns(3)

    with e4:
        options = sorted(filtered["Status"].dropna().unique().tolist())
        selected_explorer_statuses = st.multiselect(
            "📌 Status", options, default=options, key="explorer_status"
        )

    with e5:
        options = sorted(filtered["Priority"].dropna().unique().tolist())
        selected_explorer_priorities = st.multiselect(
            "🔥 Priority", options, default=options, key="explorer_priority"
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
        display = display[display["Location"].astype(str).isin(selected_locations)]

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
                display[col].fillna("").astype(str)
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


st.divider()
st.caption(
    f"🏨 StayPlease Operational Intelligence | "
    f"📊 {len(filtered):,} filtered tasks | "
    f"🏢 {filtered['Team'].nunique()} team(s) | "
    f"📁 {filtered['Source File'].nunique()} source file(s)"
)
