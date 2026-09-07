import streamlit as st
import pandas as pd
import plotly.express as px

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="StayPlease Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 StayPlease Task Dashboard")
st.caption("Operational Task Performance Dashboard")


# =========================================================
# EXPECTED EXCEL COLUMNS
# =========================================================

EXPECTED_COLUMNS = [
    "Location",
    "Task",
    "Quantity",
    "Req Department",
    "Requestor",
    "Department",
    "Assignee",
    "Status",
    "To Do",
    "Doing",
    "Pause Time",
    "Resume Time",
    "Done",
    "Time Usage",
    "Pause Duration",
    "Pause Reason",
    "Priority",
    "Message",
    "Extra"
]


# =========================================================
# LOAD SINGLE EXCEL FILE
# =========================================================

def load_excel(file):

    xls = pd.ExcelFile(file)

    frames = []

    for sheet in xls.sheet_names:

        # Skip deleted task sheet
        if sheet.strip().lower() == "deleted tasks":
            continue

        try:

            # Original report structure uses row 3 as header
            raw = pd.read_excel(
                file,
                sheet_name=sheet,
                header=2
            )

            # Remove completely empty rows
            raw = raw.dropna(how="all")

            if raw.empty:
                continue

            # Remove repeated headers
            if len(raw.columns) > 0:
                raw = raw[
                    raw.iloc[:, 0]
                    .astype(str)
                    .str.strip()
                    .ne("Location")
                ]

            # Rename columns according to expected structure
            raw.columns = EXPECTED_COLUMNS[:len(raw.columns)]

            # Add missing columns
            for col in EXPECTED_COLUMNS:
                if col not in raw.columns:
                    raw[col] = pd.NA

            # Keep consistent columns
            raw = raw[EXPECTED_COLUMNS]

            # Add team information from sheet
            team_name = sheet

            # Remove numeric prefix like:
            # 1_Housekeeping -> Housekeeping
            if "_" in team_name:
                first_part = team_name.split("_")[0]

                if first_part.isdigit():
                    team_name = "_".join(team_name.split("_")[1:])

            raw["Team"] = team_name

            frames.append(raw)

        except Exception as e:

            st.warning(
                f"⚠️ Could not read sheet '{sheet}' "
                f"from {file.name}: {e}"
            )

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


# =========================================================
# LOAD MULTIPLE EXCEL FILES
# =========================================================

@st.cache_data
def load_multiple_excels(uploaded_files):

    all_frames = []

    for file in uploaded_files:

        df_file = load_excel(file)

        if df_file.empty:
            continue

        # Add source file information
        df_file["Source File"] = file.name

        all_frames.append(df_file)

    if not all_frames:
        return pd.DataFrame()

    df = pd.concat(
        all_frames,
        ignore_index=True
    )

    return df


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.header("📤 Data Upload")

    uploaded_files = st.file_uploader(
        "Upload Task Report Excel Files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="You can upload multiple Task Report Excel files at once"
    )

    st.divider()

    st.header("🎛️ Filters")


# =========================================================
# CHECK UPLOAD
# =========================================================

if not uploaded_files:

    st.info(
        "👆 Upload one or multiple Task Report Excel files "
        "from the sidebar to start."
    )

    st.stop()


# =========================================================
# LOAD DATA
# =========================================================

with st.spinner("📂 Reading and consolidating Excel files..."):

    df = load_multiple_excels(uploaded_files)


if df.empty:

    st.error(
        "❌ No valid task data was found in the uploaded Excel files."
    )

    st.stop()


# =========================================================
# DATA CLEANING
# =========================================================

TEXT_COLUMNS = [
    "Location",
    "Task",
    "Req Department",
    "Requestor",
    "Department",
    "Assignee",
    "Status",
    "Priority",
    "Pause Reason",
    "Message",
    "Team",
    "Source File"
]

for col in TEXT_COLUMNS:

    df[col] = (
        df[col]
        .astype("string")
        .str.strip()
    )


# Normalize status

STATUS_MAP = {

    "ToDo": "To Do",
    "Todo": "To Do",
    "To Do": "To Do",

    "Doing": "Doing",

    "Pause": "Paused",
    "Paused": "Paused",

    "Done": "Done"

}

df["Status"] = df["Status"].replace(STATUS_MAP)


# Convert dates

DATE_COLUMNS = [

    "To Do",
    "Doing",
    "Pause Time",
    "Resume Time",
    "Done"

]

for col in DATE_COLUMNS:

    df[col] = pd.to_datetime(
        df[col],
        errors="coerce"
    )


# Numeric quantity

df["Quantity"] = pd.to_numeric(
    df["Quantity"],
    errors="coerce"
)


# =========================================================
# REPORT DATE
# =========================================================

df["Report Date"] = (
    df["Done"]
    .fillna(df["To Do"])
)

df["Report Date"] = pd.to_datetime(
    df["Report Date"],
    errors="coerce"
)


# =========================================================
# DURATION CALCULATION
# =========================================================

df["Duration Hours"] = (

    df["Done"]
    -
    df["To Do"]

).dt.total_seconds() / 3600


df.loc[
    df["Duration Hours"] < 0,
    "Duration Hours"
] = pd.NA


# =========================================================
# FILTERS
# =========================================================

with st.sidebar:

    # SOURCE FILE FILTER

    source_files = sorted(
        df["Source File"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_files = st.multiselect(
        "📁 Source File",
        source_files,
        default=source_files
    )


    # TEAM FILTER

    teams = sorted(
        df["Team"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_teams = st.multiselect(
        "🏢 Team / Department",
        teams,
        default=teams
    )


    # STATUS FILTER

    statuses = sorted(
        df["Status"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_statuses = st.multiselect(
        "📌 Status",
        statuses,
        default=statuses
    )


    # PRIORITY FILTER

    priorities = sorted(
        df["Priority"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_priorities = st.multiselect(
        "🔥 Priority",
        priorities,
        default=priorities
    )


    # DATE FILTER

    valid_dates = df[
        "Report Date"
    ].dropna()


    date_range = None


    if not valid_dates.empty:

        min_date = valid_dates.min().date()

        max_date = valid_dates.max().date()


        date_range = st.date_input(

            "📅 Report Date Range",

            value=(
                min_date,
                max_date
            )

        )


# =========================================================
# APPLY FILTERS
# =========================================================

filtered = df.copy()


# Source file

if selected_files:

    filtered = filtered[
        filtered["Source File"].isin(
            selected_files
        )
    ]


# Team

if selected_teams:

    filtered = filtered[
        filtered["Team"].isin(
            selected_teams
        )
    ]


# Status

if selected_statuses:

    filtered = filtered[
        filtered["Status"].isin(
            selected_statuses
        )
    ]


# Priority

if selected_priorities:

    filtered = filtered[

        filtered["Priority"].isna()

        |

        filtered["Priority"].isin(
            selected_priorities
        )

    ]


# Date range

if date_range and len(date_range) == 2:

    start_date = pd.Timestamp(
        date_range[0]
    )


    end_date = (

        pd.Timestamp(
            date_range[1]
        )

        +

        pd.Timedelta(days=1)

    )


    filtered = filtered[

        filtered["Report Date"].isna()

        |

        (

            (filtered["Report Date"] >= start_date)

            &

            (filtered["Report Date"] < end_date)

        )

    ]


# =========================================================
# KPI CALCULATIONS
# =========================================================

total_tasks = len(filtered)


completed_tasks = int(

    (
        filtered["Status"] == "Done"
    ).sum()

)


todo_tasks = int(

    (
        filtered["Status"] == "To Do"
    ).sum()

)


doing_tasks = int(

    (
        filtered["Status"] == "Doing"
    ).sum()

)


paused_tasks = int(

    (
        filtered["Status"] == "Paused"
    ).sum()

)


completion_rate = (

    completed_tasks
    /
    total_tasks
    *
    100

    if total_tasks > 0

    else 0

)


# =========================================================
# KPI DISPLAY
# =========================================================

st.subheader("📌 Dashboard Summary")


k1, k2, k3, k4, k5, k6 = st.columns(6)


k1.metric(
    "📋 Total Tasks",
    f"{total_tasks:,}"
)


k2.metric(
    "✅ Completed",
    f"{completed_tasks:,}"
)


k3.metric(
    "📝 To Do",
    f"{todo_tasks:,}"
)


k4.metric(
    "⚙️ Doing",
    f"{doing_tasks:,}"
)


k5.metric(
    "⏸️ Paused",
    f"{paused_tasks:,}"
)


k6.metric(
    "🎯 Completion Rate",
    f"{completion_rate:.1f}%"
)


st.divider()


# =========================================================
# TABS
# =========================================================

tab1, tab2, tab3 = st.tabs([

    "📈 Overview",

    "👥 Team Performance",

    "📋 Task Explorer"

])


# =========================================================
# TAB 1 — OVERVIEW
# =========================================================

with tab1:


    col1, col2 = st.columns(2)


    # STATUS PIE CHART

    status_counts = (

        filtered["Status"]

        .fillna("Unknown")

        .value_counts()

        .reset_index()

    )


    status_counts.columns = [

        "Status",

        "Tasks"

    ]


    fig_status = px.pie(

        status_counts,

        names="Status",

        values="Tasks",

        hole=0.55,

        title="Task Status Distribution"

    )


    col1.plotly_chart(

        fig_status,

        use_container_width=True

    )


    # TEAM CHART

    team_counts = (

        filtered

        .groupby("Team")

        .size()

        .reset_index(name="Tasks")

        .sort_values(
            "Tasks",
            ascending=True
        )

    )


    fig_team = px.bar(

        team_counts,

        x="Tasks",

        y="Team",

        orientation="h",

        text="Tasks",

        title="Tasks by Team"

    )


    col2.plotly_chart(

        fig_team,

        use_container_width=True

    )


    # DAILY ACTIVITY

    trend = filtered.dropna(

        subset=["Report Date"]

    ).copy()


    if not trend.empty:


        trend["Date"] = (

            trend["Report Date"]

            .dt.date

        )


        daily = (

            trend

            .groupby([
                "Date",
                "Status"
            ])

            .size()

            .reset_index(
                name="Tasks"
            )

        )


        fig_trend = px.bar(

            daily,

            x="Date",

            y="Tasks",

            color="Status",

            barmode="stack",

            title="Daily Task Activity"

        )


        st.plotly_chart(

            fig_trend,

            use_container_width=True

        )


    # URGENT TASKS

    st.subheader("🚨 Urgent Tasks")


    urgent = filtered[

        filtered["Priority"]

        .fillna("")

        .str.contains(

            "Urgent",

            case=False,

            na=False

        )

    ]


    if urgent.empty:


        st.success(
            "No urgent tasks in the current selection."
        )


    else:


        st.dataframe(

            urgent[

                [

                    "Source File",

                    "Team",

                    "Location",

                    "Task",

                    "Assignee",

                    "Status",

                    "Priority",

                    "To Do",

                    "Done"

                ]

            ]

            .sort_values(

                "To Do",

                ascending=False

            ),

            use_container_width=True,

            hide_index=True

        )


# =========================================================
# TAB 2 — TEAM PERFORMANCE
# =========================================================

with tab2:


    team_summary = (

        filtered

        .groupby("Team")

        .agg(

            Total_Tasks=(

                "Task",

                "size"

            ),

            Completed=(

                "Status",

                lambda x: (
                    x == "Done"
                ).sum()

            ),

            To_Do=(

                "Status",

                lambda x: (
                    x == "To Do"
                ).sum()

            ),

            Doing=(

                "Status",

                lambda x: (
                    x == "Doing"
                ).sum()

            ),

            Paused=(

                "Status",

                lambda x: (
                    x == "Paused"
                ).sum()

            ),

            Avg_Duration_Hours=(

                "Duration Hours",

                "mean"

            )

        )

        .reset_index()

    )


    team_summary[

        "Completion Rate %"

    ] = (

        team_summary["Completed"]

        /

        team_summary["Total_Tasks"]

        *

        100

    ).round(1)


    team_summary[

        "Avg_Duration_Hours"

    ] = (

        team_summary[

            "Avg_Duration_Hours"

        ]

        .round(2)

    )


    st.dataframe(

        team_summary,

        use_container_width=True,

        hide_index=True

    )


    # ASSIGNEE PERFORMANCE

    assignee = filtered.dropna(

        subset=["Assignee"]

    )


    if not assignee.empty:


        assignee_summary = (

            assignee

            .groupby([

                "Team",

                "Assignee"

            ])

            .size()

            .reset_index(

                name="Tasks"

            )

            .sort_values(

                "Tasks",

                ascending=False

            )

            .head(20)

        )


        fig_assignee = px.bar(

            assignee_summary,

            x="Tasks",

            y="Assignee",

            color="Team",

            orientation="h",

            title="Top 20 Assignees by Task Volume"

        )


        st.plotly_chart(

            fig_assignee,

            use_container_width=True

        )


# =========================================================
# TAB 3 — TASK EXPLORER
# =========================================================

with tab3:


    search = st.text_input(

        "🔎 Search task, location, requestor, assignee, or message"

    )


    display = filtered.copy()


    if search:


        search_columns = [

            "Task",

            "Location",

            "Requestor",

            "Assignee",

            "Message",

            "Req Department"

        ]


        mask = pd.Series(

            False,

            index=display.index

        )


        for col in search_columns:


            mask |= (

                display[col]

                .fillna("")

                .astype(str)

                .str.contains(

                    search,

                    case=False,

                    na=False

                )

            )


        display = display[mask]


    display_columns = [

        "Source File",

        "Team",

        "Location",

        "Task",

        "Quantity",

        "Req Department",

        "Requestor",

        "Department",

        "Assignee",

        "Status",

        "Priority",

        "To Do",

        "Doing",

        "Done",

        "Duration Hours",

        "Pause Reason",

        "Message"

    ]


    st.dataframe(

        display[display_columns]

        .sort_values(

            "To Do",

            ascending=False

        ),

        use_container_width=True,

        hide_index=True,

        height=600

    )


    # DOWNLOAD FILTERED DATA

    csv = (

        display[display_columns]

        .to_csv(

            index=False

        )

        .encode("utf-8")

    )


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

    f"📊 Showing {len(filtered):,} tasks | "
    f"📁 {filtered['Source File'].nunique()} Excel file(s) | "
    f"🏢 {filtered['Team'].nunique()} team(s)"

)
