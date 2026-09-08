import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image, KeepTogether
)

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


def read_excel_compatible(uploaded_file, **kwargs):
    """Read modern .xlsx and legacy Excel 97-2003 .xls files.

    Legacy .xls files require xlrd>=2.0.1 in requirements.txt.
    """
    name = str(getattr(uploaded_file, "name", "")).lower()
    engine = "xlrd" if name.endswith(".xls") else None
    try:
        return pd.read_excel(uploaded_file, engine=engine, **kwargs)
    except Exception:
        # Fallback lets pandas detect the engine for files whose extension/content differs.
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, **kwargs)


def read_single_excel(uploaded_file):
    frames = []

    try:
        xls = pd.ExcelFile(uploaded_file, engine="xlrd" if str(getattr(uploaded_file, "name", "")).lower().endswith(".xls") else None)
    except Exception:
        return pd.DataFrame()

    for sheet in xls.sheet_names:
        if str(sheet).strip().lower() == "deleted tasks":
            continue

        try:
            raw = read_excel_compatible(uploaded_file, sheet_name=sheet, header=2)
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

    # Add Property / Area classification for all dashboard pages.
    df = add_property_area_mapping(df)

    return df


def classify_stayplease_location(location):
    """Classify locations for Property and Area filtering."""
    if pd.isna(location):
        return pd.Series([pd.NA, "Unmapped", pd.NA])
    loc = str(location).strip()
    if not loc or loc.lower() in {"nan", "none", "<na>"}:
        return pd.Series([pd.NA, "Unmapped", pd.NA])

    import re
    match = re.fullmatch(r"(\d{2})(\d{2})", loc)
    if match:
        floor = int(match.group(1))
        if 73 <= floor <= 82:
            return pd.Series(["PRSJKT", "Guest Rooms", f"Floor {floor}"])
        if 83 <= floor <= 89:
            return pd.Series(["PPJKT", "Guest Rooms", f"Floor {floor}"])

    upper = loc.upper()
    if upper.startswith("PR"):
        if "RESIDENCE LOUNGE" in upper:
            return pd.Series(["PRSJKT", "Public Areas", "Residence Lounge"])
        if "LOBBY" in upper or "RECEPTION" in upper:
            return pd.Series(["PRSJKT", "Public Areas", "Lobby"])
        if "GYM" in upper or "CHANGING ROOM" in upper:
            return pd.Series(["PRSJKT", "Public Areas", "Gym"])
        if "POOL" in upper:
            return pd.Series(["PRSJKT", "Public Areas", "Pool"])
        return pd.Series(["PRSJKT", "Other Areas", "Other PRSJKT Area"])

    if upper.startswith("PP"):
        if "LOBBY" in upper or "RECEPTION" in upper:
            return pd.Series(["PPJKT", "Public Areas", "Lobby"])
        outlet_keywords = ["KEYAKI", "TEPANYAKI", "EDEN BAR", "RESTAURANT", "OUTLET", "LEVEL 90", "PDR"]
        if any(keyword in upper for keyword in outlet_keywords):
            return pd.Series(["PPJKT", "Public Areas", "Restaurant & Outlets"])
        return pd.Series(["PPJKT", "Other Areas", "Other PPJKT Area"])

    return pd.Series([pd.NA, "Unmapped", pd.NA])


def add_property_area_mapping(df):
    mapped = df["Location"].apply(classify_stayplease_location)
    mapped.columns = ["Property", "Area Type", "Specific Area"]
    return pd.concat([df, mapped], axis=1)



# =========================================================
# WORK ORDER + INCIDENT DATA SOURCES
# =========================================================

def read_work_order_excel(uploaded_file):
    try:
        raw = read_excel_compatible(uploaded_file, header=3)
    except Exception:
        return pd.DataFrame()
    raw = raw.dropna(how="all")
    if raw.empty:
        return pd.DataFrame()
    cols = ["Location", "Work Order ID", "Overdue", "Date Created", "Status", "Work Order Title", "Assigned To", "Created By", "Last Update", "Message"]
    raw = raw.iloc[:, :len(cols)].copy()
    raw.columns = cols[:len(raw.columns)]
    for c in cols:
        if c not in raw.columns:
            raw[c] = pd.NA
    raw = raw[cols]
    raw["Location"] = raw["Location"].astype("string").str.strip()
    raw["Work Order Title"] = raw["Work Order Title"].astype("string").str.strip()
    raw["Status"] = raw["Status"].astype("string").str.strip()
    raw["Date Created"] = pd.to_datetime(raw["Date Created"], errors="coerce")
    raw["Source File"] = uploaded_file.name
    raw = raw[raw["Work Order Title"].notna() & raw["Work Order Title"].ne("")].copy()
    raw = add_property_area_mapping(raw)
    return raw


def load_work_orders(files):
    frames = [read_work_order_excel(f) for f in files]
    frames = [x for x in frames if not x.empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def read_incident_excel(uploaded_file):
    """Read StayPlease Incident List Report.

    The report uses a two-row visual header followed by one row per incident.
    Timeline follow-up rows do not contain a Log Number, so they are excluded.
    """
    try:
        raw = read_excel_compatible(uploaded_file, sheet_name="Incident", header=None)
    except Exception:
        try:
            raw = read_excel_compatible(uploaded_file, header=None)
        except Exception:
            return pd.DataFrame()

    if raw.empty or raw.shape[1] < 10:
        return pd.DataFrame()

    # Actual incident records are the rows with a numeric Log Number in column 0.
    log_num = pd.to_numeric(raw.iloc[:, 0], errors="coerce")
    records = raw.loc[log_num.notna()].copy()
    if records.empty:
        return pd.DataFrame()

    # Keep only real incident records and map the report's fixed column layout.
    def col(idx):
        return records.iloc[:, idx] if idx < records.shape[1] else pd.Series(pd.NA, index=records.index)

    df = pd.DataFrame({
        "Log No": col(0),
        "Location": col(1),
        "Incident Name": col(2),
        "Status": col(3),
        "Created By": col(4),
        "Creation Time": col(5),
        "Staff Tag": col(6),
        "Guest Temp": col(7),
        "Department": col(8),
        "Deadline": col(9),
        "Guest Name": col(10),
        "VIP Level": col(11),
        "Arrival": col(12),
        "Departure": col(13),
        "More Information": col(14),
        "Guest Feedback": col(25),
        "Compensation": col(26),
        "Other Compensation": col(27),
        "Cost": col(28),
        "Issued By": col(29),
    })

    text_cols = [
        "Log No", "Location", "Incident Name", "Status", "Created By",
        "Staff Tag", "Guest Temp", "Department", "VIP Level", "Guest Name",
        "More Information", "Guest Feedback", "Compensation",
        "Other Compensation", "Issued By"
    ]
    for c in text_cols:
        df[c] = df[c].astype("string").str.strip()

    for c in ["Creation Time", "Deadline", "Arrival", "Departure"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    df["Cost"] = pd.to_numeric(df["Cost"], errors="coerce").fillna(0)

    status_norm = df["Status"].astype("string").str.strip().str.lower()
    df["Status"] = status_norm.replace({"close": "Closed", "closed": "Closed", "open": "Open"}).fillna(df["Status"])

    df["Source File"] = getattr(uploaded_file, "name", "Incident List Report")
    df = df[df["Incident Name"].notna() & df["Incident Name"].ne("")].copy()
    df = add_property_area_mapping(df)
    return df

def load_incidents(files):
    frames = [read_incident_excel(f) for f in files]
    frames = [x for x in frames if not x.empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def explode_incident_categories(data):
    if data.empty:
        return pd.DataFrame(columns=["Category", "Cases"])
    x = data[["Incident Name"]].dropna().copy()
    x["Category"] = x["Incident Name"].astype(str).str.split(",")
    x = x.explode("Category")
    x["Category"] = x["Category"].astype(str).str.replace(r"^\s*!\s*", "", regex=True).str.strip()
    x = x[x["Category"].ne("")]
    return x.groupby("Category").size().reset_index(name="Cases").sort_values("Cases", ascending=False)


def filter_auxiliary_data(data, date_col, properties, area_types, specific_areas, date_range):
    if data.empty:
        return data
    out = data[data["Area Type"] != "Unmapped"].copy()
    if properties:
        out = out[out["Property"].isin(properties)]
    if area_types:
        out = out[out["Area Type"].isin(area_types)]
    if specific_areas:
        out = out[out["Specific Area"].astype(str).isin(specific_areas)]
    if date_range and len(date_range) == 2 and date_col in out.columns:
        start = pd.Timestamp(date_range[0])
        end = pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
        out = out[out[date_col].isna() | ((out[date_col] >= start) & (out[date_col] < end))]
    return out

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


def add_count_hover(fig, label_col, count_col="Tasks", total=None, extra_cols=None):
    """Apply a consistent, lightweight Plotly hover template to ranking charts."""
    if extra_cols is None:
        extra_cols = []

    # The chart data already exists in the figure, so customdata only adds
    # lightweight tooltip metadata and does not require additional queries.
    custom_fields = [label_col, count_col] + extra_cols
    customdata = []

    for trace in fig.data:
        # Plotly Express keeps source values in trace customdata when supplied.
        # This helper is intentionally conservative and only sets the template.
        trace.hovertemplate = (
            f"<b>%{{y}}</b><br>"
            f"{count_col}: %{{x}}"
            "<extra></extra>"
        )

    return fig


def top_n_chart(data, group_col, title, n=10):
    chart_data = top_n_counts(data, group_col, n)

    if chart_data.empty:
        st.info("No data available.")
        return

    # Reverse order for readable horizontal ranking
    plot_data = chart_data.sort_values("Tasks", ascending=True)

    total_cases = int(chart_data["Tasks"].sum())
    plot_data["Share %"] = (plot_data["Tasks"] / total_cases * 100).round(1) if total_cases else 0

    fig = px.bar(
        plot_data,
        x="Tasks",
        y=group_col,
        orientation="h",
        text="Tasks",
        custom_data=[group_col, "Tasks", "Share %"],
        title=title
    )
    fig.update_traces(
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "Cases / Tasks: %{customdata[1]:,}<br>"
            "Share of Top 10: %{customdata[2]:.1f}%"
            "<extra></extra>"
        )
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
# PDF ANALYTIC REPORT
# =========================================================

def _pdf_chart_bar(data, label_col, value_col, title, xlabel, top_n=10):
    """Create a lightweight horizontal bar chart as an in-memory PNG."""
    if data is None or data.empty:
        return None

    chart = data.head(top_n).copy().sort_values(value_col, ascending=True)
    fig, ax = plt.subplots(figsize=(8.2, 4.6))
    ax.barh(chart[label_col].astype(str), chart[value_col])
    ax.set_title(title, loc="left", fontweight="bold")
    ax.set_xlabel(xlabel)
    ax.grid(axis="x", alpha=0.25)
    for i, value in enumerate(chart[value_col]):
        ax.text(value, i, f" {int(value):,}", va="center", fontsize=8)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _pdf_chart_line(data, x_col, y_col, title, ylabel):
    if data is None or data.empty:
        return None
    fig, ax = plt.subplots(figsize=(8.2, 4.2))
    ax.plot(data[x_col].astype(str), data[y_col], marker="o")
    ax.set_title(title, loc="left", fontweight="bold")
    ax.set_ylabel(ylabel)
    ax.grid(axis="y", alpha=0.25)
    ax.tick_params(axis="x", rotation=35)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def build_pdf_report(data, filter_context, work_orders=None, incidents=None):
    """Build a management-ready PDF report from the currently filtered data."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
        title="StayPlease Operational Intelligence Report",
        author="StayPlease Operational Intelligence"
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ReportTitle", parent=styles["Title"], fontSize=22,
        leading=27, textColor=colors.HexColor("#1F3A5F"), spaceAfter=8
    ))
    styles.add(ParagraphStyle(
        name="ReportSubtitle", parent=styles["Normal"], fontSize=10,
        leading=14, textColor=colors.HexColor("#666666"), spaceAfter=12
    ))
    styles.add(ParagraphStyle(
        name="SectionHeader", parent=styles["Heading2"], fontSize=15,
        leading=19, textColor=colors.HexColor("#1F3A5F"), spaceBefore=8, spaceAfter=8
    ))
    styles.add(ParagraphStyle(
        name="Insight", parent=styles["BodyText"], fontSize=10,
        leading=14, leftIndent=10, spaceAfter=5
    ))

    story = []
    generated = datetime.now().strftime("%d %b %Y %H:%M")
    total = len(data)
    completed = int((data["Status"] == "Done").sum())
    open_tasks = total - completed
    completion_rate = completed / total * 100 if total else 0
    urgent_open = int((
        data["Priority"].fillna("").str.contains("urgent", case=False, na=False)
        & (data["Status"] != "Done")
    ).sum())
    avg_resolution = data.loc[
        (data["Status"] == "Done") & data["Resolution Hours"].notna(),
        "Resolution Hours"
    ].mean()

    # COVER + EXECUTIVE SUMMARY
    story.append(Paragraph("StayPlease Operational Intelligence", styles["ReportTitle"]))
    story.append(Paragraph("Operational Analytics Report", styles["Heading2"]))
    story.append(Paragraph(
        f"Generated: {generated}<br/>"
        f"Scope: {filter_context}<br/>"
        f"Tasks included: {total:,}",
        styles["ReportSubtitle"]
    ))
    story.append(Spacer(1, 12))

    kpi_data = [
        ["Total Tasks", "Completed", "Open Tasks"],
        [f"{total:,}", f"{completed:,}", f"{open_tasks:,}"],
        ["Completion Rate", "Avg Resolution", "Urgent Open"],
        [f"{completion_rate:.1f}%", format_duration(avg_resolution), f"{urgent_open:,}"],
    ]
    kpi_table = Table(kpi_data, colWidths=[2.35*inch]*3)
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#DCE6F1")),
        ("BACKGROUND", (0,2), (-1,2), colors.HexColor("#DCE6F1")),
        ("TEXTCOLOR", (0,0), (-1,-1), colors.HexColor("#1F3A5F")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,2), (-1,2), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("FONTSIZE", (0,1), (-1,1), 18),
        ("FONTSIZE", (0,3), (-1,3), 18),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#B7C9DD")),
        ("TOPPADDING", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,-1), 9),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 18))

    # KEY FINDINGS
    story.append(Paragraph("Key Findings & Insights", styles["SectionHeader"]))
    top_tasks = top_n_counts(data, "Task", 10)
    top_locations = top_n_counts(data, "Location", 10)
    team_counts = top_n_counts(data, "Team", 10)
    insights = []
    if not top_tasks.empty:
        row = top_tasks.iloc[0]
        insights.append(f"<b>Highest-volume request:</b> {row['Task']} recorded {int(row['Tasks']):,} task(s) within the selected scope.")
    if not top_locations.empty:
        row = top_locations.iloc[0]
        insights.append(f"<b>Primary operational hotspot:</b> Location {row['Location']} recorded {int(row['Tasks']):,} task(s).")
    if not team_counts.empty:
        row = team_counts.iloc[0]
        insights.append(f"<b>Highest task volume by team:</b> {row['Team']} handled {int(row['Tasks']):,} task(s).")
    insights.append(f"<b>Completion performance:</b> {completion_rate:.1f}% of included tasks are completed, with {urgent_open:,} urgent task(s) still open.")
    if pd.notna(avg_resolution):
        insights.append(f"<b>Resolution performance:</b> Average completed-task resolution time is {format_duration(avg_resolution)}.")
    for item in insights:
        story.append(Paragraph("• " + item, styles["Insight"]))

    story.append(PageBreak())

    # OPERATIONAL ANALYTICS
    story.append(Paragraph("Operational Analytics", styles["SectionHeader"]))
    if not top_tasks.empty:
        img = _pdf_chart_bar(top_tasks, "Task", "Tasks", "Top 10 Most Requested Tasks", "Tasks")
        if img:
            story.append(Image(img, width=7.0*inch, height=3.8*inch))
    if not top_locations.empty:
        img = _pdf_chart_bar(top_locations, "Location", "Tasks", "Top 10 Locations / Rooms by Task Volume", "Tasks")
        if img:
            story.append(Image(img, width=7.0*inch, height=3.8*inch))



    story.append(PageBreak())

    # DEFECT ANALYTICS
    story.append(Paragraph("Defect Analytics", styles["SectionHeader"]))
    defect_data = data[data["Team"].fillna("").str.contains("engineering", case=False, na=False)].copy()
    if defect_data.empty:
        story.append(Paragraph("No Engineering / defect data is available for the selected scope.", styles["Normal"]))
    else:
        defect_total = len(defect_data)
        defect_done = int((defect_data["Status"] == "Done").sum())
        defect_avg = defect_data.loc[(defect_data["Status"] == "Done"), "Resolution Hours"].mean()
        defect_kpi = Table([
            ["Total Defect Tasks", "Resolved Defects", "Avg Resolution"],
            [f"{defect_total:,}", f"{defect_done:,}", format_duration(defect_avg)]
        ], colWidths=[2.35*inch]*3)
        defect_kpi.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#FCE4D6")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#E7B99C")),
            ("FONTSIZE", (0,1), (-1,1), 16),
            ("TOPPADDING", (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ]))
        story.append(defect_kpi)
        story.append(Spacer(1, 14))

        top_defects = top_n_counts(defect_data, "Task", 10)
        problematic = top_n_counts(defect_data, "Location", 10)
        if not top_defects.empty:
            img = _pdf_chart_bar(top_defects, "Task", "Tasks", "Top 10 Defects", "Cases")
            if img:
                story.append(Image(img, width=7.0*inch, height=3.8*inch))
        if not problematic.empty:
            img = _pdf_chart_bar(problematic, "Location", "Tasks", "Top 10 Problematic Rooms / Locations", "Cases")
            if img:
                story.append(Image(img, width=7.0*inch, height=3.8*inch))

        trend = defect_data.dropna(subset=["To Do"]).copy()
        if not trend.empty:
            trend["Month"] = trend["To Do"].dt.to_period("M").astype(str)
            trend_data = trend.groupby("Month").size().reset_index(name="Cases")
            trend_data["SortDate"] = pd.to_datetime(trend_data["Month"] + "-01", errors="coerce")
            trend_data = trend_data.sort_values("SortDate")
            story.append(Paragraph("Defect Trend Over Time", styles["SectionHeader"]))
            img = _pdf_chart_line(trend_data, "Month", "Cases", "Defect Trend Over Time", "Cases")
            if img:
                story.append(Image(img, width=7.0*inch, height=3.6*inch))

    story.append(PageBreak())

    # CROSS-SOURCE OPERATIONAL INSIGHTS
    story.append(Paragraph("Cross-Source Operational Insights", styles["SectionHeader"]))
    story.append(Paragraph(
        "This section compares maintenance and incident records at the same mapped location. It highlights operational overlap for review and does not establish causation.",
        styles["Normal"]
    ))
    if work_orders is None or work_orders.empty or incidents is None or incidents.empty:
        story.append(Paragraph("Work Order and Incident reports were not both available for this report scope.", styles["Normal"]))
    else:
        story.append(Paragraph(
            f"Work Orders in scope: {len(work_orders):,} &nbsp;&nbsp; | &nbsp;&nbsp; Incidents in scope: {len(incidents):,}",
            styles["Normal"]
        ))
        wo_counts = work_orders.dropna(subset=["Location"]).assign(Location=lambda x: x["Location"].astype(str)).groupby("Location").size().reset_index(name="Work Orders")
        inc_counts = incidents.dropna(subset=["Location"]).assign(Location=lambda x: x["Location"].astype(str)).groupby("Location").size().reset_index(name="Incidents")
        overlap = wo_counts.merge(inc_counts, on="Location", how="inner").sort_values(["Work Orders","Incidents"], ascending=False)
        if overlap.empty:
            story.append(Paragraph("No mapped location overlap was identified between Work Orders and Incidents in the selected scope.", styles["Normal"]))
        else:
            img = _pdf_chart_bar(overlap, "Location", "Work Orders", "Maintenance Activity at Guest-Impact Locations", "Work Orders", top_n=10)
            if img:
                story.append(Image(img, width=7.0*inch, height=3.8*inch))
            rows = [["Location", "Work Orders", "Incidents"]] + [[str(r["Location"]), str(int(r["Work Orders"])), str(int(r["Incidents"]))] for _,r in overlap.head(10).iterrows()]
            t=Table(rows, repeatRows=1, colWidths=[2.5*inch,2.0*inch,2.0*inch])
            t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#C9D5E3")),("ALIGN",(1,1),(-1,-1),"CENTER")]))
            story.append(t)

    story.append(PageBreak())

    # TEAM PERFORMANCE
    story.append(Paragraph("Team Performance", styles["SectionHeader"]))
    team_summary = (
        data.groupby("Team")
        .agg(
            Total_Tasks=("Task", "size"),
            Completed=("Status", lambda x: (x == "Done").sum()),
            Avg_Resolution_Hours=("Resolution Hours", "mean")
        ).reset_index()
    )
    if not team_summary.empty:
        team_summary["Completion Rate %"] = (team_summary["Completed"] / team_summary["Total_Tasks"] * 100).round(1)
        team_summary = team_summary.sort_values("Total_Tasks", ascending=False)
        table_rows = [["Team", "Tasks", "Completed", "Completion", "Avg Resolution"]]
        for _, row in team_summary.iterrows():
            table_rows.append([
                str(row["Team"]), f"{int(row["Total_Tasks"]):,}", f"{int(row["Completed"]):,}",
                f"{float(row["Completion Rate %"]):.1f}%", format_duration(row["Avg_Resolution_Hours"])
            ])
        t = Table(table_rows, repeatRows=1, colWidths=[1.7*inch, 0.9*inch, 1.0*inch, 1.1*inch, 1.5*inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.4, colors.HexColor("#C9D5E3")),
            ("ALIGN", (1,1), (-1,-1), "CENTER"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F5F8FB")]),
            ("TOPPADDING", (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(t)

    story.append(Spacer(1, 18))
    story.append(Paragraph("Report Notes", styles["SectionHeader"]))
    story.append(Paragraph(
        "This report is generated from the current StayPlease dashboard filters. "
        "Locations that could not be mapped to PRSJKT or PPJKT are excluded from the analysis.",
        styles["Normal"]
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# HEADER + DATA UPLOAD
# =========================================================

st.title("🏨 StayPlease Operational Intelligence")
st.caption("Task • Work Order • Incident • Operational Intelligence Dashboard")

with st.sidebar:
    st.header("📤 Data Sources")
    uploaded_files = st.file_uploader(
        "📋 Task Report (required)", type=["xlsx", "xls"], accept_multiple_files=True,
        help="Upload one or multiple StayPlease Task Report Excel files (.xlsx or legacy .xls)."
    )
    uploaded_work_orders = st.file_uploader(
        "🔧 Work Order Report (optional)", type=["xlsx", "xls"], accept_multiple_files=True
    )
    uploaded_incidents = st.file_uploader(
        "🚨 Incident List Report (optional)", type=["xlsx", "xls"], accept_multiple_files=True
    )

if not uploaded_files:
    st.info("👈 Upload one or multiple StayPlease Task Report Excel files to start. Work Order and Incident reports can then be added for expanded analytics.")
    st.stop()

with st.spinner("Reading StayPlease data sources..."):
    df = load_all_files(uploaded_files)
    work_orders = load_work_orders(uploaded_work_orders) if uploaded_work_orders else pd.DataFrame()
    incidents = load_incidents(uploaded_incidents) if uploaded_incidents else pd.DataFrame()

if df.empty:
    st.error("No valid task data was found. Please check the Task Report format.")
    st.stop()

# =========================================================
# GLOBAL FILTERS
# =========================================================

# Intentionally skip locations that cannot be mapped to PRSJKT or PPJKT.
dashboard_df = df[df["Area Type"] != "Unmapped"].copy()

with st.sidebar:
    st.divider()
    st.header("🏨 Property & Area")

    property_options = ["PRSJKT", "PPJKT"]
    selected_properties = st.multiselect("🏨 Property", property_options, default=property_options)

    area_type_options = ["Guest Rooms", "Public Areas", "Other Areas"]
    selected_area_types = st.multiselect("📍 Area Type", area_type_options, default=area_type_options)

    area_scope = dashboard_df.copy()
    if selected_properties:
        area_scope = area_scope[area_scope["Property"].isin(selected_properties)]
    if selected_area_types:
        area_scope = area_scope[area_scope["Area Type"].isin(selected_area_types)]

    specific_area_options = sorted(area_scope["Specific Area"].dropna().astype(str).unique().tolist())
    selected_specific_areas = st.multiselect(
        "📌 Specific Area", specific_area_options, default=specific_area_options,
        help="Public Areas include Lobby, Residence Lounge, Gym, Pool, and Restaurant & Outlets."
    )

    st.divider()
    st.header("🎛️ Global Filters")

    source_options = sorted(dashboard_df["Source File"].dropna().unique().tolist())
    selected_sources = st.multiselect("📁 Source File", source_options, default=source_options)

    team_options = sorted(dashboard_df["Team"].dropna().unique().tolist())
    selected_teams = st.multiselect("🏢 Team", team_options, default=team_options)

    status_options = sorted(dashboard_df["Status"].dropna().unique().tolist())
    selected_statuses = st.multiselect("📌 Status", status_options, default=status_options)

    priority_options = sorted(dashboard_df["Priority"].dropna().unique().tolist())
    selected_priorities = st.multiselect("🔥 Priority", priority_options, default=priority_options)

    valid_dates = dashboard_df["Report Date"].dropna()
    selected_dates = None
    if not valid_dates.empty:
        selected_dates = st.date_input(
            "📅 Report Date Range",
            value=(valid_dates.min().date(), valid_dates.max().date())
        )

filtered = dashboard_df.copy()
if selected_properties:
    filtered = filtered[filtered["Property"].isin(selected_properties)]
if selected_area_types:
    filtered = filtered[filtered["Area Type"].isin(selected_area_types)]
if selected_specific_areas:
    filtered = filtered[filtered["Specific Area"].astype(str).isin(selected_specific_areas)]
if selected_sources:
    filtered = filtered[filtered["Source File"].isin(selected_sources)]
if selected_teams:
    filtered = filtered[filtered["Team"].isin(selected_teams)]
if selected_statuses:
    filtered = filtered[filtered["Status"].isin(selected_statuses)]
if selected_priorities:
    filtered = filtered[filtered["Priority"].isna() | filtered["Priority"].isin(selected_priorities)]
if selected_dates and len(selected_dates) == 2:
    start_date = pd.Timestamp(selected_dates[0])
    end_date = pd.Timestamp(selected_dates[1]) + pd.Timedelta(days=1)
    filtered = filtered[
        filtered["Report Date"].isna()
        | ((filtered["Report Date"] >= start_date) & (filtered["Report Date"] < end_date))
    ]

# Apply the same Property / Area / Date scope to all additional data sources.
filtered_work_orders = filter_auxiliary_data(
    work_orders, "Date Created", selected_properties, selected_area_types, selected_specific_areas, selected_dates
) if not work_orders.empty else pd.DataFrame()
filtered_incidents = filter_auxiliary_data(
    incidents, "Creation Time", selected_properties, selected_area_types, selected_specific_areas, selected_dates
) if not incidents.empty else pd.DataFrame()

# =========================================================
# EXPORT PDF ANALYTIC REPORT
# =========================================================

property_scope = ", ".join(selected_properties) if selected_properties else "No Property Selected"
area_scope_text = ", ".join(selected_area_types) if selected_area_types else "No Area Selected"
if selected_dates and len(selected_dates) == 2:
    period_scope = f"{selected_dates[0]} to {selected_dates[1]}"
else:
    period_scope = "All available dates"

filter_context = f"Property: {property_scope} | Area: {area_scope_text} | Period: {period_scope}"

with st.sidebar:
    st.divider()
    st.header("📄 Analytic Report")
    st.caption("Export a management-ready PDF based on the current active filters.")
    if st.button("📄 Generate PDF Report", use_container_width=True):
        with st.spinner("Generating analytic report..."):
            st.session_state["stayplease_pdf_report"] = build_pdf_report(filtered, filter_context, filtered_work_orders, filtered_incidents)
            st.session_state["stayplease_pdf_name"] = (
                f"stayplease_analytic_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            )

    if st.session_state.get("stayplease_pdf_report"):
        st.download_button(
            "⬇️ Download PDF Report",
            data=st.session_state["stayplease_pdf_report"],
            file_name=st.session_state.get("stayplease_pdf_name", "stayplease_analytic_report.pdf"),
            mime="application/pdf",
            use_container_width=True
        )

# =========================================================
# TABS
# =========================================================

overview_tab, operations_tab, defects_tab, incident_tab, insights_tab, team_tab, explorer_tab, incident_explorer_tab = st.tabs([
    "🏠 Executive Overview",
    "📊 Operational Analytics",
    "🔧 Defect & Work Orders",
    "🚨 Incident Analytics",
    "🔎 Cross-Source Insights",
    "👥 Team Performance",
    "📋 Task Explorer",
    "🚨 Incident Explorer"
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

    if not filtered_work_orders.empty or not filtered_incidents.empty:
        x1,x2,x3,x4=st.columns(4)
        x1.metric("🔧 Work Orders", f"{len(filtered_work_orders):,}" if not filtered_work_orders.empty else "-")
        x2.metric("🚨 Incidents", f"{len(filtered_incidents):,}" if not filtered_incidents.empty else "-")
        open_i=int(filtered_incidents["Status"].astype(str).str.lower().eq("open").sum()) if not filtered_incidents.empty else 0
        eng_i=int(filtered_incidents["Department"].fillna("").astype(str).str.contains("engineering",case=False,na=False).sum()) if not filtered_incidents.empty else 0
        x3.metric("🔴 Open Incidents", f"{open_i:,}" if not filtered_incidents.empty else "-")
        x4.metric("🔧 Eng. Incident Involvement", f"{eng_i:,}" if not filtered_incidents.empty else "-")

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
        fig.update_traces(
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Tasks: %{value:,}<br>"
                "Share: %{percent:.1%}"
                "<extra></extra>"
            )
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        team_counts = (
            filtered.groupby("Team")
            .size()
            .reset_index(name="Tasks")
            .sort_values("Tasks", ascending=True)
        )

        team_total = int(team_counts["Tasks"].sum())
        team_counts["Share %"] = (team_counts["Tasks"] / team_total * 100).round(1) if team_total else 0
        fig = px.bar(
            team_counts, x="Tasks", y="Team",
            orientation="h", text="Tasks",
            custom_data=["Team", "Tasks", "Share %"],
            title="Tasks by Team"
        )
        fig.update_traces(hovertemplate="<b>%{customdata[0]}</b><br>Tasks: %{customdata[1]:,}<br>Share: %{customdata[2]:.1f}%<extra></extra>")
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

    hour_data = filtered.dropna(subset=["Request Hour"])
    if not hour_data.empty:
        hourly = hour_data.groupby("Request Hour").size().reset_index(name="Tasks")
        hour_total = int(hourly["Tasks"].sum())
        hourly["Share %"] = (hourly["Tasks"] / hour_total * 100).round(1) if hour_total else 0
        fig = px.bar(hourly, x="Request Hour", y="Tasks", custom_data=["Request Hour", "Tasks", "Share %"], title="🕒 Peak Request Hour")
        fig.update_traces(hovertemplate="<b>Hour %{customdata[0]:02d}:00</b><br>Tasks: %{customdata[1]:,}<br>Share: %{customdata[2]:.1f}%<extra></extra>")
        fig.update_xaxes(dtick=1)
        st.plotly_chart(fig, use_container_width=True)


# =========================================================
# DEFECT ANALYTICS
# =========================================================

with defects_tab:
    st.subheader("🔧 Defect & Work Order Analytics")

    if filtered_work_orders.empty:
        st.info("Upload a Work Order Report to unlock dedicated maintenance analytics. The existing Engineering task analytics are shown below.")
    else:
        w1, w2, w3, w4 = st.columns(4)
        wo_total = len(filtered_work_orders)
        wo_done = int(filtered_work_orders["Status"].astype(str).str.lower().eq("done").sum())
        wo_paused = int(filtered_work_orders["Status"].astype(str).str.lower().str.contains("pause").sum())
        w1.metric("🔧 Total Work Orders", f"{wo_total:,}")
        w2.metric("✅ Done", f"{wo_done:,}")
        w3.metric("⏸️ Paused", f"{wo_paused:,}")
        w4.metric("🎯 Completion Rate", f"{(wo_done/wo_total*100 if wo_total else 0):.1f}%")

        wc1, wc2 = st.columns(2)
        with wc1:
            top_n_chart(filtered_work_orders.rename(columns={"Work Order Title":"Work Order Issue"}), "Work Order Issue", "🔧 Top 10 Work Order Issues")
        with wc2:
            top_n_chart(filtered_work_orders, "Location", "📍 Top 10 Maintenance Hotspots")

        st.subheader("🔁 Recurring Maintenance Locations")
        recurring = top_n_counts(filtered_work_orders, "Location", 15)
        if not recurring.empty:
            fig = px.bar(recurring.sort_values("Tasks"), x="Tasks", y="Location", orientation="h", text="Tasks", custom_data=["Location","Tasks"], title="Locations with Repeated Work Orders")
            fig.update_traces(hovertemplate="<b>Location %{customdata[0]}</b><br>Work Orders: %{customdata[1]:,}<extra></extra>")
            fig.update_yaxes(type="category")
            st.plotly_chart(fig, use_container_width=True)

        trend_wo = filtered_work_orders.dropna(subset=["Date Created"]).copy()
        if not trend_wo.empty:
            trend_wo["Date"] = trend_wo["Date Created"].dt.date.astype(str)
            t = trend_wo.groupby("Date").size().reset_index(name="Work Orders")
            fig = px.line(t, x="Date", y="Work Orders", markers=True, custom_data=["Date","Work Orders"], title="📈 Work Order Trend")
            fig.update_traces(hovertemplate="<b>%{customdata[0]}</b><br>Work Orders: %{customdata[1]:,}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("🔧 Engineering Task Defect Analytics")

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
                recurring_total = int(recurring_locations["Cases"].sum())
                recurring_locations["Share %"] = (recurring_locations["Cases"] / recurring_total * 100).round(1) if recurring_total else 0
                fig = px.bar(
                    recurring_locations,
                    x="Cases",
                    y="Location",
                    orientation="h",
                    text="Cases",
                    custom_data=["Location", "Cases", "Share %"],
                    title=f"📍 Top Recurring Locations — {selected_defect}"
                )
                fig.update_traces(hovertemplate=f"<b>{selected_defect}</b><br>Location: %{{customdata[0]}}<br>Cases: %{{customdata[1]:,}}<br>Share of displayed locations: %{{customdata[2]:.1f}}%<extra></extra>")
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

                trend_total = int(trend_data["Cases"].sum())
                trend_data["Share %"] = (trend_data["Cases"] / trend_total * 100).round(1) if trend_total else 0
                fig = px.line(
                    trend_data,
                    x="Month",
                    y="Cases",
                    markers=True,
                    custom_data=["Month", "Cases", "Share %"],
                    title=f"📈 Trend Over Time — {selected_defect}"
                )
                fig.update_traces(hovertemplate=f"<b>{selected_defect}</b><br>Month: %{{customdata[0]}}<br>Cases: %{{customdata[1]:,}}<br>Share of selected defect: %{{customdata[2]:.1f}}%<extra></extra>")
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

        # Count ALL issues first so the room breakdown always reconciles
        # with the total shown in Top 10 Problematic Rooms / Locations.
        all_issue_types = (
            selected_location_df.dropna(subset=["Task"])
            .assign(Task=lambda x: x["Task"].astype(str).str.strip())
            .groupby("Task")
            .size()
            .reset_index(name="Cases")
            .sort_values("Cases", ascending=False)
        )

        total_location_cases = int(all_issue_types["Cases"].sum()) if not all_issue_types.empty else 0

        # Keep the chart readable: Top 10 issue types + Other Issues.
        # For "Other Issues", keep the underlying issue list so users can see
        # exactly what is grouped into the Other bucket when hovering the bar.
        top_issues = all_issue_types.head(10).copy()
        other_issue_types = all_issue_types.iloc[10:].copy()
        other_cases = int(other_issue_types["Cases"].sum())

        top_issues["Other Details"] = ""

        if other_cases > 0:
            other_breakdown = "<br>".join(
                [
                    f"• {row.Task}: {int(row.Cases):,}"
                    for row in other_issue_types.itertuples(index=False)
                ]
            )
            top_issues = pd.concat(
                [
                    top_issues,
                    pd.DataFrame([
                        {
                            "Task": "Other Issues",
                            "Cases": other_cases,
                            "Other Details": other_breakdown
                        }
                    ])
                ],
                ignore_index=True
            )

        issue_types = top_issues.sort_values("Cases", ascending=True)

        st.metric(
            f"📊 Total Defect / Issue Cases — Location {selected_location}",
            f"{total_location_cases:,}"
        )

        if issue_types.empty:
            st.info("No defect/issue type data available for this location.")
        else:
            issue_types["Share %"] = (issue_types["Cases"] / total_location_cases * 100).round(1) if total_location_cases else 0
            fig = px.bar(
                issue_types,
                x="Cases",
                y="Task",
                orientation="h",
                text="Cases",
                custom_data=["Task", "Cases", "Share %", "Other Details"],
                title=f"🔧 Types of Defects / Issues — Location {selected_location}"
            )
            fig.update_traces(
                hovertemplate=(
                    f"<b>%{{customdata[0]}}</b><br>"
                    f"Location: {selected_location}<br>"
                    f"Cases: %{{customdata[1]:,}}<br>"
                    f"Share of location total: %{{customdata[2]:.1f}}%"
                    f"%{{customdata[3]}}<extra></extra>"
                )
            )
            fig.update_layout(
                showlegend=False,
                yaxis_title="",
                xaxis_title="Cases"
            )
            fig.update_yaxes(type="category")
            st.plotly_chart(fig, use_container_width=True)

            st.caption(
                f"Breakdown reconciliation: {total_location_cases:,} total cases "
                f"shown in the selected location."
            )


# =========================================================
# INCIDENT ANALYTICS
# =========================================================

with incident_tab:
    st.subheader("🚨 Incident Analytics")
    if filtered_incidents.empty:
        st.info("Upload an Incident List Report to unlock incident analytics for the current Property / Area / Date scope.")
    else:
        total_inc = len(filtered_incidents)
        open_inc = int(filtered_incidents["Status"].astype(str).str.lower().eq("open").sum())
        closed_inc = int(filtered_incidents["Status"].astype(str).str.lower().isin(["close", "closed"]).sum())
        closure_rate = closed_inc / total_inc * 100 if total_inc else 0
        eng_inc = int(filtered_incidents["Department"].fillna("").astype(str).str.contains("engineering", case=False, na=False).sum())
        high_impact_terms = r"angry|annoyed|impatient"
        high_impact = int(filtered_incidents["Guest Temp"].fillna("").astype(str).str.contains(high_impact_terms, case=False, regex=True, na=False).sum())

        i1,i2,i3,i4,i5=st.columns(5)
        i1.metric("🚨 Total Incidents", f"{total_inc:,}")
        i2.metric("🔴 Open Incidents", f"{open_inc:,}")
        i3.metric("🟢 Closure Rate", f"{closure_rate:.1f}%")
        i4.metric("🔧 Engineering Involved", f"{eng_inc:,}")
        i5.metric("😡 High Guest Impact", f"{high_impact:,}")

        # Trend + status
        c1,c2=st.columns(2)
        with c1:
            trend_src = filtered_incidents.dropna(subset=["Creation Time"]).copy()
            if not trend_src.empty:
                trend_src["Date"] = trend_src["Creation Time"].dt.date
                trend = trend_src.groupby("Date").size().reset_index(name="Incidents")
                fig=px.line(trend,x="Date",y="Incidents",markers=True,title="📈 Incident Trend Over Time")
                fig.update_traces(hovertemplate="<b>%{x}</b><br>Incidents: %{y:,}<extra></extra>")
                fig.update_layout(xaxis_title="Date",yaxis_title="Incidents")
                st.plotly_chart(fig,use_container_width=True)
            else:
                st.info("No valid incident creation dates available.")
        with c2:
            status = filtered_incidents["Status"].fillna("Unknown").value_counts().reset_index()
            status.columns=["Status","Cases"]
            fig=px.pie(status,names="Status",values="Cases",hole=.55,title="Incident Status Distribution")
            fig.update_traces(hovertemplate="<b>%{label}</b><br>Cases: %{value:,}<br>Share: %{percent:.1%}<extra></extra>")
            st.plotly_chart(fig,use_container_width=True)

        # Categories + locations
        c3,c4=st.columns(2)
        with c3:
            cats=explode_incident_categories(filtered_incidents).head(10).sort_values("Cases")
            if not cats.empty:
                fig=px.bar(cats,x="Cases",y="Category",orientation="h",text="Cases",custom_data=["Category","Cases"],title="🏷️ Top Incident Categories")
                fig.update_traces(hovertemplate="<b>%{customdata[0]}</b><br>Incidents: %{customdata[1]:,}<extra></extra>")
                fig.update_layout(yaxis_title="",xaxis_title="Incidents")
                st.plotly_chart(fig,use_container_width=True)
        with c4:
            # Locations are room/location labels, not numeric measurements.
            # Force categorical ordering so Plotly never interprets room numbers
            # (for example 7301, 8611, 8809) as a continuous numeric axis.
            hotspots = top_n_counts(filtered_incidents, "Location", 10).rename(columns={"Tasks": "Cases"})
            if not hotspots.empty:
                hotspots["Location"] = hotspots["Location"].astype(str).str.strip()
                hotspots = hotspots.sort_values(["Cases", "Location"], ascending=[True, True]).copy()
                location_order = hotspots["Location"].tolist()

                fig = px.bar(
                    hotspots,
                    x="Cases",
                    y="Location",
                    orientation="h",
                    text="Cases",
                    custom_data=["Location", "Cases"],
                    category_orders={"Location": location_order},
                    title="📍 Top Incident Locations"
                )
                fig.update_traces(
                    hovertemplate="<b>Location %{customdata[0]}</b><br>Incidents: %{customdata[1]:,}<extra></extra>",
                    textposition="outside",
                    cliponaxis=False
                )
                fig.update_yaxes(
                    type="category",
                    categoryorder="array",
                    categoryarray=location_order,
                    title=""
                )
                fig.update_xaxes(title="Incidents", rangemode="tozero", dtick=1)
                fig.update_layout(
                    showlegend=False,
                    margin=dict(l=20, r=40, t=55, b=35),
                    bargap=0.28
                )
                st.plotly_chart(fig, use_container_width=True)

        # Guest impact + departments
        c5,c6=st.columns(2)
        with c5:
            temp=filtered_incidents["Guest Temp"].fillna("Not specified").astype(str).str.strip().replace("","Not specified").value_counts().reset_index()
            temp.columns=["Guest Temp","Cases"]
            fig=px.bar(temp,x="Guest Temp",y="Cases",text="Cases",custom_data=["Guest Temp","Cases"],title="😟 Guest Temperature / Impact")
            fig.update_traces(hovertemplate="<b>%{customdata[0]}</b><br>Cases: %{customdata[1]:,}<extra></extra>")
            st.plotly_chart(fig,use_container_width=True)
        with c6:
            dept=filtered_incidents[["Department"]].dropna().copy()
            if not dept.empty:
                dept["Department"]=dept["Department"].astype(str).str.split(",")
                dept=dept.explode("Department")
                dept["Department"]=dept["Department"].astype(str).str.strip()
                dept=dept[dept["Department"].ne("")]
                d=dept.groupby("Department").size().reset_index(name="Incident Involvement").sort_values("Incident Involvement",ascending=True)
                fig=px.bar(d,x="Incident Involvement",y="Department",orientation="h",text="Incident Involvement",custom_data=["Department","Incident Involvement"],title="👥 Department Involvement")
                fig.update_traces(hovertemplate="<b>%{customdata[0]}</b><br>Incident involvement: %{customdata[1]:,}<extra></extra>")
                fig.update_layout(yaxis_title="")
                st.plotly_chart(fig,use_container_width=True)

        # Service recovery is optional because many reports may have zero/blank values.
        st.subheader("💰 Service Recovery Analytics")
        recovery = filtered_incidents.copy()
        recovery["Has Compensation"] = (
            recovery["Compensation"].fillna("").astype(str).str.strip().ne("")
            | recovery["Other Compensation"].fillna("").astype(str).str.strip().ne("")
            | recovery["Cost"].fillna(0).gt(0)
        )
        compensated = int(recovery["Has Compensation"].sum())
        total_cost = float(recovery.loc[recovery["Cost"].notna(), "Cost"].sum())
        avg_cost = float(recovery.loc[recovery["Cost"] > 0, "Cost"].mean()) if (recovery["Cost"] > 0).any() else 0.0
        r1,r2,r3=st.columns(3)
        r1.metric("🎁 Incidents with Recovery", f"{compensated:,}")
        r2.metric("💰 Recorded Recovery Cost", f"{total_cost:,.2f}")
        r3.metric("📊 Avg Positive Cost", f"{avg_cost:,.2f}")

        compensation = recovery["Compensation"].fillna("").astype(str).str.strip()
        compensation = compensation[compensation.ne("")]
        if not compensation.empty:
            comp = compensation.value_counts().head(10).sort_values().reset_index()
            comp.columns=["Compensation","Cases"]
            fig=px.bar(comp,x="Cases",y="Compensation",orientation="h",text="Cases",title="Top Service Recovery Types")
            st.plotly_chart(fig,use_container_width=True)
        else:
            st.caption("No structured compensation type was recorded in the selected scope.")

# =========================================================
# CROSS-SOURCE INSIGHTS
# =========================================================

with insights_tab:
    st.subheader("🔎 Cross-Source Operational Insights")
    st.caption(
        "This view compares Work Order and Incident activity at the same mapped location during the selected period. "
        "It helps identify where maintenance and guest-impact records overlap; it does not prove that one caused the other."
    )

    if filtered_work_orders.empty or filtered_incidents.empty:
        st.info("Upload both Work Order Report and Incident List Report to enable cross-source insights.")
    else:
        wo_loc = filtered_work_orders.dropna(subset=["Location"]).copy()
        wo_loc["Location"] = wo_loc["Location"].astype(str).str.strip()
        inc_loc = filtered_incidents.dropna(subset=["Location"]).copy()
        inc_loc["Location"] = inc_loc["Location"].astype(str).str.strip()

        wo_counts = wo_loc.groupby("Location").size().reset_index(name="Work Orders")
        inc_counts = inc_loc.groupby("Location").size().reset_index(name="Incidents")
        open_counts = (
            inc_loc[inc_loc["Status"].astype(str).str.lower().eq("open")]
            .groupby("Location").size().reset_index(name="Open Incidents")
        )

        overlap = (
            wo_counts.merge(inc_counts, on="Location", how="inner")
            .merge(open_counts, on="Location", how="left")
            .fillna({"Open Incidents": 0})
        )

        if overlap.empty:
            st.info("No mapped locations appear in both Work Order and Incident reports for the selected scope.")
        else:
            overlap["Open Incidents"] = overlap["Open Incidents"].astype(int)
            overlap["Combined Activity"] = overlap["Work Orders"] + overlap["Incidents"]
            overlap = overlap.sort_values(
                ["Combined Activity", "Incidents", "Work Orders"], ascending=False
            )

            k1, k2, k3 = st.columns(3)
            k1.metric("📍 Shared Locations", f"{len(overlap):,}")
            k2.metric("🔧🚨 Combined Activity", f"{int(overlap['Combined Activity'].sum()):,}")
            k3.metric("🔴 Shared Locations with Open Incidents", f"{int((overlap['Open Incidents'] > 0).sum()):,}")

            with st.expander("ℹ️ How to read this analysis", expanded=True):
                st.markdown(
                    "- **Shared Location** means the same mapped location appears in both source reports.\n"
                    "- **Work Orders** show recorded maintenance activity.\n"
                    "- **Incidents** show recorded incident cases.\n"
                    "- **Combined Activity** is only a simple volume indicator to help prioritize where to review first.\n"
                    "- Use the location detail below to inspect the actual maintenance issues and incident categories before drawing conclusions."
                )

            st.subheader("📋 Location Overlap Summary")
            st.dataframe(
                overlap[["Location", "Work Orders", "Incidents", "Open Incidents", "Combined Activity"]],
                use_container_width=True, hide_index=True
            )

            chart = overlap.head(15).sort_values("Combined Activity")
            fig = px.bar(
                chart, x="Combined Activity", y="Location", orientation="h", text="Combined Activity",
                custom_data=["Location", "Work Orders", "Incidents", "Open Incidents", "Combined Activity"],
                title="📍 Locations with the Highest Cross-Source Activity"
            )
            fig.update_traces(
                hovertemplate=(
                    "<b>Location %{customdata[0]}</b><br>"
                    "Work Orders: %{customdata[1]:,}<br>"
                    "Incidents: %{customdata[2]:,}<br>"
                    "Open Incidents: %{customdata[3]:,}<br>"
                    "Combined Activity: %{customdata[4]:,}<extra></extra>"
                )
            )
            fig.update_layout(showlegend=False, yaxis_title="", xaxis_title="Records")
            fig.update_yaxes(type="category")
            st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.subheader("🔍 Location Detail Explorer")
            selected_insight_location = st.selectbox(
                "Select a shared location to inspect",
                overlap["Location"].astype(str).tolist(),
                key="cross_source_location"
            )

            selected_wo = wo_loc[wo_loc["Location"] == selected_insight_location].copy()
            selected_inc = inc_loc[inc_loc["Location"] == selected_insight_location].copy()

            d1, d2 = st.columns(2)
            with d1:
                st.markdown("#### 🔧 Maintenance activity")
                st.metric("Work Orders", f"{len(selected_wo):,}")
                if not selected_wo.empty:
                    wo_issues = (
                        selected_wo["Work Order Title"].fillna("Not specified").astype(str)
                        .value_counts().head(10).reset_index()
                    )
                    wo_issues.columns = ["Work Order Issue", "Cases"]
                    st.dataframe(wo_issues, use_container_width=True, hide_index=True)

            with d2:
                st.markdown("#### 🚨 Incident activity")
                st.metric("Incidents", f"{len(selected_inc):,}")
                if not selected_inc.empty:
                    incident_cats = explode_incident_categories(selected_inc).head(10)
                    st.markdown("**Top incident categories**")
                    st.dataframe(incident_cats, use_container_width=True, hide_index=True)
                    incident_status = (
                        selected_inc["Status"].fillna("Unknown").astype(str).value_counts()
                        .reset_index()
                    )
                    incident_status.columns = ["Status", "Cases"]
                    st.markdown("**Incident status**")
                    st.dataframe(incident_status, use_container_width=True, hide_index=True)

            open_here = int(selected_inc["Status"].astype(str).str.lower().eq("open").sum())
            if open_here > 0:
                st.warning(
                    f"Review priority: Location {selected_insight_location} has {open_here} open incident(s). "
                    "Use the issue details above to coordinate the next operational review."
                )
            else:
                st.info(
                    f"Location {selected_insight_location} has overlap between maintenance and incident records, "
                    "but no open incident is currently recorded in the selected scope."
                )

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
        "Property", "Area Type", "Specific Area",
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
# INCIDENT EXPLORER
# =========================================================

with incident_explorer_tab:
    st.subheader("🚨 Incident Explorer")
    st.caption("Inspect the underlying Incident List Report records using the current global Property / Area / Date scope.")

    if filtered_incidents.empty:
        st.info("Upload an Incident List Report to use the Incident Explorer.")
    else:
        ie1, ie2, ie3, ie4 = st.columns(4)
        with ie1:
            incident_locations = sorted(filtered_incidents["Location"].dropna().astype(str).unique().tolist())
            selected_incident_locations = st.multiselect("📍 Location", incident_locations, default=incident_locations, key="incident_explorer_location")
        with ie2:
            incident_status_options = sorted(filtered_incidents["Status"].dropna().astype(str).unique().tolist())
            selected_incident_status = st.multiselect("📌 Status", incident_status_options, default=incident_status_options, key="incident_explorer_status")
        with ie3:
            incident_dept_options = sorted(filtered_incidents["Department"].dropna().astype(str).unique().tolist())
            selected_incident_departments = st.multiselect("👥 Department", incident_dept_options, default=incident_dept_options, key="incident_explorer_department")
        with ie4:
            incident_search = st.text_input("🔎 Search", placeholder="Incident, guest, location, log no...", key="incident_explorer_search")

        incident_display = filtered_incidents.copy()
        if selected_incident_locations:
            incident_display = incident_display[incident_display["Location"].astype(str).isin(selected_incident_locations)]
        if selected_incident_status:
            incident_display = incident_display[incident_display["Status"].astype(str).isin(selected_incident_status)]
        if selected_incident_departments:
            incident_display = incident_display[incident_display["Department"].astype(str).isin(selected_incident_departments)]
        if incident_search:
            search_cols = ["Log No", "Location", "Incident Name", "Status", "Created By", "Department", "Guest Name", "More Information", "Guest Feedback"]
            mask = pd.Series(False, index=incident_display.index)
            for col in search_cols:
                mask |= incident_display[col].fillna("").astype(str).str.contains(incident_search, case=False, na=False)
            incident_display = incident_display[mask]

        # Make compensation value explicit and easy to read in the explorer.
        # The source report stores the monetary amount in Cost, while Compensation
        # and Other Compensation describe the recovery type/details.
        incident_display["Compensation Type"] = (
            incident_display["Compensation"].fillna("").astype(str).str.strip()
        )
        other_comp = incident_display["Other Compensation"].fillna("").astype(str).str.strip()
        incident_display.loc[
            incident_display["Compensation Type"].eq("") & other_comp.ne(""),
            "Compensation Type"
        ] = other_comp
        incident_display["Compensation Type"] = incident_display["Compensation Type"].replace("", "—")

        cost_numeric = pd.to_numeric(incident_display["Cost"], errors="coerce").fillna(0)
        incident_display["Compensation Amount"] = cost_numeric.apply(
            lambda x: f"IDR {x:,.0f}" if x > 0 else "—"
        )

        # Quick financial visibility for the currently filtered incident records.
        positive_cost = cost_numeric[cost_numeric > 0]
        c1, c2, c3 = st.columns(3)
        c1.metric("💰 Incidents with Compensation", f"{int((cost_numeric > 0).sum()):,}")
        c2.metric("💵 Total Compensation", f"IDR {positive_cost.sum():,.0f}")
        c3.metric(
            "📊 Avg Compensation",
            f"IDR {positive_cost.mean():,.0f}" if not positive_cost.empty else "—"
        )

        incident_columns = [
            "Property", "Area Type", "Specific Area", "Log No", "Location",
            "Incident Name", "Compensation Amount", "Compensation Type",
            "Status", "Creation Time", "Department", "Guest Temp", "VIP Level",
            "Guest Name", "Deadline", "More Information", "Guest Feedback"
        ]
        incident_table = incident_display[[c for c in incident_columns if c in incident_display.columns]].copy()
        st.caption(f"Showing {len(incident_table):,} incident record(s)")
        st.dataframe(
            incident_table.sort_values("Creation Time", ascending=False) if "Creation Time" in incident_table.columns else incident_table,
            use_container_width=True, hide_index=True, height=600,
            column_config={
                "Compensation Amount": st.column_config.TextColumn("💰 Compensation Amount"),
                "Compensation Type": st.column_config.TextColumn("🎁 Compensation Type"),
            }
        )
        incident_csv = incident_table.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Filtered Incidents (CSV)",
            data=incident_csv, file_name="stayplease_filtered_incidents.csv",
            mime="text/csv", key="download_incident_csv"
        )

st.divider()
st.caption(
    f"🏨 StayPlease Operational Intelligence | "
    f"📊 {len(filtered):,} filtered tasks | "
    f"🏨 {', '.join(sorted(filtered['Property'].dropna().unique().tolist())) if not filtered.empty else '-'} | "
    f"🏢 {filtered['Team'].nunique()} team(s) | "
    f"📁 {filtered['Source File'].nunique()} task file(s) | 🔧 {len(filtered_work_orders):,} work orders | 🚨 {len(filtered_incidents):,} incidents"
)
