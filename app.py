import html
import sqlite3
from datetime import timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import plotly.express as px

# --------------------------------------------------
# Page setup
# --------------------------------------------------

st.set_page_config(
    page_title="Endurance Training Platform",
    page_icon="🏃",
    layout="wide"
)

st.title("🏃 Endurance Training Platform")

# --------------------------------------------------
# Database connection
# --------------------------------------------------

conn = sqlite3.connect("output/training.db")
cursor = conn.cursor()

# --------------------------------------------------
# KPI Queries
# --------------------------------------------------

cursor.execute("SELECT COUNT(*) FROM activities")
total_activities = cursor.fetchone()[0]

cursor.execute("""
SELECT COUNT(*)
FROM activities
WHERE LOWER(sport) LIKE '%running%'
""")
running = cursor.fetchone()[0]

cursor.execute("""
SELECT COUNT(*)
FROM activities
WHERE LOWER(sport) LIKE '%cycling%'
""")
cycling = cursor.fetchone()[0]

cursor.execute("""
SELECT COUNT(*)
FROM activities
WHERE LOWER(sport) LIKE '%swimming%'
""")
swimming = cursor.fetchone()[0]

cursor.execute("""
SELECT COUNT(*)
FROM activities
WHERE LOWER(sport) LIKE '%training%'
""")
strength = cursor.fetchone()[0]

# --------------------------------------------------
# KPI Cards
# --------------------------------------------------

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("Activities", total_activities)
col2.metric("Running", running)
col3.metric("Cycling", cycling)
col4.metric("Swimming", swimming)
col5.metric("Strength", strength)

st.success("Database connection successful.")

# --------------------------------------------------
# Recent Activities
# --------------------------------------------------

st.subheader("Recent Activities")

cursor.execute("""
SELECT
    activity_date,
    session_name,
    sport,
    distance_km
FROM activities
ORDER BY activity_date DESC
LIMIT 20
""")

recent_activities = cursor.fetchall()

recent_df = pd.DataFrame(
    recent_activities,
    columns=[
        "Date",
        "Session",
        "Sport",
        "Distance (km)"
    ]
)

recent_df["Distance (km)"] = recent_df["Distance (km)"].round(2)

st.dataframe(
    recent_df,
    use_container_width=True,
    hide_index=True
)

# --------------------------------------------------
# Weekly Training Time by Sport
# --------------------------------------------------

st.subheader("Weekly Training Time by Sport")

cursor.execute("""
SELECT
    activity_date,
    sport,
    time_hms
FROM activities
""")

weekly_data = cursor.fetchall()

chart_rows = []

for activity_date, sport, time_hms in weekly_data:

    if time_hms is None:
        continue

    parts = str(time_hms).split(":")

    if len(parts) != 3:
        continue

    hours = (
        int(parts[0])
        + int(parts[1]) / 60
        + int(parts[2]) / 3600
    )

    activity_date = pd.to_datetime(activity_date)

    week_start = (
        activity_date
        - pd.Timedelta(days=activity_date.weekday())
    )

    chart_rows.append({
        "Week": week_start,
        "Sport": sport,
        "Hours": hours
    })

chart_df = pd.DataFrame(chart_rows)

chart_df["Sport"] = chart_df["Sport"].fillna("Other")

chart_df["Sport"] = chart_df["Sport"].replace({
    "running": "Run",
    "cycling": "Bike",
    "swimming": "Swim",
    "training": "Strength"
})

weekly_totals = (
    chart_df
    .groupby(["Week", "Sport"])["Hours"]
    .sum()
    .reset_index()
)

fig = px.bar(
    weekly_totals,
    x="Week",
    y="Hours",
    color="Sport",
    title="Weekly Training Time by Sport",
    barmode="stack",
    color_discrete_map={
        "Run": "#70AD47",
        "Bike": "#ED7D31",
        "Swim": "#5B9BD5",
        "Strength": "#8064A2",
        "Other": "#A5A5A5"
    }
)

st.plotly_chart(
    fig,
    use_container_width=True
)

# --------------------------------------------------
# Weekly Training Load Summary
# --------------------------------------------------

st.subheader("Weekly Training Load Summary")

weekly_load = (
    chart_df
    .groupby("Week")
    .agg(
        Hours=("Hours", "sum"),
        Sessions=("Hours", "count")
    )
    .reset_index()
)

weekly_load["Hours"] = weekly_load["Hours"].round(1)

weekly_load["Week"] = (
    pd.to_datetime(weekly_load["Week"])
    .dt.strftime("%Y-%m-%d")
)

weekly_load = weekly_load.sort_values(
    "Week",
    ascending=False
)

st.dataframe(
    weekly_load,
    use_container_width=True,
    hide_index=True
)

# --------------------------------------------------
# Weekly Distance by Sport
# --------------------------------------------------

st.subheader("Weekly Distance by Sport")

cursor.execute("""
SELECT
    activity_date,
    sport,
    distance_km
FROM activities
""")

distance_data = cursor.fetchall()

distance_rows = []

for activity_date, sport, distance_km in distance_data:

    if distance_km is None:
        continue

    activity_date = pd.to_datetime(activity_date)

    week_start = (
        activity_date
        - pd.Timedelta(days=activity_date.weekday())
    )

    sport_text = str(sport).lower()

    if "running" in sport_text:
        sport_category = "Run"
    elif "cycling" in sport_text:
        sport_category = "Bike"
    elif "swimming" in sport_text:
        sport_category = "Swim"
    elif "training" in sport_text:
        sport_category = "Strength"
    else:
        sport_category = "Other"

    distance_rows.append({
        "Week": week_start,
        "Sport": sport_category,
        "Distance": distance_km
    })

distance_df = pd.DataFrame(distance_rows)

weekly_distance_totals = (
    distance_df
    .groupby(["Week", "Sport"])["Distance"]
    .sum()
    .reset_index()
)

distance_chart = px.bar(
    weekly_distance_totals,
    x="Week",
    y="Distance",
    color="Sport",
    title="Weekly Distance by Sport",
    barmode="stack",
    color_discrete_map={
        "Run": "#70AD47",
        "Bike": "#ED7D31",
        "Swim": "#5B9BD5",
        "Strength": "#8064A2",
        "Other": "#A5A5A5"
    }
)

st.plotly_chart(
    distance_chart,
    use_container_width=True
)

# --------------------------------------------------
# Training Diary
# --------------------------------------------------

st.subheader("Training Diary")

sport_filter = st.selectbox(
    "Sport Filter",
    [
        "All",
        "Run",
        "Bike",
        "Swim",
        "Strength"
    ]
)

activity_limit = st.slider(
    "Number of Activities",
    min_value=10,
    max_value=200,
    value=50,
    step=10
)

cursor.execute("""
SELECT
    activity_date,
    session_name,
    sport,
    distance_km,
    time_hms
FROM activities
ORDER BY activity_date DESC
""")

diary_rows = cursor.fetchall()

filtered_rows = []

for activity_date, session_name, sport, distance_km, time_hms in diary_rows:

    sport_text = str(sport).lower()

    if "running" in sport_text:
        sport_category = "Run"
    elif "cycling" in sport_text:
        sport_category = "Bike"
    elif "swimming" in sport_text:
        sport_category = "Swim"
    elif "training" in sport_text:
        sport_category = "Strength"
    else:
        sport_category = "Other"

    if sport_filter != "All":
        if sport_category != sport_filter:
            continue

    filtered_rows.append({
        "Date": activity_date,
        "Sport": sport_category,
        "Session": session_name,
        "Distance (km)": round(distance_km, 2)
        if distance_km is not None
        else None,
        "Time": time_hms
    })

diary_df = pd.DataFrame(
    filtered_rows[:activity_limit]
)

st.dataframe(
    diary_df,
    use_container_width=True,
    hide_index=True
)

# --------------------------------------------------
# Training Calendar V1
# --------------------------------------------------

st.subheader("Training Calendar")

cursor.execute("""
SELECT
    activity_date,
    session_name,
    sport,
    distance_km,
    time_hms
FROM activities
WHERE activity_date IS NOT NULL
ORDER BY activity_date
""")

calendar_rows = cursor.fetchall()

# --------------------------------------------------
# Load Planned Sessions
# --------------------------------------------------

cursor.execute("""
SELECT
    planned_date,
    session_name,
    sport,
    planned_distance_km,
    planned_time_min,
    notes
FROM planned_sessions
""")

planned_rows = cursor.fetchall()

planned_df = pd.DataFrame(
    planned_rows,
    columns=[
        "Date",
        "Session",
        "Sport",
        "Planned Distance",
        "Planned Time",
        "Notes"
    ]
)

planned_df["Date"] = pd.to_datetime(
    planned_df["Date"],
    errors="coerce"
)

calendar_df = pd.DataFrame(
    calendar_rows,
    columns=[
        "Date",
        "Session",
        "Sport",
        "Distance (km)",
        "Time"
    ]
)

calendar_df["Date"] = pd.to_datetime(
    calendar_df["Date"],
    errors="coerce"
)

calendar_df = calendar_df.dropna(subset=["Date"])

# Use the latest activity date as the default selected date
if not calendar_df.empty:
    latest_activity_date = calendar_df["Date"].max().date()
else:
    latest_activity_date = pd.Timestamp.today().date()

def calendar_sport_category(sport_value):
    sport_text = str(sport_value).lower()

    if "running" in sport_text:
        return "Run"
    if "cycling" in sport_text:
        return "Bike"
    if "swimming" in sport_text:
        return "Swim"
    if "training" in sport_text or "strength" in sport_text:
        return "Strength"

    return "Other"

# Set the initial calendar week
if "calendar_week_start" not in st.session_state:
    initial_date = pd.Timestamp(latest_activity_date)

    st.session_state.calendar_week_start = (
        initial_date
        - pd.Timedelta(days=initial_date.weekday())
    )

# Week navigation
previous_week_column, week_label_column, next_week_column = st.columns(
    [1, 3, 1]
)

with previous_week_column:
    if st.button(
        "← Previous Week",
        use_container_width=True
    ):
        st.session_state.calendar_week_start -= pd.Timedelta(days=7)
        st.rerun()

with next_week_column:
    if st.button(
        "Next Week →",
        use_container_width=True
    ):
        st.session_state.calendar_week_start += pd.Timedelta(days=7)
        st.rerun()

week_start = pd.Timestamp(
    st.session_state.calendar_week_start
)

week_end = week_start + pd.Timedelta(days=6)

# Filter activities to the selected calendar week
week_activities = calendar_df[
    (calendar_df["Date"] >= week_start)
    & (calendar_df["Date"] < week_start + pd.Timedelta(days=7))
].copy()


# Convert activity duration from HH:MM:SS into seconds
def time_to_seconds(time_value):
    if time_value is None or pd.isna(time_value):
        return 0

    try:
        time_parts = str(time_value).split(":")

        if len(time_parts) != 3:
            return 0

        hours = int(time_parts[0])
        minutes = int(time_parts[1])
        seconds = int(float(time_parts[2]))

        return (
            hours * 3600
            + minutes * 60
            + seconds
        )

    except (TypeError, ValueError):
        return 0


week_activities["Duration Seconds"] = (
    week_activities["Time"].apply(time_to_seconds)
)

# Add the same standardised sport categories used by the calendar
week_activities["Sport Category"] = (
    week_activities["Sport"].apply(calendar_sport_category)
)

# Calculate the weekly totals
weekly_distance = (
    pd.to_numeric(
        week_activities["Distance (km)"],
        errors="coerce"
    )
    .fillna(0)
    .sum()
)

weekly_duration_seconds = int(
    pd.to_numeric(
        week_activities["Duration Seconds"],
        errors="coerce"
    )
    .fillna(0)
    .sum()
)

weekly_sessions = len(week_activities)

running_distance = (
    pd.to_numeric(
        week_activities.loc[
            week_activities["Sport Category"] == "Run",
            "Distance (km)"
        ],
        errors="coerce"
    )
    .fillna(0)
    .sum()
)

cycling_distance = (
    pd.to_numeric(
        week_activities.loc[
            week_activities["Sport Category"] == "Bike",
            "Distance (km)"
        ],
        errors="coerce"
    )
    .fillna(0)
    .sum()
)

swimming_distance = (
    pd.to_numeric(
        week_activities.loc[
            week_activities["Sport Category"] == "Swim",
            "Distance (km)"
        ],
        errors="coerce"
    )
    .fillna(0)
    .sum()
)

# Format total training time as hours and minutes
weekly_duration_hours = weekly_duration_seconds // 3600
weekly_duration_minutes = (
    weekly_duration_seconds % 3600
) // 60

weekly_duration_display = (
    f"{int(weekly_duration_hours)}h "
    f"{int(weekly_duration_minutes):02d}m"
)

with week_label_column:
    st.markdown(
        f"""
        <div style="
            text-align: center;
            padding-top: 7px;
            font-size: 18px;
            font-weight: 600;
        ">
            {week_start.strftime("%d %b %Y")}
            –
            {week_end.strftime("%d %b %Y")}
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("### Weekly Summary")

metric_1, metric_2, metric_3, metric_4 = st.columns(4)

with metric_1:
    st.metric(
        "Distance",
        f"{weekly_distance:.1f} km"
    )

with metric_2:
    st.metric(
        "Training Time",
        weekly_duration_display
    )

with metric_3:
    st.metric(
        "Sessions",
        weekly_sessions
    )

with metric_4:
    st.metric(
        "Running",
        f"{running_distance:.1f} km"
    )

sport_col_1, sport_col_2, sport_col_3 = st.columns(3)

with sport_col_1:
    st.caption(
        f"🏃 Running: {running_distance:.1f} km"
    )

with sport_col_2:
    st.caption(
        f"🚴 Cycling: {cycling_distance:.1f} km"
    )

with sport_col_3:
    st.caption(
        f"🏊 Swimming: {swimming_distance:.1f} km"
    )

selected_week_df = calendar_df[
    (calendar_df["Date"] >= week_start)
    & (calendar_df["Date"] <= week_end)
].copy()

selected_planned_week_df = planned_df[
    (planned_df["Date"] >= week_start)
    & (planned_df["Date"] <= week_end)
].copy()

# --------------------------------------------------
# Planned Week Summary
# --------------------------------------------------

planned_sessions_count = len(
    selected_planned_week_df
)

completed_sessions_count = len(
    week_activities
)

remaining_sessions_count = (
    planned_sessions_count
    - completed_sessions_count
)

if remaining_sessions_count < 0:
    remaining_sessions_count = 0

if planned_sessions_count > 0:
    compliance_pct = (
        completed_sessions_count
        / planned_sessions_count
        * 100
    )
else:
    compliance_pct = 0    

planned_run_km = (
    pd.to_numeric(
        selected_planned_week_df.loc[
            selected_planned_week_df["Sport"]
            .str.lower()
            .str.contains("run", na=False),
            "Planned Distance"
        ],
        errors="coerce"
    )
    .fillna(0)
    .sum()
)

planned_run_sessions = (
    selected_planned_week_df["Sport"]
    .str.lower()
    .str.contains("run", na=False)
    .sum()
)

planned_swims = (
    selected_planned_week_df["Sport"]
    .str.lower()
    .str.contains("swim", na=False)
    .sum()
)

planned_bikes = (
    selected_planned_week_df["Sport"]
    .str.lower()
    .str.contains("bike|cycl", na=False)
    .sum()
)

planned_strength = (
    selected_planned_week_df["Sport"]
    .str.lower()
    .str.contains("strength", na=False)
    .sum()
)

st.markdown("### Planned Week")

planned_col1, planned_col2, planned_col3, planned_col4, planned_col5 = st.columns(5)

planned_col1.metric(
    "Completed",
    f"{completed_sessions_count}/{planned_sessions_count}"
)

planned_col2.metric(
    "Compliance",
    f"{compliance_pct:.0f}%"
)

planned_col3.metric(
    "Running",
    f"{planned_run_km:.0f} km / {planned_run_sessions}"
)

planned_col4.metric(
    "Bike / Swim",
    f"{planned_bikes} / {planned_swims}"
)

planned_col5.metric(
    "Strength",
    planned_strength
)



sport_icons = {
    "Run": "🏃",
    "Bike": "🚴",
    "Swim": "🏊",
    "Strength": "🏋️",
    "Other": "●"
}

sport_colours = {
    "Run": "#70AD47",
    "Bike": "#ED7D31",
    "Swim": "#5B9BD5",
    "Strength": "#8064A2",
    "Other": "#A5A5A5"
}

completed_background_colours = {
    "Run": "#EAF4E3",
    "Bike": "#FFF0E5",
    "Swim": "#E7F3FB",
    "Strength": "#F0EAF7",
    "Other": "#F0F0F0"
}


def pretty_session_name(session_name):
    if session_name is None:
        return "Activity"

    return (
        str(session_name)
        .replace("_", " ")
        .strip()
        .title()
    )


def compact_duration(time_value):
    if time_value is None:
        return ""

    time_parts = str(time_value).split(":")

    if len(time_parts) != 3:
        return str(time_value)

    try:
        hours = int(time_parts[0])
        minutes = int(time_parts[1])
    except ValueError:
        return str(time_value)

    if hours > 0:
        return f"{hours}h {minutes:02d}m"

    return f"{minutes}m"


def distance_text(distance_value, sport_category):
    if distance_value is None:
        return ""

    try:
        distance_value = float(distance_value)
    except (TypeError, ValueError):
        return ""

    if distance_value <= 0:
        return ""

    if sport_category == "Swim":
        return f"{distance_value:.2f} km"

    return f"{distance_value:.1f} km"

# --------------------------------------------------
# Today & Tomorrow
# --------------------------------------------------

st.markdown("## Today & Tomorrow")

# Use UK local time, including daylight-saving changes.
app_today = pd.Timestamp.now(
    tz=ZoneInfo("Europe/London")
).date()

app_tomorrow = app_today + timedelta(days=1)

quick_view_dates = [
    ("TODAY", app_today),
    ("TOMORROW", app_tomorrow)
]

quick_view_columns = st.columns(
    2,
    gap="medium"
)

quick_view_planned_border_colours = {
    "Run": "#9ACD32",
    "Bike": "#F4A261",
    "Swim": "#6EC5FF",
    "Strength": "#B39DDB",
    "Other": "#F2C94C"
}

quick_view_planned_background_colours = {
    "Run": "#F2FBE8",
    "Bike": "#FFF2E6",
    "Swim": "#EAF7FF",
    "Strength": "#F3EDFF",
    "Other": "#FFF8D9"
}


def get_remaining_planned_sessions(
    completed_day_df,
    planned_day_df
):
    """
    Match completed activities against planned sessions one by one,
    using the standardised sport category.

    Example:
    Planned = Run + Bike
    Completed = Run
    Remaining = Bike
    """

    remaining_indices = list(
        planned_day_df.index
    )

    for _, completed_activity in completed_day_df.iterrows():
        completed_category = calendar_sport_category(
            completed_activity["Sport"]
        )

        matching_index = None

        for planned_index in remaining_indices:
            planned_category = calendar_sport_category(
                planned_day_df.loc[
                    planned_index,
                    "Sport"
                ]
            )

            if planned_category == completed_category:
                matching_index = planned_index
                break

        if matching_index is not None:
            remaining_indices.remove(
                matching_index
            )

    return planned_day_df.loc[
        remaining_indices
    ].copy()


def completed_quick_card(activity):
    sport_category = calendar_sport_category(
        activity["Sport"]
    )

    icon = sport_icons.get(
        sport_category,
        "●"
    )

    border_colour = sport_colours.get(
        sport_category,
        "#A5A5A5"
    )

    background_colour = (
        completed_background_colours.get(
            sport_category,
            "#F0F0F0"
        )
    )

    session_name = pretty_session_name(
        activity["Session"]
    )

    distance_display = distance_text(
        activity["Distance (km)"],
        sport_category
    )

    duration_display = compact_duration(
        activity["Time"]
    )

    details = []

    if distance_display:
        details.append(distance_display)

    if duration_display:
        details.append(duration_display)

    details_display = " · ".join(details)

    safe_session_name = html.escape(
        str(session_name)
    )

    safe_details = html.escape(
        str(details_display)
    )

    details_section = ""

    if safe_details:
        details_section = (
            '<div style="'
            'font-size:14px;'
            'margin-top:5px;'
            '">'
            f'{safe_details}'
            '</div>'
        )

    return (
        '<div style="'
        f'border-left:6px solid {border_colour};'
        f'background-color:{background_colour};'
        'padding:12px;'
        'margin:8px 0;'
        'border-radius:7px;'
        '">'
        '<div style="'
        'font-size:11px;'
        'font-weight:700;'
        f'color:{border_colour};'
        'margin-bottom:5px;'
        '">'
        'COMPLETED'
        '</div>'
        '<div style="font-size:21px;">'
        f'{icon}'
        '</div>'
        '<div style="'
        'font-size:16px;'
        'font-weight:650;'
        '">'
        f'{safe_session_name}'
        '</div>'
        f'{details_section}'
        '</div>'
    )


def planned_quick_card(
    planned_activity,
    status_text
):
    sport_category = calendar_sport_category(
        planned_activity["Sport"]
    )

    icon = sport_icons.get(
        sport_category,
        "●"
    )

    border_colour = (
        quick_view_planned_border_colours.get(
            sport_category,
            "#F2C94C"
        )
    )

    background_colour = (
        quick_view_planned_background_colours.get(
            sport_category,
            "#FFF8D9"
        )
    )

    session_name = pretty_session_name(
        planned_activity["Session"]
    )

    if str(session_name).lower() == "nan":
        session_name = "Strength"

    planned_distance = planned_activity[
        "Planned Distance"
    ]

    planned_time = planned_activity[
        "Planned Time"
    ]

    planned_notes = planned_activity[
        "Notes"
    ]

    details = []

    if pd.notna(planned_distance):
        try:
            distance_number = float(
                planned_distance
            )

            if distance_number > 0:
                if sport_category == "Swim":
                    details.append(
                        f"{distance_number:.2f} km"
                    )
                else:
                    details.append(
                        f"{distance_number:g} km"
                    )

        except (TypeError, ValueError):
            pass

    if pd.notna(planned_time):
        time_text = str(
            planned_time
        ).strip()

        if (
            time_text
            and time_text.lower() != "nan"
        ):
            details.append(
                f"{time_text} min"
            )

    details_display = " · ".join(
        details
    )

    notes_display = ""

    if pd.notna(planned_notes):
        notes_text = str(
            planned_notes
        ).strip()

        if (
            notes_text
            and notes_text.lower() != "nan"
        ):
            notes_display = notes_text

    safe_session_name = html.escape(
        str(session_name)
    )

    safe_details = html.escape(
        str(details_display)
    )

    safe_notes = html.escape(
        str(notes_display)
    )

    details_section = ""

    if safe_details:
        details_section = (
            '<div style="'
            'font-size:14px;'
            'font-weight:650;'
            'margin-top:5px;'
            '">'
            f'{safe_details}'
            '</div>'
        )

    notes_section = ""

    if safe_notes:
        notes_section = (
            '<div style="'
            'font-size:13px;'
            'line-height:1.4;'
            'margin-top:7px;'
            'color:#595959;'
            '">'
            f'{safe_notes}'
            '</div>'
        )

    return (
        '<div style="'
        f'border:1px dashed {border_colour};'
        f'border-left:6px solid {border_colour};'
        f'background-color:{background_colour};'
        'padding:12px;'
        'margin:8px 0;'
        'border-radius:7px;'
        '">'
        '<div style="'
        'font-size:11px;'
        'font-weight:700;'
        f'color:{border_colour};'
        'margin-bottom:5px;'
        '">'
        f'{html.escape(status_text)}'
        '</div>'
        '<div style="font-size:21px;">'
        f'{icon}'
        '</div>'
        '<div style="'
        'font-size:16px;'
        'font-weight:650;'
        '">'
        f'{safe_session_name}'
        '</div>'
        f'{details_section}'
        f'{notes_section}'
        '</div>'
    )


for quick_column, (
    day_label,
    quick_date
) in zip(
    quick_view_columns,
    quick_view_dates
):

    quick_completed = calendar_df[
        calendar_df["Date"].dt.date
        == quick_date
    ].copy()

    quick_planned = planned_df[
        planned_df["Date"].dt.date
        == quick_date
    ].copy()

    quick_remaining = (
        get_remaining_planned_sessions(
            quick_completed,
            quick_planned
        )
    )

    with quick_column:
        st.markdown(
            f"### {day_label}"
        )

        st.caption(
            pd.Timestamp(
                quick_date
            ).strftime(
                "%A, %d %B %Y"
            )
        )

        if (
            quick_completed.empty
            and quick_remaining.empty
        ):
            st.info(
                "No training scheduled."
            )

        for _, activity in quick_completed.iterrows():
            st.markdown(
                completed_quick_card(
                    activity
                ),
                unsafe_allow_html=True
            )

        for _, planned_activity in quick_remaining.iterrows():

            if day_label == "TODAY":
                quick_status = "TO DO TODAY"
            else:
                quick_status = "PLANNED TOMORROW"

            st.markdown(
                planned_quick_card(
                    planned_activity,
                    quick_status
                ),
                unsafe_allow_html=True
            )

st.divider()

calendar_columns = st.columns(7)

day_names = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
]

for day_number in range(7):
    calendar_date = week_start + pd.Timedelta(days=day_number)

    day_activities = selected_week_df[
        selected_week_df["Date"].dt.date
        == calendar_date.date()
    ]

    today = pd.Timestamp.today().date()

    day_planned = selected_planned_week_df[
        selected_planned_week_df["Date"].dt.date
        == calendar_date.date()
    ]

        # Match completed activities against planned sessions one by one.
    # A completed Run removes one planned Run, a completed Bike removes
    # one planned Bike, and so on. Unmatched planned sessions remain.
    remaining_planned_indices = list(day_planned.index)

    for _, completed_activity in day_activities.iterrows():
        completed_sport_category = calendar_sport_category(
            completed_activity["Sport"]
        )

        matching_planned_index = None

        for planned_index in remaining_planned_indices:
            planned_sport_category = calendar_sport_category(
                day_planned.loc[planned_index, "Sport"]
            )

            if planned_sport_category == completed_sport_category:
                matching_planned_index = planned_index
                break

        if matching_planned_index is not None:
            remaining_planned_indices.remove(
                matching_planned_index
            )

    day_planned = day_planned.loc[
        remaining_planned_indices
    ].copy()

    with calendar_columns[day_number]:
        st.markdown(
            f"**{day_names[day_number]}**  \n"
            f"{calendar_date.strftime('%d %b')}"
        )

        # Only show Rest day if neither completed nor planned sessions exist.
        if day_activities.empty and day_planned.empty:
            st.caption("Rest day")

        # --------------------------------------------------
        # Completed activities
        # --------------------------------------------------

        for _, activity in day_activities.iterrows():
            sport_category = calendar_sport_category(
                activity["Sport"]
            )

            icon = sport_icons[sport_category]
            colour = sport_colours[sport_category]

            completed_background_colour = (
                completed_background_colours.get(
                    sport_category,
                        "#F0F0F0"
                )
            )

            session_display = pretty_session_name(
                activity["Session"]
            )

            distance_display = distance_text(
                activity["Distance (km)"],
                sport_category
            )

            duration_display = compact_duration(
                activity["Time"]
            )

            activity_details = []

            if distance_display:
                activity_details.append(distance_display)

            if duration_display:
                activity_details.append(duration_display)

            details_display = " · ".join(activity_details)

            st.markdown(
                f"""
                <div style="
                    border-left: 5px solid {colour};
                    background-color: {completed_background_colour};
                    padding: 10px;
                    margin: 8px 0;
                    border-radius: 5px;
                    min-height: 85px;
                ">
                    <div style="
                        font-size: 11px;
                        font-weight: 700;
                        color: {colour};
                        margin-bottom: 4px;
                    ">
                        COMPLETED
                    </div>
                    <div style="font-size: 19px;">
                        {icon}
                    </div>
                    <div style="font-weight: 600;">
                        {session_display}
                    </div>
                    <div style="
                        font-size: 13px;
                        margin-top: 4px;
                    ">
                        {details_display}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        # --------------------------------------------------
        # Planned sessions
        # --------------------------------------------------

        for _, planned_activity in day_planned.iterrows():
            planned_sport_category = calendar_sport_category(
                planned_activity["Sport"]
            )

            if calendar_date.date() < today:
                planned_status = "MISSED"
            else:
                planned_status = "PLANNED"

            planned_icon = sport_icons.get(
                planned_sport_category,
                "●"
            )

            planned_session_display = pretty_session_name(
                planned_activity["Session"]
            )

            if str(planned_session_display).lower() == "nan":
                planned_session_display = "Strength"

            planned_distance = planned_activity[
                "Planned Distance"
            ]

            planned_time = planned_activity[
                "Planned Time"
            ]

            planned_notes = planned_activity["Notes"]

            planned_details = []

            if pd.notna(planned_distance):
                try:
                    planned_distance_number = float(
                        planned_distance
                    )

                    if planned_distance_number > 0:
                        if planned_sport_category == "Swim":
                            planned_details.append(
                                f"{planned_distance_number:.2f} km"
                            )
                        else:
                            planned_details.append(
                                f"{planned_distance_number:g} km"
                            )

                except (TypeError, ValueError):
                    pass

            if pd.notna(planned_time):
                planned_time_text = str(
                    planned_time
                ).strip()

                if planned_time_text:
                    planned_details.append(
                        f"{planned_time_text} min"
                    )

            planned_details_display = " · ".join(
                planned_details
            )

                        # Prepare safe display text
            safe_session_name = html.escape(
                str(planned_session_display)
            )

            safe_details = html.escape(
                str(planned_details_display)
            )

            notes_display = ""

            if pd.notna(planned_notes):
                notes_text = str(planned_notes).strip()

                if notes_text and notes_text.lower() != "nan":
                    notes_display = html.escape(notes_text)

            # Build the optional distance/time line
            details_section = ""

            if safe_details:
                details_section = (
                    '<div style="'
                    'font-size:13px;'
                    'margin-top:4px;'
                    'font-weight:600;'
                    'color:#4A4200;'
                    '">'
                    f'{safe_details}'
                    '</div>'
                )

            # Build the optional notes line
            notes_section = ""

            if notes_display:
                notes_section = (
                    '<div style="'
                    'font-size:12px;'
                    'margin-top:6px;'
                    'color:#665500;'
                    'line-height:1.4;'
                    '">'
                    f'{notes_display}'
                    '</div>'
                )

                        # Planned card colours by sport
            planned_border_colours = {
                "Run": "#9ACD32",
                "Bike": "#F4A261",
                "Swim": "#6EC5FF",
                "Strength": "#B39DDB",
                "Other": "#F2C94C"
            }

            planned_background_colours = {
                "Run": "#F2FBE8",
                "Bike": "#FFF2E6",
                "Swim": "#EAF7FF",
                "Strength": "#F3EDFF",
                "Other": "#FFF8D9"
            }

            planned_border_colour = (
                planned_border_colours.get(
                    planned_sport_category,
                    "#F2C94C"
                )
            )

            planned_background_colour = (
                planned_background_colours.get(
                    planned_sport_category,
                    "#FFF8D9"
                )
            )    

            if planned_status == "MISSED":
                status_colour = "#C62828"
                planned_border_colour = "#D9534F"
                planned_background_colour = "#FDECEC"
            else:
                status_colour = planned_border_colour

            planned_card_html = (
                '<div style="'
                f'border:1px dashed {planned_border_colour};'
                f'border-left:5px solid {planned_border_colour};'
                f'background-color:{planned_background_colour};'
                'padding:10px;'
                'margin:8px 0;'
                'border-radius:5px;'
                'min-height:85px;'
                '">'
                '<div style="'
                'font-size:11px;'
                'font-weight:700;'
                f'color:{status_colour};'
                'margin-bottom:4px;'
                '">'
                f'{planned_status}'
                '</div>'
                '<div style="font-size:19px;">'
                f'{planned_icon}'
                '</div>'
                '<div style="font-weight:600;">'
                f'{safe_session_name}'
                '</div>'
                f'{details_section}'
                f'{notes_section}'
                '</div>'
            )

            st.markdown(
                planned_card_html,
                unsafe_allow_html=True
            )

conn.close()