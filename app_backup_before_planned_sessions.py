import sqlite3
from datetime import timedelta
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

weekly_duration_seconds = (
    week_activities["Duration Seconds"].sum()
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

    with calendar_columns[day_number]:
        st.markdown(
            f"**{day_names[day_number]}**  \n"
            f"{calendar_date.strftime('%d %b')}"
        )

        if day_activities.empty:
            st.caption("Rest day")
        else:
            for _, activity in day_activities.iterrows():
                sport_category = calendar_sport_category(
                    activity["Sport"]
                )

                icon = sport_icons[sport_category]
                colour = sport_colours[sport_category]

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
                        background-color: rgba(128, 128, 128, 0.08);
                        padding: 10px;
                        margin: 8px 0;
                        border-radius: 5px;
                        min-height: 85px;
                    ">
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