import sqlite3

import pandas as pd


PLAN_FILE = "data/training_plan.xlsx"
DATABASE_FILE = "output/training.db"


print("Loading training plan...")

plan_df = pd.read_excel(
    PLAN_FILE,
    sheet_name=0
)


# --------------------------------------------------
# Ensure optional analysis columns exist
# --------------------------------------------------

if "Drift Segment" not in plan_df.columns:
    plan_df["Drift Segment"] = None

if "Work Segments" not in plan_df.columns:
    plan_df["Work Segments"] = None


# --------------------------------------------------
# Database connection
# --------------------------------------------------

conn = sqlite3.connect(DATABASE_FILE)
cursor = conn.cursor()


# --------------------------------------------------
# Planned sessions table
# --------------------------------------------------

cursor.execute("""
CREATE TABLE IF NOT EXISTS planned_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    planned_date TEXT,
    session_name TEXT,
    sport TEXT,
    planned_distance_km REAL,
    planned_time_min TEXT,
    notes TEXT,
    drift_segment TEXT,
    work_segments TEXT
)
""")


# --------------------------------------------------
# Add new columns to an existing older database
# --------------------------------------------------

existing_columns = {
    row[1]
    for row in cursor.execute(
        "PRAGMA table_info(planned_sessions)"
    ).fetchall()
}

if "drift_segment" not in existing_columns:
    cursor.execute("""
    ALTER TABLE planned_sessions
    ADD COLUMN drift_segment TEXT
    """)

if "work_segments" not in existing_columns:
    cursor.execute("""
    ALTER TABLE planned_sessions
    ADD COLUMN work_segments TEXT
    """)


# --------------------------------------------------
# Clear old planned sessions
# --------------------------------------------------

cursor.execute(
    "DELETE FROM planned_sessions"
)


# --------------------------------------------------
# Import training plan
# --------------------------------------------------

rows_imported = 0

for _, row in plan_df.iterrows():

    # ----------------------------------------------
    # Date
    # ----------------------------------------------

    parsed_date = pd.to_datetime(
        row["Date"],
        errors="coerce"
    )

    if pd.isna(parsed_date):
        print(
            "Skipping row with invalid date:",
            row["Date"]
        )
        continue

    planned_date = str(
        parsed_date
    )[:10]


    # ----------------------------------------------
    # Optional analysis instructions
    # ----------------------------------------------

    drift_segment = None

    if pd.notna(
        row["Drift Segment"]
    ):
        drift_text = str(
            row["Drift Segment"]
        ).strip()

        if (
            drift_text
            and drift_text.lower() != "nan"
        ):
            drift_segment = drift_text


    work_segments = None

    if pd.notna(
        row["Work Segments"]
    ):
        work_text = str(
            row["Work Segments"]
        ).strip()

        if (
            work_text
            and work_text.lower() != "nan"
        ):
            work_segments = work_text


    # ----------------------------------------------
    # Other fields
    # ----------------------------------------------

    planned_distance = None

    if pd.notna(
        row["Planned Distance (km)"]
    ):
        planned_distance = row[
            "Planned Distance (km)"
        ]


    planned_time = None

    if pd.notna(
        row["Planned Time (min)"]
    ):
        planned_time = str(
            row["Planned Time (min)"]
        ).strip()


    notes = None

    if pd.notna(
        row["Notes"]
    ):
        notes_text = str(
            row["Notes"]
        ).strip()

        if (
            notes_text
            and notes_text.lower() != "nan"
        ):
            notes = notes_text


    # ----------------------------------------------
    # Insert row
    # ----------------------------------------------

    cursor.execute("""
    INSERT INTO planned_sessions (
        planned_date,
        session_name,
        sport,
        planned_distance_km,
        planned_time_min,
        notes,
        drift_segment,
        work_segments
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        planned_date,
        row["Session"],
        row["Sport"],
        planned_distance,
        planned_time,
        notes,
        drift_segment,
        work_segments
    ))

    rows_imported += 1


# --------------------------------------------------
# Save changes
# --------------------------------------------------

conn.commit()


print()
print("=" * 40)
print("PLANNED SESSIONS IMPORT COMPLETE")
print("=" * 40)
print(f"Rows imported: {rows_imported}")


conn.close()