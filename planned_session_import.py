import sqlite3
import pandas as pd

PLAN_FILE = "data/training_plan.xlsx"
DATABASE_FILE = "output/training.db"

print("Loading training plan...")

plan_df = pd.read_excel(
    PLAN_FILE,
    sheet_name=0
)

conn = sqlite3.connect(DATABASE_FILE)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS planned_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    planned_date TEXT,
    session_name TEXT,
    sport TEXT,
    planned_distance_km REAL,
    planned_time_min TEXT,
    notes TEXT
)
""")

cursor.execute(
    "DELETE FROM planned_sessions"
)

rows_imported = 0

for _, row in plan_df.iterrows():

    cursor.execute("""
    INSERT INTO planned_sessions (
        planned_date,
        session_name,
        sport,
        planned_distance_km,
        planned_time_min,
        notes
    )
    VALUES (?, ?, ?, ?, ?, ?)
    """, (
        str(row["Date"]),
        row["Session"],
        row["Sport"],
        row["Planned Distance (km)"],
        str(row["Planned Time (min)"])
        if not pd.isna(row["Planned Time (min)"])
        else None,
        row["Notes"]
    ))

    rows_imported += 1

conn.commit()

print()
print("=" * 40)
print("PLANNED SESSIONS IMPORT COMPLETE")
print("=" * 40)
print(f"Rows imported: {rows_imported}")

conn.close()