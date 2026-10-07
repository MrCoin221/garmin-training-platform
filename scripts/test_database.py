import sqlite3

conn = sqlite3.connect("output/training.db")
cursor = conn.cursor()

print()
print("=" * 40)
print("DATABASE TEST")
print("=" * 40)

cursor.execute("SELECT COUNT(*) FROM activities")

activity_count = cursor.fetchone()[0]

print(f"Activities in database: {activity_count}")

print()
print("Latest 10 activities")
print("-" * 40)

cursor.execute("""
SELECT
    activity_date,
    session_name,
    sport,
    distance_km
FROM activities
ORDER BY activity_date DESC
LIMIT 10
""")

rows = cursor.fetchall()

for row in rows:
    print(row)

conn.close()