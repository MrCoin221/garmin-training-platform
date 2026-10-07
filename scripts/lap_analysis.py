from fitparse import FitFile
import os
from openpyxl import Workbook

data_folder = "data"
output_folder = "output"
output_file = os.path.join(output_folder, "lap_summary.xlsx")

files = os.listdir(data_folder)
fit_files = [file for file in files if file.endswith(".fit")]


def seconds_to_hms(total_seconds):
    """Convert seconds into hh:mm:ss text."""
    if total_seconds is None:
        return None

    total_seconds = int(total_seconds)

    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def speed_to_pace(speed_mps):
    """Convert speed in metres per second into min/km pace."""
    if speed_mps is None or speed_mps <= 0:
        return None

    seconds_per_km = 1000 / speed_mps
    minutes = int(seconds_per_km // 60)
    seconds = int(seconds_per_km % 60)

    return f"{minutes:02d}:{seconds:02d}"


def classify_session_type(session_name):
    """Determine session type from file name."""
    name_lower = session_name.lower()

    if "brick_run" in name_lower:
        return "Brick Run"
    elif "easy_run" in name_lower:
        return "Easy Run"
    elif "interval_run" in name_lower:
        return "Interval Run"
    elif "long_run" in name_lower:
        return "Long Run"
    elif "park_run" in name_lower:
        return "Park Run"
    elif "recovery_run" in name_lower:
        return "Recovery Run"
    elif "race_day" in name_lower:
        return "Race Day"
    elif "tempo_run" in name_lower:
        return "Tempo Run"
    elif "threshold_1" in name_lower:
        return "Threshold 1"
    elif "threshold_2" in name_lower:
        return "Threshold 2"
    else:
        return "Other"


if not fit_files:
    print("No .fit files found in the data folder.")

else:
    rows = []

    for file_name in fit_files:
        file_path = os.path.join(data_folder, file_name)
        print(f"\nReading file: {file_name}")

        # Split file name into date and session name
        base_name = file_name.replace(".fit", "")
        parts = base_name.split("_", 1)

        if len(parts) == 2:
            date = parts[0]
            session_name = parts[1]
        else:
            date = ""
            session_name = base_name

        session_type = classify_session_type(session_name)

        fitfile = FitFile(file_path)

        # Read session-level environment information
        sub_sport = None
        total_ascent_session = None

        for session in fitfile.get_messages("session"):
            for data in session:
                if data.name == "sub_sport":
                    sub_sport = data.value
                elif data.name == "total_ascent":
                    total_ascent_session = data.value

        sub_sport_text = str(sub_sport).lower() if sub_sport is not None else ""
        name_lower = session_name.lower()

        if "treadmill" in sub_sport_text:
            environment = "Treadmill"
        elif "indoor" in sub_sport_text:
            environment = "Indoor"
        elif "treadmill" in name_lower:
            environment = "Treadmill"
        elif total_ascent_session is None:
            environment = "Unknown"
        else:
            environment = "Outdoor"

        lap_number = 0

        for lap in fitfile.get_messages("lap"):
            lap_number += 1

            lap_distance_m = None
            lap_time_s = None
            lap_avg_hr = None
            lap_max_hr = None
            lap_avg_speed_mps = None
            lap_ascent_m = None

            for data in lap:
                if data.name == "total_distance":
                    lap_distance_m = data.value
                elif data.name == "total_timer_time":
                    lap_time_s = data.value
                elif data.name == "avg_heart_rate":
                    lap_avg_hr = data.value
                elif data.name == "max_heart_rate":
                    lap_max_hr = data.value
                elif data.name == "enhanced_avg_speed":
                    lap_avg_speed_mps = data.value
                elif data.name == "avg_speed" and lap_avg_speed_mps is None:
                    lap_avg_speed_mps = data.value
                elif data.name == "total_ascent":
                    lap_ascent_m = data.value

            lap_distance_km = round(lap_distance_m / 1000, 2) if lap_distance_m is not None else None
            lap_time_hms = seconds_to_hms(lap_time_s)
            lap_pace = speed_to_pace(lap_avg_speed_mps)

            rows.append([
                date,
                session_name,
                session_type,
                environment,
                lap_number,
                lap_distance_km,
                lap_time_hms,
                round(lap_avg_speed_mps, 3) if lap_avg_speed_mps is not None else None,
                lap_pace,
                lap_avg_hr,
                lap_max_hr,
                lap_ascent_m
            ])

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Lap Summary"

    ws.append([
        "Date",
        "Session Name",
        "Session Type",
        "Environment",
        "Lap Number",
        "Lap Distance (km)",
        "Lap Time (hh:mm:ss)",
        "Lap Avg Speed (m/s)",
        "Lap Pace (min/km)",
        "Lap Avg Heart Rate (bpm)",
        "Lap Max Heart Rate (bpm)",
        "Lap Ascent (m)"
    ])

    for row in rows:
        ws.append(row)

    column_widths = {
        "A": 14,
        "B": 24,
        "C": 18,
        "D": 14,
        "E": 12,
        "F": 18,
        "G": 20,
        "H": 20,
        "I": 20,
        "J": 26,
        "K": 26,
        "L": 16
    }

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    wb.save(output_file)

    print(f"\nLap summary saved to: {output_file}")