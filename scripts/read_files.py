from fitparse import FitFile
import os
from openpyxl import Workbook

data_folder = "data"
output_folder = "output"
output_file = os.path.join(output_folder, "session_summary.xlsx")

files = os.listdir(data_folder)
fit_files = [file for file in files if file.endswith(".fit")]

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

        # Determine session type from file name
        name_lower = session_name.lower()

        if "brick_run" in name_lower:
            session_type = "Brick Run"
        elif "easy_run" in name_lower:
            session_type = "Easy Run"
        elif "interval_run" in name_lower:
            session_type = "Interval Run"
        elif "long_run" in name_lower:
            session_type = "Long Run"
        elif "park_run" in name_lower:
            session_type = "Park Run"
        elif "recovery_run" in name_lower:
            session_type = "Recovery Run"
        elif "race_day" in name_lower:
            session_type = "Race Day"
        elif "tempo_run" in name_lower:
            session_type = "Tempo Run"
        elif "threshold_1" in name_lower:
            session_type = "Threshold 1"
        elif "threshold_2" in name_lower:
            session_type = "Threshold 2"
        else:
            session_type = "Other"
        
        for session in fitfile.get_messages("session"):
            for data in session:
        
        # Determine environment
        sub_sport_text = str(sub_sport).lower() if sub_sport is not None else ""

        if "treadmill" in sub_sport_text:
            environment = "Treadmill"
        elif "indoor" in sub_sport_text:
            environment = "Indoor"
        elif "treadmill" in name_lower:
            environment = "Treadmill"
        else:
            environment = "Outdoor"

        fitfile = FitFile(file_path)

        # Session-level summary fields
        total_distance_m = None
        total_timer_time_s = None
        avg_heart_rate_bpm = None
        avg_speed_mps = None
        total_ascent_m = None
        total_calories = None
        sport = None
        sub_sport = None

        for session in fitfile.get_messages("session"):
            for data in session:
                if data.name == "total_distance":
                    total_distance_m = data.value
                elif data.name == "total_timer_time":
                    total_timer_time_s = data.value
                elif data.name == "avg_heart_rate":
                    avg_heart_rate_bpm = data.value
                elif data.name == "enhanced_avg_speed":
                    avg_speed_mps = data.value
                elif data.name == "total_ascent":
                    total_ascent_m = data.value
                elif data.name == "total_calories":
                    total_calories = data.value
                elif data.name == "sport":
                    sport = data.value
                elif data.name == "sub_sport":
                    sub_sport = data.value

        # Distance in km
        distance_km = round(total_distance_m / 1000, 2) if total_distance_m is not None else None

        # Time in hh:mm:ss
        time_hms = None
        if total_timer_time_s is not None:
            total_seconds = int(total_timer_time_s)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            time_hms = f"{hours}:{minutes:02d}:{seconds:02d}"
    
        # Pace from speed
        avg_pace = None
        if avg_speed_mps is not None and avg_speed_mps > 0:
            avg_pace_sec_per_km = 1000 / avg_speed_mps
            pace_min = int(avg_pace_sec_per_km // 60)
            pace_sec = int(avg_pace_sec_per_km % 60)
            avg_pace = f"{pace_min}:{pace_sec:02d}"
    
        # Print summary
        print(f"Date: {date}")
        print(f"Session Name: {session_name}")
        print(f"Session Type: {session_type}")
        print(f"Sport: {sport}")
        print(f"Sub Sport: {sub_sport}")
        print(f"Environment: {environment}")
        print(f"Distance: {distance_km} km")
        print(f"Time: {time_hms}")
        print(f"Average Heart Rate: {avg_heart_rate_bpm} bpm")
        print(f"Average Speed: {avg_speed_mps:.3f} m/s" if avg_speed_mps is not None else "Average Speed: None")
        print(f"Average Pace: {avg_pace} min/km")
        print(f"Total Ascent: {total_ascent_m} m")
        print(f"Calories: {total_calories}")

        # Empty manual fields for now
        notes = ""
        rpe = ""

        rows.append([date, session_name, session_type, sport, sub_sport, environment, distance_km, time_hms, avg_heart_rate_bpm, round(avg_speed_mps, 3) if avg_speed_mps is not None else None, avg_pace, total_ascent_m, total_calories, notes, rpe])

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Session Summary"

    # Header row
    ws.append(["Date", "Session Name", "Session Type", "Sport", "Sub Sport", "Environment", "Distance (km)", "Time (hh:mm:ss)", "Average Heart Rate (bpm)", "Average Speed (m/s)", "Average Pace (min/km)", "Total Ascent (m)", "Calories", "Notes", "RPE"])

    # Data rows
    for row in rows:
        ws.append(row)

    # Optional: set column widths to make it easier to read
    column_widths = {"A": 14, "B": 18, "C": 18, "D": 12, "E": 14, "F": 14, "G": 14, "H": 16, "I": 24, "J": 20, "K": 22, "L": 18, "M": 12, "N": 30, "O": 8}

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    wb.save(output_file)

    print(f"\nSession summaries saved to: {output_file}")