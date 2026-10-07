import subprocess
import sys
from pathlib import Path

scripts_to_run = [
    "read_garmin.py",
    "lap_analysis.py",
    "tempo_work_block_analysis.py",
    "trend_analysis.py",
]

project_root = Path(__file__).resolve().parent.parent
scripts_folder = project_root / "scripts"

print("========================================")
print("GARMIN TRAINING ANALYSIS - RUN ALL")
print("========================================")
print(f"Project folder: {project_root}")
print("")

for script_name in scripts_to_run:
    script_path = scripts_folder / script_name

    print("----------------------------------------")
    print(f"Running: {script_name}")
    print("----------------------------------------")

    if not script_path.exists():
        print(f"ERROR: Could not find {script_path}")
        print("Stopping run.")
        sys.exit(1)

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=project_root
    )

    if result.returncode != 0:
        print("")
        print(f"ERROR: {script_name} failed.")
        print("Stopping run so you can fix the issue above.")
        sys.exit(result.returncode)

    print(f"Completed: {script_name}")
    print("")

print("========================================")
print("ALL SCRIPTS COMPLETED SUCCESSFULLY")
print("========================================")
print("Updated files should now be in the output folder:")
print("- session_summary.xlsx")
print("- lap_summary.xlsx")
print("- tempo_work_block_summary.xlsx")
print("- trend_summary.xlsx")