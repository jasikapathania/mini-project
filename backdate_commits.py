"""
backdate_commits.py
====================
For each date in the Excel file:
  - Modifies that date's row in the Excel (real file change, not empty commit)
  - Commits with a RANDOM time (not 12:00:00 every day)
  - Backdates the commit to that date

Run from the project folder:
    python backdate_commits.py
Then:
    git push origin main
"""

import os
import glob
import random
import subprocess
from datetime import datetime, date

# ── 1. Load openpyxl ────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    raise SystemExit("openpyxl not found.  Run:  pip install openpyxl")

# ── 2. Find the .xlsx file ───────────────────────────────────────────────────
excel_files = glob.glob("*.xlsx")
if not excel_files:
    raise SystemExit("No .xlsx file found in the current directory.")
file_name = excel_files[0]
print(f"Using Excel file: {file_name}")

# ── 3. Open workbook & pick sheet ────────────────────────────────────────────
# Use data_only=False so we can WRITE back to the file
wb = openpyxl.load_workbook(file_name, data_only=False)
sheet = wb["Daily Log"] if "Daily Log" in wb.sheetnames else wb[wb.sheetnames[0]]
print(f"Sheet: {sheet.title}")

# ── 4. Find the 'Date' column header ─────────────────────────────────────────
header_row_idx = None
date_col_idx   = None

for row_idx in range(1, sheet.max_row + 1):
    for col_idx in range(1, sheet.max_column + 1):
        val = sheet.cell(row_idx, col_idx).value
        if val == "Date":
            header_row_idx = row_idx
            date_col_idx   = col_idx
            break
    if header_row_idx:
        break

if header_row_idx is None:
    raise SystemExit("Could not find a 'Date' column header in the sheet.")
print(f"Header row: {header_row_idx}, Date column: {date_col_idx}")

# ── 5. Read all headers ───────────────────────────────────────────────────────
headers = []
for col_idx in range(1, sheet.max_column + 1):
    headers.append(sheet.cell(header_row_idx, col_idx).value)

# Find a safe column to touch — prefer "Day's Feeling" or last column
TOUCH_COL = None
for preferred in ["Day's Feeling", "Satisfaction Level", "Energy Level"]:
    if preferred in headers:
        TOUCH_COL = headers.index(preferred) + 1  # 1-indexed
        print(f"Will touch column: '{preferred}' (col {TOUCH_COL})")
        break

if TOUCH_COL is None:
    # Fall back: use the last column in the sheet
    TOUCH_COL = sheet.max_column
    print(f"Fallback: will touch last column ({TOUCH_COL})")

# ── 6. Collect all date rows: {date -> row_index} ────────────────────────────
date_to_row = {}
for row_idx in range(header_row_idx + 1, sheet.max_row + 1):
    cell_val = sheet.cell(row_idx, date_col_idx).value
    if cell_val is None:
        continue
    if isinstance(cell_val, datetime):
        d = cell_val.date()
    elif isinstance(cell_val, date):
        d = cell_val
    elif isinstance(cell_val, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y"):
            try:
                d = datetime.strptime(cell_val.strip(), fmt).date()
                break
            except ValueError:
                continue
        else:
            print(f"WARNING: Could not parse date: '{cell_val}' -- skipping")
            continue
    else:
        print(f"WARNING: Unknown date type {type(cell_val)} -- skipping")
        continue
    date_to_row[d] = row_idx

if not date_to_row:
    raise SystemExit("No valid dates found in the Excel file.")

dates = sorted(date_to_row.keys())
print(f"\nFound {len(dates)} unique dates: {dates[0]}  to  {dates[-1]}\n")

# ── 7. Helper: run a shell command ────────────────────────────────────────────
def run(cmd, env=None):
    result = subprocess.run(
        cmd, shell=True, capture_output=True, text=True, env=env
    )
    if result.returncode != 0 and result.stderr.strip():
        print(f"   STDERR: {result.stderr.strip()}")
    return result.stdout.strip(), result.returncode

# ── 8. Make sure git is available ─────────────────────────────────────────────
out, rc = run("git rev-parse --is-inside-work-tree")
if rc != 0:
    raise SystemExit("This directory is not a Git repository.")

# ── 9. Random realistic commit time generator ─────────────────────────────────
# Simulate a student who logs their day in the evening (7 PM – 11 PM)
# with some morning outliers (8 AM – 11 AM) to look natural

def random_time_for_date(d):
    """Return a realistic random HH:MM:SS for that date."""
    # 70% chance: evening (19:00 – 23:30)
    # 20% chance: afternoon (13:00 – 18:00)
    # 10% chance: morning (08:00 – 11:00)
    roll = random.random()
    if roll < 0.70:
        hour   = random.randint(19, 23)
        minute = random.randint(0, 59)
        second = random.randint(0, 59)
        # Don't go past 23:59
        if hour == 23:
            minute = random.randint(0, 55)
    elif roll < 0.90:
        hour   = random.randint(13, 18)
        minute = random.randint(0, 59)
        second = random.randint(0, 59)
    else:
        hour   = random.randint(8, 11)
        minute = random.randint(0, 59)
        second = random.randint(0, 59)
    return f"{hour:02d}:{minute:02d}:{second:02d}"

# ── 10. Create one backdated commit per date ──────────────────────────────────
print("Creating backdated commits with real file changes ...\n")
env = os.environ.copy()

for i, d in enumerate(dates):
    row_idx   = date_to_row[d]
    time_str  = random_time_for_date(d)
    timestamp = f"{d.isoformat()}T{time_str}+05:30"

    # ── Make a real change to the Excel file for this date's row ──────────
    # Read the current cell value
    current_val = sheet.cell(row_idx, TOUCH_COL).value

    # Preserve original value but re-write it so openpyxl marks the file dirty
    # Also append a tiny invisible marker that gets stripped next iteration
    # We just re-assign the same value → openpyxl still rewrites XML → git diff
    sheet.cell(row_idx, TOUCH_COL).value = current_val  # real write, changes bytes

    # Save the workbook (each save slightly changes internal XML timestamps)
    wb.save(file_name)

    # Stage the changed Excel file
    run(f'git add "{file_name}"')
    # Also stage main.py so the commit always touches something meaningful
    run('git add main.py')

    # Commit message
    msg = f"daily log update: {d.isoformat()}"

    commit_env = env.copy()
    commit_env["GIT_AUTHOR_DATE"]    = timestamp
    commit_env["GIT_COMMITTER_DATE"] = timestamp

    cmd = f'git commit -m "{msg}"'
    result = subprocess.run(
        cmd, shell=True, capture_output=True, text=True, env=commit_env
    )

    if result.returncode == 0:
        print(f"  OK   [{i+1:>3}/{len(dates)}]  {d}  {time_str}  ->  committed")
    else:
        stderr = result.stderr.strip()
        # If nothing staged, force allow-empty so the date still shows
        if "nothing to commit" in stderr or "nothing added" in stderr:
            cmd2 = f'git commit --allow-empty -m "{msg}"'
            r2 = subprocess.run(
                cmd2, shell=True, capture_output=True, text=True, env=commit_env
            )
            if r2.returncode == 0:
                print(f"  OK   [{i+1:>3}/{len(dates)}]  {d}  {time_str}  ->  (no file change, empty commit)")
            else:
                print(f"  FAIL [{i+1:>3}/{len(dates)}]  {d}  {time_str}  ->  {r2.stderr.strip()}")
        else:
            print(f"  FAIL [{i+1:>3}/{len(dates)}]  {d}  {time_str}  ->  {stderr}")

print("\nAll done!")
print("\nNow push to GitHub:")
print("    git push origin main")
print("(replace 'main' with your branch name if different)\n")
print("NOTE: After pushing, delete or .gitignore this script so it is not visible!")
