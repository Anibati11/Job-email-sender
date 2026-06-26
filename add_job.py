#!/usr/bin/env python3
"""
add_job.py
──────────
Paste a job description into input_jd.txt, then run this script.
It saves the JD as a timestamped file in the jobs/ queue folder
and clears input_jd.txt for the next paste.

Usage:
  1. Paste JD into input_jd.txt
  2. python add_job.py
  3. Repeat for each job
  4. python email_script.py   ← processes all queued jobs
"""

import datetime
from pathlib import Path

INPUT = Path("input_jd.txt")
JOBS  = Path("jobs")

JOBS.mkdir(exist_ok=True)

if not INPUT.exists():
    INPUT.write_text("", encoding="utf-8")
    print("Created input_jd.txt — paste a job description into it and re-run.")
else:
    raw = INPUT.read_text(encoding="utf-8").strip()
    if not raw:
        print("input_jd.txt is empty — paste a job description first.")
    else:
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out   = JOBS / f"job_{stamp}.txt"
        out.write_text(raw, encoding="utf-8")
        INPUT.write_text("", encoding="utf-8")
        queued = len(list(JOBS.glob("*.txt")))
        print(f"✅ Job saved → {out.name}  ({queued} job(s) queued)")
