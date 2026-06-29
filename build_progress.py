"""build_progress.py — emit student-progress.json for the booking page's
"returning student" lookup.

Anzon wants 舊生 to enter their email and see their class status, then book the
next lesson. To do that on a STATIC public page WITHOUT exposing the student
roster, this file is keyed by the **SHA-256 hash of the lowercased email** and
the value contains ONLY lesson counts — never a name, email, contact or date.

Worst case if the file is public: someone who already knows a student's exact
email learns which lesson number that student is on. They can NOT enumerate or
download the student list (hashes aren't reversible; no PII in the file).

Completion is inferred by DATE (a lesson dated before today = already taken),
because the Schedule sheet only carries Scheduled/Rescheduled statuses.
"""
import sys, json, hashlib, datetime as dt
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path
from zoneinfo import ZoneInfo
import openpyxl

LON = ZoneInfo("Europe/London")
XLSX = Path(r"G:/My Drive/AI_Development/02_freelance/03_ai-teaching/students/attendance.xlsx")
OUT = Path(__file__).parent / "docs" / "student-progress.json"
TOTAL = 5
DEAD = {"cancelled", "canceled", "dropped", "void"}

def h(email):
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()

def main():
    today = dt.datetime.now(LON).date()
    wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
    # email map (Students: header row 3, data from row 4; email in col 2 'Contact')
    email_of = {}
    for r in wb["Students"].iter_rows(min_row=4, values_only=True):
        if r and r[0] and len(r) > 1 and r[1] and "@" in str(r[1]):
            email_of[str(r[0]).strip()] = str(r[1]).strip().lower()
    # gather lessons per student (Schedule: data from row 5)
    per = {}   # name -> {"done": set, "all": set}
    for r in wb["Schedule"].iter_rows(min_row=5, values_only=True):
        if not r or not r[0] or not r[1]:
            continue
        name = str(r[0]).strip()
        status = str(r[4]).strip().lower() if len(r) > 4 and r[4] else ""
        if status in DEAD:
            continue
        try:
            lesson = int(r[1])
        except (TypeError, ValueError):
            continue
        datestr = str(r[2])[:10] if len(r) > 2 and r[2] else ""
        d = per.setdefault(name, {"done": set(), "all": set()})
        d["all"].add(lesson)
        try:
            if datestr and dt.date.fromisoformat(datestr) < today:
                d["done"].add(lesson)
        except ValueError:
            pass
    wb.close()

    out = {}
    for name, d in per.items():
        email = email_of.get(name)
        if not email:
            continue
        done = len(d["done"])
        # next = first lesson not yet taken (handles reschedules: a cancelled
        # future lesson is still "not done", so it surfaces as the one to rebook)
        nxt = next((n for n in range(1, TOTAL + 1) if n not in d["done"]), TOTAL)
        finished = done >= TOTAL
        out[h(email)] = {"done": done, "next": nxt,
                         "total": TOTAL, "finished": finished}

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({
        "generated": dt.datetime.now(LON).isoformat(timespec="minutes"),
        "note": "keyed by sha256(lowercased email); values are lesson counts only",
        "students": out,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ wrote progress for {len(out)} students (hashed) -> {OUT}")

if __name__ == "__main__":
    main()
