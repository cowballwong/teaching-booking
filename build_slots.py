"""build_slots.py — generate the PUBLIC open-slots.json for the booking page.

Calendar-aware: opens the 4 Saturday slots (UK 07:30/08:45/10:00/11:15) for the
next ~2 months, MINUS (a) slots already booked in the Excel master, (b) days
covered by a BLOCKING UK Family event. Blocking rule (learned from Anzon):
  • all-day OR multi-day (spans >1 calendar day) event -> BLOCK the whole day
    (trips / annual leave / school camping — family genuinely away)
  • same-day timed event (a 2-hour birthday party etc.) -> do NOT block; collect
    as a FLAG for Anzon to decide (not every calendar entry means he's busy)

⚠️ PRIVACY: open-slots.json contains ONLY open date/time — never event names,
student names or emails. Flags (with event names) print to console / a private
log for Anzon only, never to the public json.
"""
import sys, json, datetime as dt
sys.stdout.reconfigure(encoding="utf-8")
from zoneinfo import ZoneInfo
from pathlib import Path
LON = ZoneInfo("Europe/London")
sys.path.insert(0, r"G:/My Drive/AI_Development/06_shared/03_calendar/scripts")
import openpyxl
from calendar_api import CalendarAPI

WEEKS = 9                                   # ~2 months of Saturdays
SLOTS_UK = ["07:30", "08:45", "10:00", "11:15"]
OUT = Path(__file__).parent / "docs" / "open-slots.json"
XLSX = r"G:/My Drive/AI_Development/02_freelance/03_ai-teaching/students/attendance.xlsx"

def hk(uk):
    h, m = map(int, uk.split(":")); return f"{(h+7)%24:02d}:{m:02d}"

def today():
    # plain script (not a workflow) — real date is fine here
    return dt.datetime.now(LON).date()

def saturdays(start, n):
    d = start
    while d.weekday() != 5:
        d += dt.timedelta(days=1)
    out = []
    for _ in range(n):
        out.append(d); d += dt.timedelta(days=7)
    return out

def load_blocking_and_flags(start, end):
    api = CalendarAPI()
    ukf = api.get_calendar_ids().get("UK Family")
    evs = api.get_events(start.isoformat(), end.isoformat(), calendar_id=ukf) if ukf else []
    blocking, flags = [], []
    for e in evs:
        s = e.get("start", {}); en = e.get("end", {}); summ = e.get("summary", "(no title)")
        if "dateTime" in s:
            sd = dt.datetime.fromisoformat(s["dateTime"]).astimezone(LON)
            ed = dt.datetime.fromisoformat(en["dateTime"]).astimezone(LON)
            if ed.date() > sd.date():
                blocking.append((sd.date(), ed.date(), summ, "multi-day"))
            else:
                flags.append((sd, ed, summ))
        elif "date" in s:
            sd = dt.date.fromisoformat(s["date"]); ed = dt.date.fromisoformat(en["date"])
            blocking.append((sd, ed - dt.timedelta(days=1), summ, "all-day"))
    return blocking, flags

def day_block(day, blocking):
    for sd, ed, summ, kind in blocking:
        if sd <= day <= ed:
            return f"{summ} ({kind})"
    return None

def slot_flags(day, uk, flags):
    h, m = map(int, uk.split(":")); ls = dt.datetime(day.year, day.month, day.day, h, m, tzinfo=LON); le = ls + dt.timedelta(hours=1)
    return [f"{day} {uk} ↔ {summ}" for sd, ed, summ in flags if ls < ed and sd < le]

def load_taken():
    wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True); ws = wb["Schedule"]
    rows = list(ws.iter_rows(values_only=True))
    hidx = next(i for i, r in enumerate(rows[:8]) if r and any(c and "student" in str(c).lower() for c in r) and any(c and "date" in str(c).lower() for c in r))
    hdr = rows[hidx]; col = lambda n: next((i for i, h in enumerate(hdr) if h and n in str(h).lower()), None)
    ci = {k: col(k) for k in ["date", "time", "status"]}
    taken = set()
    for r in rows[hidx + 1:]:
        st = str(r[ci['status']]) if ci['status'] is not None and r[ci['status']] else ""
        if not st.lower().startswith("schedul"): continue
        if not r[ci['date']] or not r[ci['time']]: continue
        taken.add((str(r[ci['date']])[:10], str(r[ci['time']])[:5]))
    return taken

def main():
    start = today()
    sats = saturdays(start, WEEKS)
    end = sats[-1] + dt.timedelta(days=1)
    blocking, flags = load_blocking_and_flags(start, end)
    taken = load_taken()
    open_slots, all_flags = [], []
    for day in sats:
        if day < start:  # never offer a past day
            continue
        if day_block(day, blocking):
            continue
        for uk in SLOTS_UK:
            if (day.isoformat(), uk) in taken:
                continue
            fl = slot_flags(day, uk, flags)
            if fl:
                all_flags += fl
                continue  # has a same-day event overlap -> hold back, let Anzon decide
            open_slots.append({"date": day.isoformat(), "weekday": "Sat",
                               "uk": uk, "hk": hk(uk)})
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({
        "generated": dt.datetime.now(LON).isoformat(timespec="minutes"),
        "tz_note": "UK = Europe/London; HK = UK+7 (BST)",
        "slots": open_slots,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ wrote {len(open_slots)} open slots over {WEEKS} Saturdays -> {OUT}")
    if all_flags:
        print("\n🟡 FLAGS (same-day timed overlaps — held back, your call):")
        for f in all_flags: print("  - " + f)

if __name__ == "__main__":
    main()
