"""build_slots.py — generate the PUBLIC open-slots.json for the booking page.

Calendar-aware: opens the 5 Saturday slots (UK 06:15/07:30/08:45/10:00/11:15) for
the next ~4 months, MINUS (a) slots already booked in the Excel master, (b) days
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
HKG = ZoneInfo("Asia/Hong_Kong")
sys.path.insert(0, r"G:/My Drive/AI_Development/06_shared/07_browser/scripts")
import openpyxl

WEEKS = 16                                  # through end of Oct (Anzon, 2026-07-15)
SLOTS_UK = ["06:15", "07:30", "08:45", "10:00", "11:15"]
IRIS_MARKERS = {"al", "ld"}                 # Iris leave / work — never block teaching
OUT = Path(__file__).parent / "docs" / "open-slots.json"
XLSX = r"G:/My Drive/AI_Development/02_freelance/03_ai-teaching/students/attendance.xlsx"

def hk(uk, day):
    """UK local -> HK local for THAT date. Must be date-aware: HK is UK+7 during
    BST but UK+8 once the UK clocks go back (last Sun of Oct — 2026-10-25), so a
    hardcoded +7 silently shows HK students the wrong time for late-Oct slots."""
    h, m = map(int, uk.split(":"))
    return (dt.datetime(day.year, day.month, day.day, h, m, tzinfo=LON)
            .astimezone(HKG).strftime("%H:%M"))

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
    # Read availability from the AUTHORITATIVE iCloud "UK Family" calendar via
    # CalDAV — real-time, no Google-import sync lag. (AI Teaching LESSONS live on
    # the separate Google "AI Teaching" calendar; this is only Anzon's life cal.)
    import io, contextlib, icloud_caldav
    with contextlib.redirect_stdout(io.StringIO()):       # hush its console prints
        evs = icloud_caldav.list_events(start.isoformat(), end.isoformat(),
                                        calendar="UK Family", json_out=False)
    blocking, flags = [], []
    for e in evs:
        summ = e.get("summary", "(no title)")
        # "AL" (Iris annual leave) / "LD" (Iris working days) are IRIS's schedule —
        # they do NOT block Anzon's teaching. Only genuine whole-family trips (named
        # events) or an explicit cancel block a Saturday. (Anzon corrected 2026-06-29.)
        if summ.strip().lower() in IRIS_MARKERS:
            continue
        s = (e.get("start") or "").strip(); en = (e.get("end") or "").strip()
        if not s:
            continue
        if len(s) == 10:                          # all-day "YYYY-MM-DD" (DTEND exclusive)
            sd = dt.date.fromisoformat(s)
            ed = dt.date.fromisoformat(en[:10]) if en else sd + dt.timedelta(days=1)
            # Only MULTI-day all-day events (genuine whole-family trips) block.
            # Single-day all-day entries (Bank shift, kids' PE, AL/LD) are personal
            # markers, NOT family-away days — never block teaching. Genuine single-day
            # cancellations go through blocked_dates.json. (Anzon corrected 2026-06-29.)
            if (ed - sd).days >= 2:
                blocking.append((sd, ed - dt.timedelta(days=1), summ, "all-day"))
        else:                                     # timed "YYYY-MM-DD HH:MM:SS+TZ"
            sdt = dt.datetime.fromisoformat(s).astimezone(LON)
            edt = dt.datetime.fromisoformat(en).astimezone(LON) if en else sdt
            if edt.date() > sdt.date():
                blocking.append((sdt.date(), edt.date(), summ, "multi-day"))
            else:
                flags.append((sdt, edt, summ))
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
        st = str(r[ci['status']]).strip().lower() if ci['status'] is not None and r[ci['status']] else ""
        # a slot is taken by any BOOKED lesson — Scheduled OR Rescheduled (or Done).
        # (NOT just startswith "schedul", which wrongly skipped "Rescheduled".)
        if st not in ("scheduled", "rescheduled", "done"): continue
        if not r[ci['date']] or not r[ci['time']]: continue
        taken.add((str(r[ci['date']])[:10], str(r[ci['time']])[:5]))
    return taken

def load_manual_blocks():
    """Saturdays Anzon explicitly cancelled / is away (blocked_dates.json) — these
    stay closed regardless of calendar markers (AL/LD don't auto-block)."""
    try:
        data = json.loads(Path(XLSX).parent.joinpath("blocked_dates.json").read_text(encoding="utf-8"))
        return {b["date"] for b in data.get("blocked", []) if b.get("date")}
    except Exception:
        return set()

def main():
    start = today()
    sats = saturdays(start, WEEKS)
    end = sats[-1] + dt.timedelta(days=1)
    blocking, flags = load_blocking_and_flags(start, end)
    manual_blocked = load_manual_blocks()
    taken = load_taken()
    open_slots, all_flags = [], []
    for day in sats:
        if day < start:  # never offer a past day
            continue
        if day.isoformat() in manual_blocked:   # explicitly cancelled / away
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
                               "uk": uk, "hk": hk(uk, day)})
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({
        "generated": dt.datetime.now(LON).isoformat(timespec="minutes"),
        "tz_note": "UK = Europe/London; HK = Asia/Hong_Kong (UK+7 in BST, UK+8 in GMT) — computed per slot date",
        "slots": open_slots,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ wrote {len(open_slots)} open slots over {WEEKS} Saturdays -> {OUT}")
    if all_flags:
        print("\n🟡 FLAGS (same-day timed overlaps — held back, your call):")
        for f in all_flags: print("  - " + f)

if __name__ == "__main__":
    main()
