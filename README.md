# AI Teaching — Booking

Self-serve lesson booking page for Anzon Wong's 1-on-1 AI Teaching students.
Live: https://cowballwong.github.io/teaching-booking/

## How it works
- `build_slots.py` — calendar-aware slot generator. Opens the 4 Saturday slots
  (UK 07:30 / 08:45 / 10:00 / 11:15) for the next ~2 months, **minus** slots
  already booked in the Excel master, **minus** days covered by a blocking
  "UK Family" calendar event (all-day / multi-day = away → block; same-day timed
  = flag for Anzon, held back). Writes `docs/open-slots.json` containing **only**
  open date/time — never calendar event names or student data.
- `docs/index.html` — the public page. Student picks an open slot + name + email +
  new-vs-existing lesson → writes a `pending` doc to Firestore `teach_bookings`
  (reuses the `worldcup-bet-2026` Firebase project). The page can never **read**
  bookings back — Firestore rules forbid it — so no student PII is exposed.
- LillyRose drains `teach_bookings` server-side (Admin SDK), syncs into
  `students/attendance.xlsx` + the "AI Teaching" Google Calendar (Meet link),
  emails the student, then re-runs `build_slots.py` to refresh availability.

## Rebuild availability
```
py -3.13 build_slots.py        # regenerates docs/open-slots.json
git add docs/open-slots.json && git commit -m "refresh slots" && git push
```

Mirror of source lives at
`G:/My Drive/AI_Development/02_freelance/03_ai-teaching/students/booking/`.
