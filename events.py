"""Calendar helpers: time zones, event normalisation and rendering, free time.

All times shown to the agent are in the user's time zone (TIMEZONE, or the
mailbox's own setting); Graph returns UTC and we convert here.
"""
from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import formatting as fmt

# Exchange mailbox settings use Windows time-zone names. Common ones -> IANA.
WINDOWS_TZ = {
    "UTC": "UTC", "Coordinated Universal Time": "UTC",
    "GMT Standard Time": "Europe/London", "Greenwich Standard Time": "Atlantic/Reykjavik",
    "W. Europe Standard Time": "Europe/Berlin", "Romance Standard Time": "Europe/Paris",
    "Central Europe Standard Time": "Europe/Budapest", "Central European Standard Time": "Europe/Warsaw",
    "GTB Standard Time": "Europe/Bucharest", "FLE Standard Time": "Europe/Kiev",
    "E. Europe Standard Time": "Europe/Chisinau", "Turkey Standard Time": "Europe/Istanbul",
    "Russian Standard Time": "Europe/Moscow", "Egypt Standard Time": "Africa/Cairo",
    "South Africa Standard Time": "Africa/Johannesburg", "Israel Standard Time": "Asia/Jerusalem",
    "Jordan Standard Time": "Asia/Amman", "Arab Standard Time": "Asia/Riyadh",
    "Arabic Standard Time": "Asia/Baghdad", "Arabian Standard Time": "Asia/Dubai",
    "Iran Standard Time": "Asia/Tehran", "Pakistan Standard Time": "Asia/Karachi",
    "India Standard Time": "Asia/Kolkata", "Bangladesh Standard Time": "Asia/Dhaka",
    "SE Asia Standard Time": "Asia/Bangkok", "Singapore Standard Time": "Asia/Singapore",
    "China Standard Time": "Asia/Shanghai", "Tokyo Standard Time": "Asia/Tokyo",
    "Korea Standard Time": "Asia/Seoul", "AUS Eastern Standard Time": "Australia/Sydney",
    "New Zealand Standard Time": "Pacific/Auckland", "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago", "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles", "SA Pacific Standard Time": "America/Bogota",
    "E. South America Standard Time": "America/Sao_Paulo", "Morocco Standard Time": "Africa/Casablanca",
}
IANA_TO_WINDOWS = {v: k for k, v in WINDOWS_TZ.items() if k != "Coordinated Universal Time"}
IANA_TO_WINDOWS.update({"Etc/UTC": "UTC", "Asia/Muscat": "Arabian Standard Time"})

DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
SHOW_AS = ("free", "tentative", "busy", "oof", "workingElsewhere")


def zone(name: str | None) -> ZoneInfo | None:
    """A ZoneInfo for an IANA or (common) Windows time-zone name, else None."""
    if not name:
        return None
    try:
        return ZoneInfo(WINDOWS_TZ.get(name, name))
    except (ZoneInfoNotFoundError, ValueError):
        return None


def graph_tz_name(tz: ZoneInfo) -> str:
    """The name to send to Graph: Windows name when known (always accepted), else IANA."""
    return IANA_TO_WINDOWS.get(tz.key, tz.key)


def _naive(value: str) -> datetime:
    m = re.match(r"^(\d{4}-\d\d-\d\d)(?:[T ](\d\d:\d\d(?::\d\d)?))?", value or "")
    if not m:
        raise ValueError(value)
    return datetime.fromisoformat(m.group(1) + "T" + (m.group(2) or "00:00"))


def from_graph(dt: dict | None, tz: ZoneInfo) -> datetime | None:
    """Graph dateTimeTimeZone -> aware datetime in tz."""
    if not dt or not dt.get("dateTime"):
        return None
    src = zone(dt.get("timeZone") or "UTC") or timezone.utc
    return _naive(dt["dateTime"]).replace(tzinfo=src).astimezone(tz)


def to_graph(dt: datetime, tz: ZoneInfo) -> dict:
    return {"dateTime": dt.astimezone(tz).strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": graph_tz_name(tz)}


def is_date(value: str) -> bool:
    return bool(re.fullmatch(r"\s*\d{4}-\d\d-\d\d\s*", value or "")) or \
        (value or "").strip().lower() in ("today", "tomorrow")


def parse_local(value: str, tz: ZoneInfo, now: datetime) -> datetime:
    """'2026-11-18T09:00' (local), '...Z' / '+04:00' (converted), '2026-11-18'
    (local midnight), 'now', 'today', 'tomorrow'. Raises ValueError."""
    v = (value or "").strip()
    low = v.lower()
    local_now = now.astimezone(tz)
    if low == "now":
        return local_now.replace(second=0, microsecond=0)
    if low in ("today", "tomorrow"):
        d = local_now.date() + timedelta(days=1 if low == "tomorrow" else 0)
        return datetime.combine(d, time(0), tz)
    s = v.replace("Z", "+00:00").replace("z", "+00:00")
    s = re.sub(r"(\d\d:\d\d:\d\d)\.\d+", r"\1", s)
    if re.fullmatch(r"\d{4}-\d\d-\d\d", s):
        s += "T00:00"
    dt = datetime.fromisoformat(s)
    return dt.replace(tzinfo=tz) if dt.tzinfo is None else dt.astimezone(tz)


def day_label(d: date) -> str:
    return f"{d:%a} {d.isoformat()}"


def duration_text(minutes: int) -> str:
    h, m = divmod(int(minutes), 60)
    if h and m:
        return f"{h}h{m:02d}"
    return f"{h}h" if h else f"{m} min"


# ---------------------------------------------------------------- events
EVENT_FIELDS = ("id,subject,start,end,isAllDay,location,organizer,attendees,responseStatus,showAs,"
                "isOnlineMeeting,onlineMeeting,isCancelled,isOrganizer,type,seriesMasterId,sensitivity,"
                "importance,categories,bodyPreview,webLink,responseRequested,hasAttachments")
EVENT_FULL_FIELDS = EVENT_FIELDS + (",body,recurrence,reminderMinutesBeforeStart,isReminderOn,"
                                    "allowNewTimeProposals,locations")

RESPONSES = {"accepted": "accepted", "tentativelyAccepted": "tentative", "declined": "declined",
             "notResponded": "not responded", "none": "not responded", "organizer": "organiser"}


def event_times(e: dict, tz: ZoneInfo) -> tuple[datetime | date | None, datetime | date | None]:
    """(start, end) as local datetimes, or dates for all-day events (end exclusive)."""
    if e.get("isAllDay"):
        s, en = (e.get("start") or {}).get("dateTime"), (e.get("end") or {}).get("dateTime")
        return (_naive(s).date() if s else None, _naive(en).date() if en else None)
    return from_graph(e.get("start"), tz), from_graph(e.get("end"), tz)


def attendee(a: dict) -> dict:
    out = fmt.address(a)
    out["type"] = a.get("type") or "required"
    out["response"] = RESPONSES.get(((a.get("status") or {}).get("response")) or "none", "not responded")
    return out


def recurrence_text(r: dict | None) -> str | None:
    if not r:
        return None
    p, rng = r.get("pattern") or {}, r.get("range") or {}
    n = p.get("interval") or 1
    kind = p.get("type") or ""
    days = ", ".join(d.capitalize() for d in p.get("daysOfWeek") or [])
    if kind == "daily":
        text = "daily" if n == 1 else f"every {n} days"
    elif kind == "weekly":
        text = ("weekly" if n == 1 else f"every {n} weeks") + (f" on {days}" if days else "")
    elif kind == "absoluteMonthly":
        text = ("monthly" if n == 1 else f"every {n} months") + f" on day {p.get('dayOfMonth')}"
    elif kind == "relativeMonthly":
        text = f"monthly on the {p.get('index', '')} {days}"
    elif kind.endswith("Yearly"):
        text = "yearly"
    else:
        text = kind or "recurring"
    if rng.get("type") == "endDate" and rng.get("endDate"):
        text += f" until {rng['endDate']}"
    elif rng.get("type") == "numbered" and rng.get("numberOfOccurrences"):
        text += f", {rng['numberOfOccurrences']} times"
    return text


def event(e: dict, tz: ZoneInfo, body_chars: int = 0) -> dict:
    s, en = event_times(e, tz)
    all_day = bool(e.get("isAllDay"))
    om = e.get("onlineMeeting") or {}
    organizer = fmt.address(e.get("organizer"))
    attendees = [attendee(a) for a in e.get("attendees") or []]
    d = {
        "event_id": e.get("id"),
        "subject": e.get("subject") or "(no subject)",
        "all_day": all_day or None,
        "start": s.isoformat() if s else None,
        "end": en.isoformat() if en else None,
        "location": (e.get("location") or {}).get("displayName") or None,
        "online_meeting": (om.get("joinUrl") or True) if e.get("isOnlineMeeting") else None,
        "organizer": organizer or None,
        "is_organizer": bool(e.get("isOrganizer")),
        "my_response": RESPONSES.get((e.get("responseStatus") or {}).get("response") or "none"),
        "response_requested": e.get("responseRequested"),
        "show_as": e.get("showAs"),
        "cancelled": e.get("isCancelled") or None,
        "recurring": (e.get("type") in ("occurrence", "exception", "seriesMaster")) or None,
        "series_id": e.get("seriesMasterId") or (e.get("id") if e.get("type") == "seriesMaster" else None),
        "private": (e.get("sensitivity") in ("private", "confidential")) or None,
        "importance": e.get("importance") if e.get("importance") not in (None, "normal") else None,
        "categories": e.get("categories") or None,
        "attachments": e.get("hasAttachments") or None,
        "attendees": attendees or None,
        "preview": fmt.shorten(fmt.clean_links(e.get("bodyPreview") or ""), 160) or None,
        "web_link": e.get("webLink"),
    }
    if body_chars:
        text = fmt.body_text(e.get("body"))
        d["body"] = text[:body_chars] + ("…" if len(text) > body_chars else "")
        d["preview"] = None
        d["recurrence"] = recurrence_text(e.get("recurrence"))
        if e.get("isReminderOn") is False:
            d["reminder"] = "off"
        elif e.get("reminderMinutesBeforeStart") is not None:
            d["reminder"] = f"{e['reminderMinutesBeforeStart']} min before"
    return {k: v for k, v in d.items() if v is not None and v != ""}


def needs_response(d: dict) -> bool:
    return (not d.get("is_organizer") and not d.get("cancelled") and d.get("response_requested") is not False
            and d.get("my_response") == "not responded" and bool(d.get("organizer")))


def _marks(d: dict) -> str:
    marks = []
    if d.get("cancelled"):
        marks.append("CANCELLED")
    if needs_response(d):
        marks.append("NEEDS RESPONSE")
    elif d.get("is_organizer"):
        if d.get("attendees"):
            marks.append("you organise")
    elif d.get("my_response") in ("accepted", "tentative", "declined"):
        marks.append(d["my_response"])
    if d.get("show_as") not in (None, "busy", "unknown") and not (
            d["show_as"] == "tentative" and d.get("my_response") in ("tentative", "not responded")):
        marks.append(f"shown as {d['show_as']}")
    if d.get("recurring"):
        marks.append("🔁")
    if d.get("private"):
        marks.append("private")
    if d.get("importance"):
        marks.append(f"!{d['importance']}")
    if d.get("attachments"):
        marks.append("📎")
    if d.get("categories"):
        marks.append("[" + ", ".join(d["categories"]) + "]")
    return " · ".join(marks)


def _local(value: str) -> datetime | date:
    return date.fromisoformat(value) if len(value) == 10 else datetime.fromisoformat(value)


def when_text(d: dict, with_day: bool = True) -> str:
    s, e = _local(d["start"]), _local(d["end"]) if d.get("end") else None
    if d.get("all_day"):
        last = (e - timedelta(days=1)) if e else s
        text = (day_label(s) + " ") if with_day else ""
        return text + "all day" + (f" to {day_label(last)}" if last != s else "")
    text = (day_label(s.date()) + " ") if with_day else ""
    text += f"{s:%H:%M}"
    if e:
        text += f"–{e:%H:%M}" if e.date() == s.date() else f" → {day_label(e.date())} {e:%H:%M}"
    return text


def attendee_summary(atts: list[dict] | None) -> str | None:
    if not atts:
        return None
    counts: dict[str, int] = {}
    for a in atts:
        counts[a["response"]] = counts.get(a["response"], 0) + 1
    parts = [f"{n} {k}" for k, n in counts.items() if k != "organiser"]
    return f"{len(atts)} attendee{'s' if len(atts) != 1 else ''}" + (f" ({', '.join(parts)})" if parts else "")


def render_events(events: list[dict]) -> str:
    out: list[str] = []
    day = None
    for d in events:
        s = _local(d["start"])
        dd = s if isinstance(s, date) and not isinstance(s, datetime) else s.date()
        if dd != day:
            day = dd
            out.append(day_label(dd))
        where = []
        if d.get("location"):
            where.append(d["location"])
        if d.get("online_meeting"):
            where.append("Teams")
        marks = _marks(d)
        out.append(f"{when_text(d, with_day=False)} {d['subject']}"
                   + "".join(f" · {w}" for w in where) + (f" · {marks}" if marks else ""))
        who = []
        if d.get("organizer") and not d.get("is_organizer"):
            who.append("organiser " + fmt.addr_text(d["organizer"]))
        summ = attendee_summary(d.get("attendees"))
        if summ:
            who.append(summ)
        if who:
            out.append(fmt._PAD + " · ".join(who))
        if d.get("preview"):
            out.append(fmt._PAD + d["preview"])
        out.append(f"{fmt._PAD}id {d['event_id']}"
                   + (f" · series {d['series_id']}" if d.get("series_id") and d["series_id"] != d["event_id"]
                      else ""))
    return "\n".join(out)


def render_event(d: dict, tz_name: str) -> str:
    lines = [f"Subject: {d['subject']}", f"When: {when_text(d)} ({tz_name})"]
    if d.get("recurrence"):
        lines.append(f"Repeats: {d['recurrence']}")
    elif d.get("recurring"):
        lines.append("Repeats: yes (this is one occurrence of a series)")
    if d.get("location"):
        lines.append(f"Where: {d['location']}")
    if d.get("online_meeting"):
        om = d["online_meeting"]
        lines.append("Teams meeting" + (f": {om}" if isinstance(om, str) else ""))
    if d.get("organizer"):
        lines.append("Organiser: " + ("you" if d.get("is_organizer") else fmt.addr_text(d["organizer"])))
    if not d.get("is_organizer"):
        lines.append(f"Your response: {d.get('my_response')}" + (" (NEEDS RESPONSE)" if needs_response(d) else ""))
    if d.get("attendees"):
        lines.append(f"Attendees ({len(d['attendees'])}):")
        for a in d["attendees"]:
            lines.append(f"{fmt._PAD}{fmt.addr_text(a)} · {a['type']} · {a['response']}")
    extra = []
    if d.get("show_as"):
        extra.append(f"shown as {d['show_as']}")
    if d.get("reminder"):
        extra.append(f"reminder {d['reminder']}")
    for key, mark in (("cancelled", "CANCELLED"), ("private", "private"), ("attachments", "📎")):
        if d.get(key):
            extra.append(mark)
    if d.get("importance"):
        extra.append(f"!{d['importance']}")
    if d.get("categories"):
        extra.append("[" + ", ".join(d["categories"]) + "]")
    if extra:
        lines.append(" · ".join(extra))
    lines.append(f"Event id: {d['event_id']}")
    if d.get("series_id") and d["series_id"] != d["event_id"]:
        lines.append(f"Series id: {d['series_id']} (use it to change or answer the whole series)")
    if d.get("web_link"):
        lines.append(f"Open in Outlook: {d['web_link']}")
    lines.append("")
    lines.append(d.get("body") or "(no description)")
    return "\n".join(lines)


# ---------------------------------------------------------------- free time
def parse_hhmm(value: str) -> time:
    m = re.match(r"^\s*(\d{1,2}):(\d\d)", value or "")
    if not m:
        raise ValueError(value)
    h, mi = int(m.group(1)), int(m.group(2))
    return time(23, 59) if h >= 24 else time(h, mi)


def parse_days(value: str) -> list[int]:
    """'mon-fri' or 'sun,mon,tue,wed,thu' -> weekday numbers (Monday = 0)."""
    short = [d[:3] for d in DAYS]
    out: set[int] = set()
    for part in re.split(r"[\s,;]+", (value or "").strip().lower()):
        if not part:
            continue
        if "-" in part:
            a, b = (short.index(x[:3]) for x in part.split("-", 1))
            i = a
            while True:
                out.add(i)
                if i == b:
                    break
                i = (i + 1) % 7
        else:
            out.add(short.index(part[:3]))
    return sorted(out)


def days_text(days: list[int]) -> str:
    names = [d[:3].capitalize() for d in DAYS]
    if len(days) > 2 and days == list(range(days[0], days[-1] + 1)):
        return f"{names[days[0]]}–{names[days[-1]]}"
    return ", ".join(names[d] for d in days)


def merge(intervals: list[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    out: list[tuple[datetime, datetime]] = []
    for s, e in sorted(intervals):
        if out and s <= out[-1][1]:
            out[-1] = (out[-1][0], max(out[-1][1], e))
        else:
            out.append((s, e))
    return out


def _round_up(dt: datetime, step: int = 5) -> datetime:
    dt = dt.replace(second=0, microsecond=0) + (timedelta(minutes=1) if dt.second or dt.microsecond else
                                                 timedelta(0))
    extra = (-dt.minute) % step
    return dt + timedelta(minutes=extra)


def free_slots(window_start: datetime, window_end: datetime, busy: list[tuple[datetime, datetime]],
               minutes: int, tz: ZoneInfo, work: dict | None) -> list[tuple[datetime, datetime]]:
    """Gaps of at least `minutes` inside [window_start, window_end) not covered by
    `busy`. work = {days: [0..6], start: time, end: time, tz: ZoneInfo} limits each
    day to working hours (in the working-hours time zone); None = whole days."""
    need = timedelta(minutes=minutes)
    busy = merge([(s.astimezone(tz), e.astimezone(tz)) for s, e in busy if e > s])
    slots: list[tuple[datetime, datetime]] = []
    day = window_start.astimezone(tz).date()
    last = window_end.astimezone(tz).date()
    while day <= last:
        if work:
            wz = work.get("tz") or tz
            if day.weekday() not in work["days"]:
                day += timedelta(days=1)
                continue
            lo = datetime.combine(day, work["start"], wz).astimezone(tz)
            hi = datetime.combine(day, work["end"], wz).astimezone(tz)
        else:
            lo = datetime.combine(day, time(0), tz)
            hi = datetime.combine(day + timedelta(days=1), time(0), tz)
        lo, hi = max(lo, window_start), min(hi, window_end)
        cursor = _round_up(lo)
        for s, e in busy:
            if e <= cursor or s >= hi:
                continue
            if s - cursor >= need:
                slots.append((cursor, s))
            cursor = max(cursor, _round_up(e))
        if hi - cursor >= need:
            slots.append((cursor, hi))
        day += timedelta(days=1)
    return slots


def render_slots(d: dict) -> str:
    who = ", ".join(d["participants"][1:])
    head = (f"Free for {duration_text(d['minutes'])}" + (f" with {who}" if who else "")
            + f" (times {d['time_zone']}" + (f", working hours {d['working_hours']}" if d.get("working_hours")
                                              else "") + ")")
    out = [head]
    day = None
    for s in d["slots"]:
        st, en = datetime.fromisoformat(s["start"]), datetime.fromisoformat(s["end"])
        if st.date() != day:
            day = st.date()
            out.append(day_label(day))
        out.append(f"{fmt._PAD}{st:%H:%M}–{en:%H:%M} ({duration_text(s['minutes'])})")
    if not d["slots"]:
        out.append("No free time found in this range.")
    for u in d.get("unavailable") or []:
        out.append(f"(free/busy not available for {u['email']}: {u['reason']}; they are ignored)")
    if d.get("more"):
        out.append(f"(more free slots exist: showing the first {len(d['slots'])})")
    return "\n".join(out)
