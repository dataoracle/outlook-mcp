"""Offline tests for the calendar tools. Reuses the fake mailbox from
test_mail.py and adds a small fake calendar in front of it.

Run:  python tests/test_calendar.py      (or: pytest tests/)
Pass -v to print sample output.
"""
import json
import sys
from datetime import datetime, time, timezone
from pathlib import Path
from urllib.parse import unquote
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parent))

import httpx  # noqa: E402

import test_mail as h  # noqa: E402  (sets up the environment and the server)
from test_mail import ALICE, BOB, ME, call, j, server  # noqa: E402

import events as cal  # noqa: E402

NOW = datetime(2026, 11, 17, 6, 0, tzinfo=timezone.utc)  # Tuesday 10:00 in Dubai
server._now = lambda: NOW


def ev(eid, subject, start, end, *, organizer=ME, attendees=(), response="organizer", is_org=True,
       all_day=False, show_as="busy", online=False, etype="singleInstance", series=None, requested=True):
    return {"id": eid, "subject": subject, "isAllDay": all_day, "showAs": show_as,
            "start": {"dateTime": start + ".0000000", "timeZone": "UTC"},
            "end": {"dateTime": end + ".0000000", "timeZone": "UTC"},
            "location": {"displayName": "Room 4" if not online else ""}, "organizer": organizer,
            "attendees": [{**a, "type": "required", "status": {"response": r}} for a, r in attendees],
            "responseStatus": {"response": response}, "isOrganizer": is_org, "isOnlineMeeting": online,
            "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/abc"} if online else None,
            "isCancelled": False, "type": etype, "seriesMasterId": series, "sensitivity": "normal",
            "importance": "normal", "categories": [], "bodyPreview": "", "hasAttachments": False,
            "webLink": f"https://outlook.office.com/calendar/item/{eid}", "responseRequested": requested,
            "body": {"contentType": "text", "content": f"Agenda for {subject}"}}


def reset_calendar():
    global EVENTS, ACTIONS, NEXT
    NEXT = [0]
    ACTIONS = []
    EVENTS = {e["id"]: e for e in [
        ev("EV-0001-STANDUP", "Daily standup", "2026-11-17T05:00:00", "2026-11-17T05:30:00", organizer=ALICE,
           attendees=[(ME, "accepted")], response="accepted", is_org=False, online=True, etype="occurrence",
           series="EV-SERIES-STANDUP"),
        ev("EV-0002-BUDGET", "Budget review", "2026-11-17T08:00:00", "2026-11-17T09:00:00", organizer=BOB,
           attendees=[(ME, "none"), (ALICE, "accepted")], response="notResponded", is_org=False,
           show_as="tentative"),
        ev("EV-0003-FOCUS", "Focus time", "2026-11-18T06:00:00", "2026-11-18T08:00:00"),
        ev("EV-0004-SYNC", "Project sync", "2026-11-19T10:00:00", "2026-11-19T10:30:00",
           attendees=[(ALICE, "accepted"), (BOB, "none")]),
        ev("EV-0005-HOLIDAY", "Public holiday", "2026-11-20T00:00:00", "2026-11-21T00:00:00", all_day=True,
           show_as="free"),
    ]}


def _t(dt):
    return dt["dateTime"][:19]


def calendar(req: httpx.Request) -> httpx.Response | None:
    p = unquote(req.url.path).removeprefix("/v1.0")
    q = dict(req.url.params)
    body = json.loads(req.content) if req.content else None
    if p == "/me/calendars":
        return j({"value": [{"id": "CAL-DEFAULT-0001", "name": "Calendar", "isDefaultCalendar": True,
                             "canEdit": True, "owner": {"name": "Me", "address": "me@contoso.example"}},
                            {"id": "CAL-TEAM-000002", "name": "Team", "canEdit": False,
                             "owner": {"name": "Alice", "address": "alice@contoso.example"}}]})
    if p in ("/me/calendar/calendarView", "/me/calendars/CAL-TEAM-000002/calendarView"):
        assert q["$orderby"] == "start/dateTime" and q["startDateTime"].endswith("Z"), q
        rows = [e for e in EVENTS.values()
                if _t(e["end"]) > q["startDateTime"][:19] and _t(e["start"]) < q["endDateTime"][:19]]
        rows.sort(key=lambda e: _t(e["start"]))
        return j({"value": rows})
    if p == "/me/calendar/getSchedule":
        ACTIONS.append(("getSchedule", body))
        out = []
        for who in body["schedules"]:
            if who == "me@contoso.example":
                items = [{"status": e["showAs"], "start": e["start"], "end": e["end"]} for e in EVENTS.values()]
                out.append({"scheduleId": who, "scheduleItems": items, "workingHours": {
                    "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
                    "startTime": "08:00:00.0000000", "endTime": "17:00:00.0000000",
                    "timeZone": {"name": "Arabian Standard Time"}}})
            elif who == "alice@contoso.example":
                out.append({"scheduleId": who, "scheduleItems": [
                    {"status": "busy", "start": {"dateTime": "2026-11-17T07:00:00.0000000", "timeZone": "UTC"},
                     "end": {"dateTime": "2026-11-17T08:00:00.0000000", "timeZone": "UTC"}}]})
            else:
                out.append({"scheduleId": who, "error": {"message": "The user is outside your organization"}})
        return j({"value": out})
    if p in ("/me/calendar/events", "/me/calendars/CAL-TEAM-000002/events") and req.method == "POST":
        NEXT[0] += 1
        e = {**body, "id": f"EV-NEW-{NEXT[0]:04d}-ZZZZ", "isOrganizer": True, "type": "singleInstance",
             "webLink": "https://outlook.office.com/calendar/item/new",
             "organizer": ME, "responseStatus": {"response": "organizer"}}
        if body.get("isOnlineMeeting"):
            e["onlineMeeting"] = {"joinUrl": "https://teams.microsoft.com/l/meetup-join/new"}
        EVENTS[e["id"]] = e
        ACTIONS.append(("create", p, body))
        return j(e, 201)
    if p.startswith("/me/messages/INVITE-"):
        assert q["$expand"] == "microsoft.graph.eventMessage/event($select=id)", q
        return j({"id": p.rsplit("/", 1)[1], "subject": "Budget review", "event": {"id": "EV-0002-BUDGET"}})
    if p.startswith("/me/events/"):
        parts = p.split("/")
        e = EVENTS.get(parts[3])
        if e is None:
            return j({"error": {"code": "ErrorItemNotFound", "message": "not found"}}, 404)
        if len(parts) == 5:
            ACTIONS.append((parts[4], parts[3], body))
            if parts[4] == "decline":
                del EVENTS[parts[3]]
            return httpx.Response(202)
        if req.method == "GET":
            return j(e)
        if req.method == "PATCH":
            ACTIONS.append(("patch", parts[3], body))
            e.update(body)
            return j(e)
        if req.method == "DELETE":
            ACTIONS.append(("delete", parts[3], None))
            del EVENTS[parts[3]]
            return httpx.Response(204)
    return None


def graph(req: httpx.Request) -> httpx.Response:
    res = calendar(req)
    if res is not None:
        assert 'IdType="ImmutableId"' in req.headers.get("Prefer", ""), req.headers
        return res
    return h.graph(req)


server.graph._http = httpx.AsyncClient(base_url="https://graph.microsoft.com/v1.0",
                                       transport=httpx.MockTransport(graph))
reset_calendar()


def created():
    return [a for a in ACTIONS if a[0] == "create"]


# ------------------------------------------------------------------ tests
def test_whoami_shows_calendar_time_zone():
    out = call("whoami")
    assert "Calendar: times shown in Asia/Dubai (from mailbox settings)" in out, out
    assert "calendar" in out.rsplit("Server:", 1)[1], out


def test_list_events_today_in_local_time():
    reset_calendar()
    out = call("list_events", days=1)
    assert out.startswith("2 events, Tue 2026-11-17 (times Asia/Dubai)"), out
    assert "09:00–09:30 Daily standup · Teams · accepted · 🔁" in out, out
    assert "12:00–13:00 Budget review · Room 4 · NEEDS RESPONSE" in out, out
    assert "organiser Bob Smith <bob@partner.example> · 2 attendees (1 not responded, 1 accepted)" in out, out
    assert "series EV-SERIES-STANDUP" in out, out


def test_list_events_week_filters():
    reset_calendar()
    out = call("list_events", start="2026-11-16", end="2026-11-22")
    assert "5 events, Mon 2026-11-16 to Sun 2026-11-22" in out, out
    assert "Fri 2026-11-20\nall day Public holiday · Room 4 · shown as free" in out, out
    assert "14:00–14:30 Project sync · Room 4 · you organise" in out, out
    out = call("list_events", days=14, needs_response=True)
    assert out.startswith("1 event needing your response") and "Budget review" in out, out
    out = call("list_events", days=14, query="focus")
    assert "Focus time" in out and "Budget" not in out, out
    out = call("list_events", days=14, calendar="nope")
    assert "No calendar named 'nope'" in out and "Team" in out, out


def test_get_event():
    reset_calendar()
    out = call("get_event", event_id="EV-0001-STANDUP")
    assert "When: Tue 2026-11-17 09:00–09:30 (Asia/Dubai)" in out, out
    assert "Teams meeting: https://teams.microsoft.com/l/meetup-join/abc" in out, out
    assert "Organiser: Alice Johnson <alice@contoso.example>" in out and "Your response: accepted" in out, out
    assert "Series id: EV-SERIES-STANDUP" in out and out.endswith("Agenda for Daily standup"), out


def test_find_free_time_with_attendees():
    reset_calendar()
    out = call("find_free_time", duration_minutes=45, attendees="alice@contoso.example; eve@gmail.example",
               end="2026-11-20")
    sched = [a for a in ACTIONS if a[0] == "getSchedule"][-1][1]
    assert sched["schedules"] == ["me@contoso.example", "alice@contoso.example", "eve@gmail.example"], sched
    assert sched["startTime"] == {"dateTime": "2026-11-17T06:00:00", "timeZone": "UTC"}, sched
    expected = ("Tue 2026-11-17\n  10:00–11:00 (1h)\n  13:00–17:00 (4h)\n"
                "Wed 2026-11-18\n  08:00–10:00 (2h)\n  12:00–17:00 (5h)\n"
                "Thu 2026-11-19\n  08:00–14:00 (6h)\n  14:30–17:00 (2h30)\n"
                "Fri 2026-11-20\n  08:00–17:00 (9h)")
    assert expected in out, out
    assert "working hours Mon–Fri 08:00–17:00" in out and "free/busy not available for eve@gmail.example" in out
    out = call("find_free_time", duration_minutes=90, start="2026-11-17", end="2026-11-17",
               tentative_is_busy=False)
    assert "Tue 2026-11-17\n  10:00–17:00 (7h)" in out, out


def test_create_appointment_without_attendees():
    reset_calendar()
    out = call("create_event", subject="Dentist", start="2026-11-18T15:00", duration_minutes=60,
               reminder_minutes=-1, sensitivity="private")
    assert out.startswith("Created 'Dentist', Wed 2026-11-18 15:00–16:00 (Asia/Dubai)\nEvent id: EV-NEW-"), out
    body = created()[-1][2]
    assert body["start"] == {"dateTime": "2026-11-18T15:00:00", "timeZone": "Arabian Standard Time"}, body
    assert body["isReminderOn"] is False and "attendees" not in body, body
    out = call("create_event", subject="Leave", start="2026-11-25", end="2026-11-26", all_day=True)
    body = created()[-1][2]
    assert body["end"]["dateTime"] == "2026-11-27T00:00:00" and body["showAs"] == "free", body
    assert "Wed 2026-11-25 all day to Thu 2026-11-26" in out, out


def test_create_meeting_sends_invites_and_respects_allowlist():
    reset_calendar()
    out = call("create_event", subject="Kickoff", start="2026-11-19T09:00", attendees="eve@gmail.example")
    assert "not allowed" in out and not created(), out
    out = call("create_event", subject="Kickoff", start="2026-11-19T09:00", attendees="Alice")
    assert "name" in out.lower() and not created(), out
    out = call("create_event", subject="Weekly 1:1", start="2026-11-19T09:00",
               attendees="alice@contoso.example", online_meeting=True, repeat="weekly",
               repeat_until="2026-12-31", body="**Agenda**: plans", format="markdown")
    assert "invitations sent to 1 attendee" in out and "Teams link: https://teams" in out, out
    body = created()[-1][2]
    assert body["attendees"] == [{"emailAddress": {"address": "alice@contoso.example"}, "type": "required"}], body
    assert body["recurrence"]["pattern"] == {"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]}
    assert body["recurrence"]["range"]["endDate"] == "2026-12-31" and "<strong>Agenda</strong>" in \
        body["body"]["content"], body
    server.ENABLE_SEND = False
    try:
        out = call("create_event", subject="X", start="2026-11-19T09:00", attendees="alice@contoso.example")
        assert "ENABLE_SEND_TOOLS=false" in out and len(created()) == 1, out
    finally:
        server.ENABLE_SEND = True


def test_create_meeting_elicit_approval():
    reset_calendar()
    server.SEND_CONFIRMATION = "elicit"
    prompts = []

    async def decline(message, response_type, params, ctx):
        from fastmcp.client.elicitation import ElicitResult
        return ElicitResult(action="decline")

    async def approve(message, response_type, params, ctx):
        prompts.append(message)
        return "send"

    try:
        args = dict(subject="Kickoff", start="2026-11-19T09:00", attendees="alice@contoso.example",
                    optional_attendees="bob@partner.example")
        out = call("create_event", decline, **args)
        assert "did not approve" in out and not created(), out
        out = call("create_event", approve, **args)
        assert out.startswith("Created 'Kickoff'") and len(created()) == 1, out
        assert prompts[0].startswith("Send these meeting invitations as you?\nSubject: Kickoff\n"
                                     "When: Thu 2026-11-19 09:00–09:30 (Asia/Dubai)"), prompts
        assert "Required: alice@contoso.example" in prompts[0], prompts
        assert "Optional: bob@partner.example" in prompts[0], prompts
        out = call("create_event", approve, subject="Solo", start="2026-11-19T09:00")  # no attendees: no prompt
        assert out.startswith("Created 'Solo'") and len(prompts) == 1, out
    finally:
        server.SEND_CONFIRMATION = "client"


def test_update_event():
    reset_calendar()
    out = call("update_event", event_id="EV-0004-SYNC", start="2026-11-19T16:00")
    assert "Thu 2026-11-19 16:00–16:30" in out and "update sent to 2 attendees" in out, out
    patch = ACTIONS[-1][2]
    assert patch["end"] == {"dateTime": "2026-11-19T16:30:00", "timeZone": "Arabian Standard Time"}, patch
    out = call("update_event", event_id="EV-0004-SYNC", remove_attendees="bob@partner.example",
               add_attendees="eve@gmail.example")
    assert "not allowed" in out, out
    out = call("update_event", event_id="EV-0004-SYNC", remove_attendees="bob@partner.example")
    assert "update sent to 1 attendee" in out, out
    out = call("update_event", event_id="EV-0002-BUDGET", subject="Renamed")
    assert "attendee of this meeting" in out, out
    out = call("update_event", event_id="EV-0002-BUDGET", categories="Finance", show_as="busy")
    assert out.startswith("Updated 'Budget review'") and "update sent" not in out, out
    assert ACTIONS[-1][2] == {"showAs": "busy", "categories": ["Finance"]}, ACTIONS[-1]
    out = call("update_event", event_id="EV-0003-FOCUS", end="2026-11-18T13:00")
    assert "Wed 2026-11-18 10:00–13:00" in out and "update sent" not in out, out


def test_respond_to_event():
    reset_calendar()
    out = call("respond_to_event", response="accept", message_id="INVITE-0001-BUDGET", comment="See you")
    assert out.startswith("Accepted 'Budget review', Tue 2026-11-17 12:00–13:00") and "organiser was notified" \
        in out, out
    assert ACTIONS[-1] == ("accept", "EV-0002-BUDGET", {"sendResponse": True, "comment": "See you"}), ACTIONS[-1]
    out = call("respond_to_event", response="accept", event_id="EV-0002-BUDGET", propose_start="2026-11-18")
    assert "only be proposed" in out, out
    out = call("respond_to_event", response="decline", event_id="EV-0002-BUDGET",
               propose_start="2026-11-18T14:00")
    a = ACTIONS[-1]
    assert a[0] == "decline" and a[2]["proposedNewTime"]["end"]["dateTime"] == "2026-11-18T15:00:00", a
    assert "with a proposed new time" in out and "removed from your calendar" in out, out
    out = call("respond_to_event", response="tentative", event_id="EV-0004-SYNC")
    assert "You organise this meeting" in out, out
    server.ENABLE_SEND = False
    try:
        out = call("respond_to_event", response="accept", event_id="EV-0001-STANDUP")
        assert "send_response=False" in out, out
        out = call("respond_to_event", response="accept", event_id="EV-0001-STANDUP", send_response=False)
        assert "organiser was not notified" in out and ACTIONS[-1][2] == {"sendResponse": False}, out
    finally:
        server.ENABLE_SEND = True


def test_cancel_event():
    reset_calendar()
    out = call("cancel_event", event_id="EV-0002-BUDGET")
    assert "not its organiser" in out, out
    out = call("cancel_event", event_id="EV-0004-SYNC", comment="No longer needed")
    assert "cancellation sent to 2 attendees" in out, out
    assert ACTIONS[-1] == ("cancel", "EV-0004-SYNC", {"Comment": "No longer needed"}), ACTIONS[-1]
    out = call("cancel_event", event_id="EV-0003-FOCUS")
    assert out.startswith("Deleted 'Focus time'") and ACTIONS[-1][0] == "delete", out


def test_calendar_helpers():
    tz = ZoneInfo("Asia/Dubai")
    assert cal.parse_days("sun-thu") == [0, 1, 2, 3, 6] and cal.parse_days("fri-mon") == [0, 4, 5, 6]
    assert cal.days_text([0, 1, 2, 3, 4]) == "Mon–Fri" and cal.duration_text(150) == "2h30"
    assert cal.zone("Arabian Standard Time").key == "Asia/Dubai" and cal.zone("Nowhere/Zone") is None
    assert cal.parse_local("2026-11-18T05:00:00Z", tz, NOW).hour == 9
    assert cal.parse_local("tomorrow", tz, NOW).isoformat() == "2026-11-18T00:00:00+04:00"
    s = datetime(2026, 11, 17, 9, 7, tzinfo=tz)
    e = datetime(2026, 11, 17, 12, 0, tzinfo=tz)
    busy = [(datetime(2026, 11, 17, 10, 0, tzinfo=tz), datetime(2026, 11, 17, 10, 52, tzinfo=tz))]
    slots = cal.free_slots(s, e, busy, 30, tz, None)
    assert [(a.strftime("%H:%M"), b.strftime("%H:%M")) for a, b in slots] == [("09:10", "10:00"),
                                                                                ("10:55", "12:00")], slots
    work = {"days": [1], "start": time(9, 30), "end": time(11, 0), "tz": tz}
    slots = cal.free_slots(s, e, busy, 30, tz, work)
    assert [(a.strftime("%H:%M"), b.strftime("%H:%M")) for a, b in slots] == [("09:30", "10:00")], slots


if __name__ == "__main__":
    h.VERBOSE = "-v" in sys.argv
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_")]
    for name, fn in tests:
        fn()
        print(f"ok  {name}")
    print(f"{len(tests)} passed")
