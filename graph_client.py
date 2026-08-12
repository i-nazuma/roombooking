"""Microsoft Graph access: app-only (client credentials) auth against the
room mailbox's calendar.

No human ever signs in — the app authenticates as itself via CLIENT_SECRET,
relying on the Calendars.ReadWrite Application permission (admin-consented)
on the app registration to read/write the room mailbox's calendar directly.
"""

import json
import os
import re
import time
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote
from zoneinfo import ZoneInfo

import msal
import requests

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]

with open(CONFIG_PATH, encoding="utf-8") as f:
    CONFIG = json.load(f)

_LOCAL_TZ = ZoneInfo(CONFIG["timezone"])


def _to_utc_iso(local_naive_dt):
    # calendarView's startDateTime/endDateTime query params are always
    # interpreted as UTC by Graph, regardless of the Prefer: outlook.timezone
    # header (that header only affects how *response* dateTimes are
    # rendered). Naive local datetimes must be converted explicitly.
    return local_naive_dt.replace(tzinfo=_LOCAL_TZ).astimezone(timezone.utc).isoformat()

_app = msal.ConfidentialClientApplication(
    client_id=CONFIG["client_id"],
    client_credential=CONFIG["client_secret"],
    authority=f"https://login.microsoftonline.com/{CONFIG['tenant_id']}",
)

_ROOM_PATH = f"/users/{quote(CONFIG['room_mailbox'])}"
_ROOM_MAILBOX_LOWER = CONFIG["room_mailbox"].lower()


class BookingConflict(Exception):
    """Raised when the requested slot overlaps an existing event."""

    def __init__(self, conflicting_event):
        self.subject = conflicting_event.get("subject") or "(kein Titel)"
        self.start = conflicting_event["start"]["dateTime"][11:16]
        self.end = conflicting_event["end"]["dateTime"][11:16]
        super().__init__(f"{self.subject} ({self.start}–{self.end})")


def acquire_token():
    """Return a valid app-only access token (MSAL caches/renews internally)."""
    result = _app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown authentication error"))
        raise RuntimeError(f"Authentication failed: {error}")
    return result["access_token"]


def _graph_get(path, params=None, tz=None):
    headers = {"Authorization": f"Bearer {acquire_token()}"}
    if tz:
        headers["Prefer"] = f'outlook.timezone="{tz}"'
    resp = requests.get(f"{GRAPH_BASE}{path}", headers=headers, params=params, timeout=15)
    resp.raise_for_status()
    return resp.json()


def _graph_post(path, body, tz=None):
    headers = {"Authorization": f"Bearer {acquire_token()}", "Content-Type": "application/json"}
    if tz:
        headers["Prefer"] = f'outlook.timezone="{tz}"'
    resp = requests.post(f"{GRAPH_BASE}{path}", headers=headers, json=body, timeout=15)
    resp.raise_for_status()
    return resp.json()


def _graph_delete(path):
    headers = {"Authorization": f"Bearer {acquire_token()}"}
    resp = requests.delete(f"{GRAPH_BASE}{path}", headers=headers, timeout=15)
    resp.raise_for_status()


def _parse_graph_datetime(value):
    # Graph returns fractional seconds with up to 7 digits (e.g. ".0000000"),
    # which datetime.fromisoformat rejects on Python < 3.11. Truncate to 6.
    value = re.sub(r"(\.\d{6})\d+", r"\1", value)
    return datetime.fromisoformat(value)


def get_agenda(for_date=None):
    """Events on the room calendar for one day (local time), earliest first.

    `for_date` is a `datetime.date`; defaults to today.
    """
    tz = CONFIG["timezone"]
    day = for_date or date.today()
    start_of_day = datetime(day.year, day.month, day.day)
    end_of_day = start_of_day + timedelta(days=1)

    data = _graph_get(
        f"{_ROOM_PATH}/calendarView",
        params={
            "startDateTime": _to_utc_iso(start_of_day),
            "endDateTime": _to_utc_iso(end_of_day),
            "$select": "subject,start,end,organizer,isCancelled",
            "$orderby": "start/dateTime",
            "$top": 50,
        },
        tz=tz,
    )

    agenda = []
    for event in data.get("value", []):
        if event.get("isCancelled"):
            continue
        organizer_address = (
            event.get("organizer", {}).get("emailAddress", {}).get("address", "")
        )
        agenda.append(
            {
                "id": event["id"],
                "start": event["start"]["dateTime"][11:16],
                "end": event["end"]["dateTime"][11:16],
                "titel": event.get("subject") or "(kein Titel)",
                "gebuchtVon": event.get("organizer", {})
                .get("emailAddress", {})
                .get("name", "Unbekannt"),
                "loeschbar": organizer_address.lower() == _ROOM_MAILBOX_LOWER,
                "_startDateTime": event["start"]["dateTime"],
                "_endDateTime": event["end"]["dateTime"],
            }
        )
    return agenda


def get_current_status(agenda):
    """Derive busy/free from an already-fetched agenda (no extra Graph call)."""
    now = datetime.now()
    for event in agenda:
        start = _parse_graph_datetime(event["_startDateTime"])
        end = _parse_graph_datetime(event["_endDateTime"])
        if start <= now <= end:
            return {"belegt": True, "freiAb": event["end"], "belegtBis": event["end"]}
    return {"belegt": False, "freiAb": None, "belegtBis": None}


def _find_conflict(start, end):
    """Return the first non-cancelled event overlapping [start, end], or None."""
    tz = CONFIG["timezone"]
    data = _graph_get(
        f"{_ROOM_PATH}/calendarView",
        params={
            "startDateTime": _to_utc_iso(start),
            "endDateTime": _to_utc_iso(end),
            "$select": "subject,start,end,isCancelled",
            "$top": 1,
        },
        tz=tz,
    )
    for event in data.get("value", []):
        if not event.get("isCancelled"):
            return event
    return None


def create_booking(employee, duration_minutes, start=None):
    """Create an event on the room mailbox's own calendar.

    Raises BookingConflict (without creating anything) if the slot overlaps
    an existing event.
    """
    tz = CONFIG["timezone"]
    start = start or datetime.now()
    end = start + timedelta(minutes=duration_minutes)

    conflict = _find_conflict(start, end)
    if conflict:
        raise BookingConflict(conflict)

    body = {
        "subject": f"Raumbuchung – {employee}",
        "body": {"contentType": "Text", "content": f"Gebucht über das Tablet für {employee}."},
        "start": {"dateTime": start.isoformat(), "timeZone": tz},
        "end": {"dateTime": end.isoformat(), "timeZone": tz},
    }
    return _graph_post(f"{_ROOM_PATH}/events", body, tz=tz)


_employees_cache = {"names": None, "fetched_at": 0.0}
_EMPLOYEES_CACHE_TTL_SECONDS = 7 * 24 * 60 * 60  # 1 week


def get_employees():
    """Display names of enabled members of the configured Entra ID group,
    sorted alphabetically.

    Falls back to config.json's static "employees" list if
    `employee_group_id` isn't set or the Graph call fails (e.g.
    GroupMember.Read.All not yet granted/consented) — booking should keep
    working even if the directory lookup is temporarily unavailable. A
    failed lookup doesn't get cached, so the next request retries Graph
    immediately rather than being stuck on the fallback for the TTL.
    """
    now = time.monotonic()
    if _employees_cache["names"] is not None and now - _employees_cache["fetched_at"] < _EMPLOYEES_CACHE_TTL_SECONDS:
        return _employees_cache["names"]

    group_id = CONFIG.get("employee_group_id")
    if not group_id:
        return CONFIG.get("employees", [])

    try:
        names = []
        path = f"/groups/{group_id}/members"
        params = {"$select": "displayName,accountEnabled", "$top": 999}
        while path:
            data = _graph_get(path, params=params)
            for member in data.get("value", []):
                if member.get("accountEnabled", True) and member.get("displayName"):
                    names.append(member["displayName"])
            next_link = data.get("@odata.nextLink")
            path = next_link[len(GRAPH_BASE):] if next_link else None
            params = None
        names.sort()
    except Exception:  # noqa: BLE001 - fall back rather than break booking
        return CONFIG.get("employees", [])

    _employees_cache["names"] = names
    _employees_cache["fetched_at"] = now
    return names


def delete_booking(event_id):
    """Delete an event, but only if it was booked by this app (organizer ==
    the room mailbox itself). Raises PermissionError otherwise.
    """
    event = _graph_get(f"{_ROOM_PATH}/events/{event_id}", params={"$select": "organizer"})
    organizer_address = event.get("organizer", {}).get("emailAddress", {}).get("address", "")
    if organizer_address.lower() != _ROOM_MAILBOX_LOWER:
        raise PermissionError("Dieser Termin wurde nicht über das Tablet gebucht.")
    _graph_delete(f"{_ROOM_PATH}/events/{event_id}")
