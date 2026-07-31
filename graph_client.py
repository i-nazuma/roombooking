"""Microsoft Graph access: app-only (client credentials) auth against the
room mailbox's calendar.

No human ever signs in — the app authenticates as itself via CLIENT_SECRET,
relying on the Calendars.ReadWrite Application permission (admin-consented)
on the app registration to read/write the room mailbox's calendar directly.
"""

import json
import os
import re
from datetime import datetime, timedelta
from urllib.parse import quote

import msal
import requests

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]

with open(CONFIG_PATH, encoding="utf-8") as f:
    CONFIG = json.load(f)

_app = msal.ConfidentialClientApplication(
    client_id=CONFIG["client_id"],
    client_credential=CONFIG["client_secret"],
    authority=f"https://login.microsoftonline.com/{CONFIG['tenant_id']}",
)

_ROOM_PATH = f"/users/{quote(CONFIG['room_mailbox'])}"


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


def _parse_graph_datetime(value):
    # Graph returns fractional seconds with up to 7 digits (e.g. ".0000000"),
    # which datetime.fromisoformat rejects on Python < 3.11. Truncate to 6.
    value = re.sub(r"(\.\d{6})\d+", r"\1", value)
    return datetime.fromisoformat(value)


def get_today_agenda():
    """Today's events on the room calendar (local time), earliest first."""
    tz = CONFIG["timezone"]
    now = datetime.now()
    start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_day = start_of_day + timedelta(days=1)

    data = _graph_get(
        f"{_ROOM_PATH}/calendarView",
        params={
            "startDateTime": start_of_day.isoformat(),
            "endDateTime": end_of_day.isoformat(),
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
        agenda.append(
            {
                "start": event["start"]["dateTime"][11:16],
                "end": event["end"]["dateTime"][11:16],
                "titel": event.get("subject") or "(kein Titel)",
                "gebuchtVon": event.get("organizer", {})
                .get("emailAddress", {})
                .get("name", "Unbekannt"),
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


def create_booking(employee, duration_minutes):
    """Create an event on the room mailbox's own calendar."""
    tz = CONFIG["timezone"]
    now = datetime.now()
    end = now + timedelta(minutes=duration_minutes)

    body = {
        "subject": f"Raumbuchung – {employee}",
        "body": {"contentType": "Text", "content": f"Gebucht über das Tablet für {employee}."},
        "start": {"dateTime": now.isoformat(), "timeZone": tz},
        "end": {"dateTime": end.isoformat(), "timeZone": tz},
    }
    return _graph_post(f"{_ROOM_PATH}/events", body, tz=tz)
