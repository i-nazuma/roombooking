# roombooking

small hackathon

# Besprechungsraum-Anzeige (Room Display PWA)

A small PWA + Python backend that shows live availability of the meeting
room's Outlook calendar on a tablet (big status card: **Belegt / Frei**,
"Wieder frei ab …", today's agenda) and lets you book the room, which
creates a real event on that calendar via Microsoft Graph.

Architecture: **app-only** (client credentials) access to the room's
resource mailbox (`besprechungszimmer@coredat.com`) — the app authenticates
as itself via a client secret, so nobody ever has to sign in on the tablet.
`server.py` runs directly on the tablet, serving both the app and the
`/api`.

---

## 1. App registration (already done)

An Entra ID app registration already exists for this
(`dc3d1b08-9bd1-4778-abcf-a9155e478ee2`, tenant
`7c47ee9e-02e3-49a0-9b2d-13d70357e4af`) with the Microsoft Graph
**Application permission** `Calendars.ReadWrite`, admin-consented — that's
what lets the server both read and create events on the room mailbox
without a delegated user login.

If you ever need to recreate it: Entra ID → App registrations → New
registration → Certificates & secrets → new client secret → API permissions
→ Microsoft Graph → **Application permissions** → `Calendars.ReadWrite` →
Grant admin consent (needs a Global/Exchange admin).

### Recommended: scope the app to only this one mailbox

By default, `Calendars.ReadWrite` (Application) lets the app read/write
**every** mailbox in the tenant. Locking it to just the room mailbox is
optional but a good idea — ask whoever manages Exchange Online to run:

```powershell
Connect-ExchangeOnline

New-DistributionGroup -Name "RoomDisplayScope" -Type Security -Members besprechungszimmer@coredat.com

New-ApplicationAccessPolicy `
  -AppId dc3d1b08-9bd1-4778-abcf-a9155e478ee2 `
  -PolicyScopeGroupId "RoomDisplayScope" `
  -AccessRight RestrictAccess `
  -Description "Room display app can only access the meeting room mailbox"
```

---

## 2. Configure the server

```bash
pip install -r requirements.txt
cp config.example.json config.json
```

Fill in `config.json`:

```json
{
  "client_id": "...",
  "client_secret": "...",
  "tenant_id": "...",
  "room_mailbox": "besprechungszimmer@coredat.com",
  "timezone": "Europe/Vienna",
  "employees": ["Ibrahim", "Alex", "Oli", "Steffi", "Antonia", "Nico", "Mario", "Sonja"]
}
```

`config.json` is gitignored — **never commit it**, it holds the client
secret. Treat that secret like a password: anyone with it can read/write
every mailbox the app has been granted access to.

---

## 3. Run it (on the tablet itself)

```bash
python server.py
```

No login prompt — the app authenticates itself on every request using the
client secret. The app is served at `http://<tablet-ip>:5000`.

---

## 4. Show it on the tablet

1. Open the tablet's browser to `http://localhost:5000`.
2. Add to Home screen / install as PWA (uses `manifest.json`: fullscreen,
   landscape, icon).
3. Tap the fullscreen icon top-right for a distraction-free look, and
   disable screen lock/sleep on the tablet since it'll run continuously.

---

## How it behaves

- Polls `/api/status` and `/api/agenda` every 30 seconds.
- **Belegt** (amber) while an event on the room's calendar is currently
  running, showing "Wieder frei ab HH:MM Uhr". **Frei** (green) otherwise.
  Only shown while viewing today — hidden while browsing another day, since
  it's inherently about right-now.
- The agenda shows one day at a time (today by default). Swipe left/right
  on it, or use the ‹ › arrows, to browse other days; a "Heute" button
  appears once you've navigated away from today.
- "Raum jetzt buchen" creates a real event starting now, for the chosen
  duration. "Termin planen" opens the same panel with an added date/time
  picker to book a specific future slot instead (defaults to whatever day
  you're currently viewing).
- Both booking flows are rejected with a clear "already booked" message if
  the requested slot overlaps an existing event on the room's calendar —
  no double-bookings.
- Agenda entries booked through the tablet (organizer = the room mailbox
  itself) show a delete button. Entries organized by anyone else — i.e.
  real meetings that simply invited the room via Outlook — don't; this is
  re-checked server-side on every delete, not just hidden in the UI.
- Cancelled events are filtered out.
- If a fetch to `/api` fails (network hiccup, Graph error), the tablet keeps
  showing the last successfully loaded data with an "Offline – letzter
  Stand" badge instead of breaking.

## Notes / things you may want to adjust

- The employee list is a static array in `config.json`, not pulled from the
  organization's directory (avoids needing extra directory-read
  permissions).
- Since the app authenticates as itself (not as a specific person), every
  booking's organizer on the calendar will be the room mailbox itself —
  who actually booked it is only recorded in the event's subject/body
  (`Raumbuchung – {employee}`).
