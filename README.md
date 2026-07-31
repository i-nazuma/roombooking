# roombooking
small hackathon

# Besprechungsraum-Anzeige (Room Display PWA)

A small PWA + backend that reads a dedicated Outlook **room resource mailbox** via
Microsoft Graph and shows live availability on a tablet, styled like your reference
screenshot (big status card: **Belegt / Frei**, "Wieder frei ab …", and today's
agenda below).

Architecture: a tiny Node/Express server holds your Graph credentials and calls
Microsoft Graph app-only (client credentials flow) once a minute; it serves both the
`/api/status` JSON and the static PWA. The tablet's browser never sees your client
secret — it only talks to your server.

---

## 1. Register an app in Microsoft Entra ID (Azure AD)

1. Go to **entra.microsoft.com** → **App registrations** → **New registration**.
2. Name it something like `Room Display`, leave redirect URI empty, register.
3. Note down, from the app's **Overview** page:
    - **Application (client) ID** → `CLIENT_ID`
    - **Directory (tenant) ID** → `TENANT_ID`
4. Go to **Certificates & secrets** → **New client secret** → copy the value
   immediately (shown once) → `CLIENT_SECRET`.
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Application permissions** → add **`Calendars.Read`** (or `Calendars.ReadWrite`
   if you'd ever want to write to it) → **Grant admin consent** (needs a Global
   Admin / Exchange Admin) for this to work.

### Recommended: scope the app to only this one mailbox

By default `Calendars.Read` (application) lets the app read **every** mailbox in
the tenant. Since you only need one room, lock it down with an Exchange Online
PowerShell **Application Access Policy**:

```powershell
Connect-ExchangeOnline

New-DistributionGroup -Name "RoomDisplayScope" -Type Security -Members besprechungsraum@firma.com

New-ApplicationAccessPolicy `
  -AppId <CLIENT_ID> `
  -PolicyScopeGroupId "RoomDisplayScope" `
  -AccessRight RestrictAccess `
  -Description "Room display app can only read the meeting room mailbox"
```

This is optional but strongly recommended — ask your Microsoft 365 admin if you're
not the one managing Exchange.

---

## 2. Configure the server

```bash
cp .env.example .env
```

Fill in `.env`:

```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
ROOM_MAILBOX=besprechungsraum@firma.com   # the room resource mailbox's address
ROOM_DISPLAY_NAME=Besprechungsraum
TIMEZONE=Europe/Vienna
PORT=3000
```

---

## 3. Run it

You need a machine that stays on and reachable by the tablet over your local
network (a Raspberry Pi, a small VPS, an old mini-PC, or even a NAS with Node
support all work well — this does **not** need to run on the tablet itself).

```bash
npm install
npm start
```

The server now serves everything at `http://<server-ip>:3000`.

For production, keep it alive with a process manager, e.g.:

```bash
npm install -g pm2
pm2 start server.js --name room-display
pm2 save
pm2 startup
```

---

## 4. Show it on the Android tablet

1. On the tablet, open Chrome and go to `http://<server-ip>:3000`.
2. Tap the Chrome menu → **Add to Home screen** (or Chrome may prompt to install
   automatically) → this installs it as a standalone PWA using `manifest.json`
   (fullscreen, landscape, your icon).
3. Open the installed app from the home screen.
4. Tap the fullscreen icon in the top-right corner for a distraction-free kiosk
   look, and disable the tablet's screen lock/sleep (Settings → Display → Sleep →
   Never, or use "Stay awake while charging" in Developer Options) since it'll sit
   mounted on a wall permanently plugged in.

### For a fully locked-down kiosk (recommended for a wall-mounted tablet)

A plain installed PWA can still be swiped away or exited. For a true kiosk that
survives reboots and can't be backed out of, use a kiosk-launcher app such as
**Fully Kiosk Browser** (free tier is enough):
- Set the start URL to `http://<server-ip>:3000`
- Enable "Autostart on boot" and "Keep screen on"
- Enable "Motion detection" or scheduled screen-on/off if you want it to sleep
  overnight

---

## How the display behaves

- Polls `/api/status` every 60 seconds.
- **Belegt** (amber, lock icon) while a non-free event on the room's calendar is
  currently running, showing "Wieder frei ab HH:MM Uhr".
- **Frei** (green, unlock icon) otherwise, showing the next upcoming appointment
  if there is one today.
- Today's full agenda is listed below, with the current meeting (if any)
  highlighted.
- If the tablet loses network or the server is unreachable, it keeps showing the
  **last successfully loaded data** with an "Offline – letzter Stand" badge
  (matching your screenshot) instead of going blank.

## Notes / things you may want to adjust

- Meetings marked "Free" in Outlook (showAs: free) are excluded from occupancy —
  only Busy/Tentative/OOF/Working-elsewhere block the room.
- Cancelled events are filtered out.
- Colors, type, and spacing live in `public/styles.css` if you want to match your
  company branding more closely.
- If your organization uses a different resource-booking flow (e.g. a shared
  calendar instead of a bookable room mailbox), the same `/api/status` endpoint
  can point `ROOM_MAILBOX` at that shared calendar's mailbox address — no other
  changes needed as long as Graph can read a `calendarView` for it.
