# Google Sheets Attendance Check-In

A cross-platform, one-class-per-day attendance system using Google Sheets and
Google Apps Script.

## Features

- Personalized, private tokenized check-in links
- Works on iPhone, Android, tablets, and desktop
- Server-generated `America/Chicago` date and time
- One check-in per person per date
- `People`, `Attendance`, and `Settings` sheets
- Automatic annual report tabs grouped by date
- Admin menu for initialization, registration, link refresh, and report rebuilds
- Hosted by Google Apps Script; no always-on local machine required

## Files

- `Code.gs` — backend, Sheet setup, admin menu, check-in logic, and reports
- `Index.html` — mobile-friendly check-in page
- `appsscript.json` — Apps Script manifest
- `SETUP.md` — complete installation and deployment instructions
- `.clasp.json.example` / `.claspignore` — optional GitHub-to-Apps-Script sync

Start with [SETUP.md](SETUP.md).

## Security notes

Treat each personal link like a membership card. Anyone who has a link can use
it to check in as that person. Tokens are long and random, phone numbers are not
embedded in links, inactive people are rejected, and duplicate daily check-ins
are blocked. For stronger in-person verification, add a rotating classroom code
or staff-operated scanner in a future version.

Do not make the Google Sheet public. Only administrators need access to it.
