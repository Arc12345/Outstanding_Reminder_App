# Outstanding Reminders

Automatic payment reminders to vendors on **WhatsApp** (official Business API) or **SMS**, from an Excel list.

Upload the outstanding list once. Every vendor gets a personalised message — their name, amount and date — on
their own schedule: every week, every month, every day, or from a chosen date. When amounts change, upload the
updated Excel; when they don't, the same reminders repeat. Pause or stop at any time.

Built for a small accounts team: one Windows app, no server, no monthly software fee
(WhatsApp's own per-message charge still applies).

| Home: upload and check | Per-vendor schedule | Status, pause and stop |
|---|---|---|
| ![Home](docs/screenshots/home-upload.png) | ![Schedule](docs/screenshots/vendor-schedule.png) | ![Status](docs/screenshots/status-pause.png) |

---

## Features

- **Excel in, reminders out.** Reads `.xlsx` or `.csv`, finds the name / mobile / amount columns, cleans Indian
  mobile numbers (`+91`, `0`, spaces, dashes), reads amounts like `₹25,000`, `Rs. 18,000.00` or Tally-style
  `12,500 Dr` / `3,000 Cr`, merges repeated rows, and lists every row it can't send with the reason.
- **Per-vendor schedules.** Every week (default), every month (on a chosen date), every day (except Sunday),
  or don't send — plus an optional *start from* date. Set them on screen or with optional Excel columns.
- **Automatic sending.** Windows Task Scheduler starts the app once a day; it sends only to vendors who are due.
  Nobody gets two reminders in the same day, week or month.
- **Pause / resume / stop.** Skip 1–4 weeks or pause until a date; it starts again by itself.
- **Safety checks.** Business hours only, a per-run limit, a WhatsApp setup check before every automatic run,
  and an automatic stop after 5 failures in a row (e.g. when the WhatsApp balance runs out).
- **Honest dates.** The message says *"as of <date the list was uploaded>"*, so a resent list never shows a newer
  date than its amounts.
- **History.** Every message, its status and any error, with CSV export.
- **Three ways to send:** WhatsApp Cloud API (recommended), SMS through an Android phone running
  [RestSMS](https://github.com/Xcreen/RestSMS), or *dry run* for practice.

## Quick start

### Option A — run from source (any OS)

```bash
pip install -r requirements.txt
python app.py
```

The browser opens at <http://127.0.0.1:5050>. First-time password: `change-me` (change it in **Settings**).
The app starts in **dry run**, so nothing is sent while you try it. Try it with `docs/sample_outstanding.xlsx`.

### Option B — the Windows app (`.exe`)

- **From GitHub:** open the **Actions** tab → **Build Windows app** → **Run workflow**. When it finishes, download
  `Outstanding-Reminders-windows` from the run's *Artifacts*. Pushing a tag such as `v1.0.0` also attaches the
  `.exe` to a GitHub Release.
- **On your own PC:** install Python 3.10+ ("Add Python to PATH"), then double-click `build_windows.bat`.
  The app appears in `dist\`.

Put `Outstanding Reminders.exe` in a permanent folder and double-click it. Its data files are created next to it.

> **Windows may block an unsigned app.** For *"Windows protected your PC"* click **More info → Run anyway**.
> **Smart App Control** has no such button: see [Running on Windows](#running-on-windows).

## Setting up

1. **Settings → Message details:** company name and the contact number vendors should call.
2. **Settings → Send with:** choose **WhatsApp (official Business API)** and fill in the connection details.
   Step-by-step guide: [docs/WHATSAPP_SETUP.md](docs/WHATSAPP_SETUP.md).
3. **Settings → Check WhatsApp setup:** all three checks should be green.
4. **Settings → Most messages per run:** at least the number of vendors on your list.
5. **Home → Upload updated list → Save as the vendor list**, then **Turn on automatic sending**.

Everyday use for non-technical staff: [docs/HOW_TO_USE.md](docs/HOW_TO_USE.md).

## The Excel file

The first row must be headings. Three columns are required; two are optional.

| Column | Example headings the app recognises | Notes |
|---|---|---|
| Vendor name | `Vendor Name`, `Party Name`, `Customer Name`, `Ledger` | required |
| Mobile number | `Mobile Number`, `Mobile No`, `Phone`, `WhatsApp` | 10-digit Indian mobile; `+91` / `0` are fine |
| Outstanding | `Outstanding`, `Amount Due`, `Balance`, `Amount` | 0 or negative = not messaged |
| Frequency | `Frequency`, `How often`, `Reminder` | optional: `Weekly`, `Monthly`, `Daily`, `Don't send` |
| Start date | `Start date`, `Remind from`, `Reminder date` | optional: e.g. `05-11-2026`; monthly reminders use its day |

Blank *Frequency* / *Start date* cells keep whatever was chosen in the app. See `docs/sample_outstanding.xlsx`.

## How automatic sending works

- Turning automatic sending on creates a Windows Task Scheduler task, **"Outstanding Reminders - weekly send"**,
  that runs `Outstanding Reminders.exe --auto-send` every day at the chosen time. It can wake the PC from sleep,
  and catches up if the PC was off. The PC must be on and the user logged in.
- Each run reads the saved vendor list (`weekly-list.xlsx`) and works out who is due today:

  | Schedule | Due on | If the PC was off |
  |---|---|---|
  | Every week | the weekly day (default Monday) | sent the next day |
  | Every month | the *start from* date's day of month (31st → last day) | sent within 2 days; a Sunday date goes out Monday |
  | Every day | Monday to Saturday | — |
  | Start from | nothing before that date | — |

- Each vendor's last reminder date is stored, so a vendor is never messaged twice in the same period, even if the
  check runs more than once.
- Before sending, the run checks: turned on, not paused, 9 AM–7 PM, list readable, within the per-run limit, and —
  for WhatsApp — that the number, token and approved template all work. If any check fails, nothing is sent and the
  reason is shown on Home and written to `auto-send.log`.

## Running on Windows

| Message | What to do |
|---|---|
| *Windows protected your PC* (SmartScreen) | **More info → Run anyway** (once). |
| *Smart App Control blocked an app* | No per-app exception exists. Try **Properties → Unblock** on the downloaded zip before extracting; otherwise turn Smart App Control off (Windows Security → App & browser control). Because the app runs every day, turning it back on would block those runs. |
| Antivirus quarantines the `.exe` | Mark it as safe; PyInstaller apps are sometimes flagged by mistake. |

The permanent fix is a code-signed `.exe` (a paid certificate in the company's name).

## Using it from another PC

- **Move it:** stop automatic sending, copy the whole folder (`.exe`, `reminders.db`, `weekly-list.xlsx`), then on
  the new PC press **Turn on automatic sending**. Never run automatic sending on two PCs — vendors would get
  duplicates.
- **Share it on the office network:** start it with `--host 0.0.0.0` (e.g. a shortcut ending in
  `"…\Outstanding Reminders.exe" --host 0.0.0.0`) and open `http://<that PC's IP>:5050` from other PCs.

## Project structure

```
app.py                       server, Excel reading, validation, schedules, senders, storage, CLI
ui/index.html                the app's screens (HTML, CSS and JavaScript in one file)
ui/login.html                login page
tests/                       pytest tests (number/amount parsing, validation, SMS length, schedule rules)
assets/icon.ico              app icon
build_windows.bat            builds the .exe on a Windows PC
.github/workflows/           builds the .exe on GitHub
docs/                        user guide, WhatsApp setup, sample Excel, screenshots
```

Files created next to the app at runtime (all git-ignored): `reminders.db`, `weekly-list.xlsx`, `auto-send.log`.

### Command line

```
python app.py [--host 127.0.0.1] [--port 5050] [--no-browser] [--auto-send]
```

| Environment variable | Purpose |
|---|---|
| `SMS_APP_PASSWORD` | Initial password (until changed in Settings) |
| `SMS_APP_DB` | Path of the database file (default: `reminders.db` next to the app) |
| `WA_API_BASE` | WhatsApp API base URL — only for testing against a stand-in server |

### Tests

```bash
pip install -r requirements-dev.txt
python -m pytest -q
```

## Security notes

- **Never commit `reminders.db`.** It contains the WhatsApp access token. `.gitignore` already excludes it, along
  with vendor lists (`*.xlsx`, `*.csv`) and logs.
- The token is stored only on the PC running the app, and is never shown again in full in the browser.
- Change the default password before use. Use `--host 0.0.0.0` only on a trusted office network.
- Consider keeping this repository **private** if it will contain anything company-specific.

## Limitations

- Automatic scheduling uses Windows Task Scheduler (the app itself also runs on macOS/Linux, without scheduling).
- "Accepted by WhatsApp" means WhatsApp accepted the message; delivered/read counts are in WhatsApp Manager.
  The app does not receive replies (no webhook server).
- WhatsApp only allows approved templates for business-initiated messages; changing the wording needs re-approval.
- SMS through a personal SIM is limited by the phone and the operator; use WhatsApp for large lists.

## License

No license file is included yet. Without one, others may view the code but have no right to reuse it.
Add a `LICENSE` file (for example MIT) if you want to allow that.
