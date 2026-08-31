# Cafe Tan90 Attendance

A staff attendance + task tracker for Cafe Tan90 (Miw, Nureen, Richa). One page, one URL — a "Staff view" tab for logging in/out and marking tasks done, and a PIN-gated "Admin view" tab for setting daily tasks and managing the employee roster — backed by a live Google Sheet as the record, no manual export, ever.

## Why this runs on Google Apps Script, not a normal web host

The backend needs to create Google Sheets, write to them, and read them back in real time. Apps Script is the only place that can do that directly (via `SpreadsheetApp` / `DriveApp`), for free, with no server to run or pay for. Hosting this on Vercel/Netlify/etc. instead would mean rebuilding the whole backend against the Google Sheets API with a service-account credential to manage — a real option if that's ever wanted, but a different, bigger project than what's here. This repo is the Apps Script source, kept in git for version history and review, not for git-based deployment.

**Deploying still means the one-time Apps Script setup in [SETUP.md](./SETUP.md)** — copy/pasting these files into script.google.com and clicking Deploy. That step has to happen from the owning Google account, so it can't be done from here; SETUP.md walks through it in about 5 minutes.

## What's in here

- `appscript/Code.gs` — server-side logic: employee roster (add/rename/remove, up to 8, tracked by a stable slot id so a rename or removal never loses or splits history), standing/recurring tasks (auto-seeded onto every matching day), monthly spreadsheet auto-creation (one Google Sheet per calendar month, a tab per day, pre-built in advance), attendance log-in/out, task assignment + completion tracking, admin PIN check, and payroll (per-employee hourly rate history with effective dates, exact-to-the-second hours from log-in/out, day and browsable month totals).
- `appscript/Staff.html` — the single page that's actually served (`doGet()` always renders this one). Has the shared-screen Staff tab (log in/out, mark tasks done) and pulls in `Admin.html` as its Admin tab via Apps Script's `include()` — one URL, two tabs, no separate admin link to remember.
- `appscript/Admin.html` — **not served on its own** — a body-only partial included into `Staff.html`'s Admin tab. PIN-gated (the tab is visible to anyone, but nothing behind it renders until the PIN is entered), split into two sub-tabs: **Tasks** (one unified task list for any date, each task markable "Repeats" on chosen days of the week right on its own row and reorderable by dragging its ≡ handle, a "Plan tasks ahead" calendar for one-off tasks across a week or month, employee management, and a shortcut to open the live Sheet) and **Payroll** (set each employee's hourly rate — with an effective date so a raise can start today, be backdated, or be scheduled ahead — and browse exact-hours pay for any day or any month, past or present). Typing into any task or employee name field auto-capitalizes each word as you go.
- `appscript/Stylesheet.html` — shared styling (navy/card look, matching the Enrich Lead Allocation demo this was modeled after).
- `appscript/appsscript.json` — Apps Script project manifest (timezone, web app config).
- `preview.html` — a standalone, in-browser mockup (no backend) of both pages, useful for reviewing the design/flow without deploying anything.
- `SETUP.md` — step-by-step deployment guide, plus explanations of how the monthly-doc and roster automation actually work.

## Quick start

1. Read [SETUP.md](./SETUP.md) and follow the one-time setup (about 5 minutes).
2. Open `preview.html` in a browser first if you just want to see the design before deploying anything.

## Local development notes

There's no build step — every file here is pasted directly into the Apps Script editor as-is. If you edit `Code.gs` or any `.html` file, copy the updated contents into the corresponding file in your Apps Script project and redeploy (Deploy → Manage deployments → New version) to make the change live.
