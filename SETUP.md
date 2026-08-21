# Cafe Tan90 Attendance — setup guide

This turns into a small web app hosted by Google (no separate server, no cost). Once it's deployed, it lives at a permanent URL you can bookmark or pin as a shortcut on the shared device.

## What you're setting up

- **Staff page** (default URL) — Miw, Nureen, and Richa share this on one device. Each taps their name's "Log in" / "Log out" button, and "Mark done" on each task as they finish it.
- **Admin page** (URL + `?page=admin`) — PIN-protected, for setting each day's task assignments, managing the employee roster, and jumping straight to the raw Google Sheet.
- **Google Sheets backend** — every log-in/out and every "Mark done" writes straight into a Google Sheet. No manual export, ever. A new spreadsheet is created automatically the first time anyone uses the app in a new calendar month, named "Cafe Tan90 Attendance — <Month Year>" (e.g. "Cafe Tan90 Attendance — August 2026"), with one tab per day of that month (28-31 tabs) — each tab holding that day's attendance log, task assignments, and task-completion times. One document per month, so a month's worth of pages all live in one place.

## How the automation actually works (the questions you asked)

**"How is a day marked done, and how is it known that it's done?"**
There's no button for this and nothing to remember — it's automatic, driven purely by the calendar date. Every action (log in, log out, mark done, viewing the Staff page) always reads and writes whichever day-tab matches *today's* date in Bangkok time. There's no separate "close the day" step because a day doesn't need to be explicitly closed — it just accumulates whatever happened on it. If you want a human-readable sense of whether a day is "finished," that's just: did everyone log out, and are all their assigned tasks showing a green "Done" pill? You can see that at a glance on the Staff page or by opening the day's tab in Sheets — but nothing needs to be marked or triggered.

**"How does tomorrow's page show up? What about next month's?"**
It already exists before you need it. The moment the app is used for the first time in a given calendar month, it builds the *entire* month's spreadsheet up front — every day tab for that month, pre-created, sitting there ready. So "tomorrow's page" (or "next month's doc") isn't something that gets created at midnight or by a trigger; it was already sitting there from the first use of the month. The only thing that happens automatically at a date rollover is that the Staff page starts *reading and writing* the next tab instead of today's — because "today" is just computed from the clock, not stored anywhere. (The very first time the app is opened in a new month, building all 28-31 tabs takes a few extra seconds — after that it's instant for the rest of the month.)

In short: nothing to schedule, nothing to press. Open the Staff page any day, any month, and it's already on the right page.

## One-time setup (about 5 minutes)

1. Go to **script.google.com** and sign in with the Google account you want to own this (this account's Drive is where the monthly spreadsheets will live).
2. Click **New project**.
3. Rename the project (top left) to something like "Cafe Tan90 Attendance".
4. You'll see a default `Code.gs` file already open — delete its contents and paste in the contents of **Code.gs** from this bundle.
5. Add the HTML files: click the **+** next to "Files" → **HTML** → name it exactly `Staff` (Apps Script adds `.html` itself) → paste in **Staff.html**'s contents.
6. Repeat step 5 for `Admin` (paste **Admin.html**) and `Stylesheet` (paste **Stylesheet.html**).
7. Open **appsscript.json**: click the gear icon (Project Settings) on the left, tick "Show `appsscript.json` manifest file in editor", then open it from the file list and replace its contents with the **appsscript.json** from this bundle.
8. **Set your admin PIN**: in `Code.gs`, find the `setAdminPin` function near the bottom, change `'2468'` to whatever PIN you want, then select `setAdminPin` from the function dropdown at the top and click **Run** once (it'll ask you to authorize — that's normal, click through it since this is your own script). That's the only time you need to run anything manually.
9. Click **Deploy → New deployment**.
   - Click the gear next to "Select type" and choose **Web app**.
   - Description: anything, e.g. "v1".
   - Execute as: **Me**.
   - Who has access: **Anyone** (this lets the shared device use it without anyone signing into a Google account — appropriate since it's one shared screen).
   - Click **Deploy**, then **Authorize access** and click through the Google warning (it's your own script, this is expected for personal Apps Script projects).
10. Copy the **Web app URL** it gives you. That's your Staff page link. Open it once to confirm it loads and shows today's date with Miw/Nureen/Richa.
11. Bookmark that URL on the shared device (or add it to the home screen if it's a tablet/phone) so staff never have to type it.
12. The **admin link** is the same URL with `?page=admin` added at the end, e.g. `https://script.google.com/macros/s/XXXXXXXX/exec?page=admin`. Bookmark that separately somewhere only the admin uses.

## Using it day to day

- **Staff:** open the Staff page link (or it's already sitting there on the shared device), tap your name's "Log in" when you arrive, "Log out" when you leave, and "Mark done" on each task as you finish it (only shows up once you're assigned it and logged in). That's the whole interaction.
- **Admin — tasks:** each morning, open the admin link, enter the PIN, add today's tasks and tick who's doing what, click **Save tasks**. Use **Copy yesterday's tasks** if today looks the same — one click instead of re-typing everything (this copies the task list and assignments, but never copies over yesterday's "done" times, so today always starts fresh).
- **Admin — employees:** the admin page has an **Employees** section — type a new name into the box next to anyone's row to rename them, tap **+ Add employee** to bring on a new hire (up to 8 active at once), or tap the **✕** next to someone to remove them. Removing someone doesn't delete anything — their past attendance stays exactly as it was in the Sheet, they just stop showing up on the Staff page and task list from today onward. A rename or add/remove takes a few seconds to apply because it quietly updates every remaining day-tab in the current month so the roster stays consistent for the rest of it (already-past days are left untouched, since that's history).
- **Data:** the admin page has an **"Open this month's Sheet"** button that jumps straight to today's tab — no need to dig through Drive. (If you ever do want to browse Drive directly, look for a folder called "Cafe Tan90 Attendance Sheets" with one spreadsheet per month, 28-31 tabs each.) Nothing to export — it's already there, live, as soon as someone taps a button.
- **Staff task table:** tasks are listed as rows, with the active employees as columns across the top (better than the other way around once there are more than a couple of tasks — nothing to scroll sideways through). Each person's column shows their present/not-in-yet status in the header, and fades until they've logged in.

## If you ever want to change something later

- **Add/remove/rename employees:** no code changes needed — do this from the Admin page's Employees section (see above).
- **Change the PIN:** edit and re-run `setAdminPin` in the script editor — no redeploy needed.
- **More than 8 tasks a day, or more than 8 employees:** raise `MAX_TASK_COLUMNS` or `MAX_EMPLOYEES` in `Code.gs` (also widens the sheet layout — future day tabs pick it up automatically; existing tabs already built this month keep their old layout).
- **Every code change needs a new deployment version** to go live: Deploy → Manage deployments → pick the pencil/edit icon → Version: **New version** → Deploy. Just saving the file in the editor is not enough on its own.

## A couple of judgment calls I made — flag if you'd rather have it differently

- **"Who has access: Anyone"** means no Google login is required on the shared device — this matches "staff should think as little as possible," but it also means anyone with the URL could open it. If that's a concern, we can switch to "Anyone within [your Workspace domain]" if you're on Google Workspace, or add a lightweight PIN gate to the Staff page too.
- **Employee cap is 8** — plenty of headroom for a small cafe team, but if you ever expect to grow past that, say so and I'll raise the limit before you deploy.
- **Removing an employee mid-shift** (after they've already logged in today) will still blank their name on today's tab going forward from that point — their already-logged in/out times for today stay in the row, just without a name label next to them until you look at row position. This is a rare edge case (removing someone typically happens when they've left for good, not mid-shift), but flagging it in case it matters to you.
- **Present Today never un-checks on log-out** — once someone logs in, they're marked present for the rest of the day even after logging out. That's usually what "present today" should mean, but say the word if you want it to reset.
- **Log in/out and Mark done are all one-shot per day** — the buttons disable themselves once used, so nobody can double-tap and overwrite a time by accident. If someone needs to correct a mistake, that's a quick manual edit directly in the Google Sheet (open the day's tab, fix the cell).
- **A task can only be marked done if that person is assigned it and already logged in** — enforced on the server too, not just hidden in the UI, so it can't be bypassed.
