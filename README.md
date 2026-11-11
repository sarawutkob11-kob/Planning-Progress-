# Planning Progress: Parts 1–6 + multi-project (core data, alerts, 2-way sync, desktop app, roles, weekly report)

This part reads 3 sheets of `UAT Planning.xlsx` (UAT Plan, Test Progress, Defect List) and cleans up the values. It then stores them in a local SQLite database along with an append-only audit log, a PMO review queue and data-quality flags.

It only reads the workbook. It copies the workbook to `snapshots/` first and never opens or writes the original.

## Setup (once)

**Python 3.14** (tested with 3.14.6 on Windows 11). `pyproject.toml` allows 3.11+, but only 3.14.6 is tested.

```powershell
python -m venv $env:USERPROFILE\.venvs\uat-tracker          # outside OneDrive on purpose
$py = "$env:USERPROFILE\.venvs\uat-tracker\Scripts\python.exe"
& $py -m pip install -r requirements.lock                   # the exact, tested versions
& $py -m pip install -e . --no-deps                         # the app itself
```
`requirements.lock` holds the exact versions of the working environment (tests included). `pyproject.toml` only
sets ranges (up to the next major version). After upgrading a package and running the tests, regenerate it:
`& $py -m pip freeze --exclude-editable > requirements.lock` (keep the 3 comment lines at the top).

Then edit `config.toml` (copied from `config.example.toml`): workbook path, data folder and date window.

## Run

```powershell
$py = "$env:USERPROFILE\.venvs\uat-tracker\Scripts\python.exe"
& $py -m uat_tracker import     # copy + import + summary
& $py -m uat_tracker review     # open PMO review items (values the app refused to guess)
& $py -m uat_tracker flags      # active data flags
& $py -m uat_tracker summary
& $py -m pytest                 # unit tests
```

App data (phase 1): `%LOCALAPPDATA%\PlanningProgress\projects\<id>\` → `tracker.db`, `snapshots\` (one folder per project; see *Multi-project add-on* at the end).

## Rules implemented

| Rule | Behaviour |
|---|---|
| Header mapping | Columns are found by the first line of the header text, case-insensitive. The table stops at the first blank header, which keeps the legend table out. If a header appears twice, the leftmost one is used. A blank header in Defect col A counts as `No.` |
| Year > 2500 | Subtract 543 (Buddhist Era → CE). Applies to both text and real Excel dates. |
| Multi-date cell | The last value typed is the current date, and the earlier values become reschedule history in typed order. Separators: `=>`, `->`, spaces, new lines. |
| Text in a date cell | Text in parentheses and other words go to `*_note`. |
| Day/month | If one part is > 12, the reading is clear. If both are ≤ 12, a zero-padded `DD/MM/YYYY` is read as day/month, and anything else goes to review as ambiguous. |
| Date window | A date outside `date_window_start..end` goes to review. This catches `29/09/2027`, `1/4/1900` and dates that Excel already swapped. |
| Typos | `24/09/206`, `06-Now-2026`, `30/102026` → review. A bad value in the history goes to review on its own, and the valid current date is kept. |
| Time columns | A date found in Start/Finish time goes to review and stays where it is. |
| Status | `Not Start`, `In progress`, `Blocked`, `Complete`. Blank = no status. Unknown → review. Defects use `Open`, `Assigned`, `Fixed`, `Retest`, `Close`, `Rejected`. |
| Flags | Pass+Fail ≠ Executed, blank #Fail, missing owner (UAT Responsible, Defect Assigned To), missing Report By Email, Complete without Actual Finish (Test Progress). |
| IDs | UAT `{Group}-{TaskNo}` (repeats get `b`, `c`…, blanks become `x1`, `x2`…), `TP-{No.}`, `DEF-{No.}`. Rows are bound by content until Part 4 adds an ID column. |
| Audit | One entry when an item is created (all fields), then one per changed field. The log can't be updated or deleted, and each entry carries a hash of the previous one. A row that disappears from Excel is marked inactive and is never deleted. |
| Re-import | An unchanged file writes nothing, and an identical snapshot is reused. |

Dates are stored as ISO (`YYYY-MM-DD`) so they sort correctly. They're shown as `DD-Mon-YYYY` everywhere else, including the audit log and CLI output.

---

# Part 2: Traffic Light + email alerts

Only **Test Progress** drives automatic emails. The job re-imports Excel before every run.

**Dry-run is ON by default** (`[alerts] dry_run = true`). In dry-run every email goes to `pmo_email` only. The subject starts with `[DRY RUN]` and a banner shows the real To/CC. An `.html` copy is saved in `<data_dir>\previews\<date>\`. With no `pmo_email` set, only the previews are written. Set `dry_run = false` only when the PMO decides to go live.

## Before the first run

1. `pip install -e ".[test]"` again (adds `pywin32`).
2. In `config.toml`, `[alerts]`: set `pmo_email`, and `send_account` if Outlook has several accounts.
3. In the workbook, add **`Report By Email`** (right after Report By, one address) and **`CC Email`** (`;`-separated). Until then every email is skipped and flagged `missing_report_email`.
4. Bank holidays. Without data for the current year the job sends nothing and warns the PMO. Either:
   - register a key at <https://portal.api.bot.or.th/> and set `bot_api_key` (check `bot_url` there), then `holidays fetch`; or
   - `holidays import holidays.csv` (`date,name` per line, date `YYYY-MM-DD` or `DD-Mon-YYYY`), or `holidays add 2026-12-07 "Name"`.
5. Schedule it: `scheduler install` (creates the task *Planning Progress - Daily Alerts - <project>*: 08:30 Mon–Fri, only when you are logged on, missed runs are not caught up). `scheduler show` / `scheduler remove`. Once the new tasks are created, tasks with the old names (*UAT Tracker - Daily Alerts / Weekly Report* and the single-project *Planning Progress - Daily Alerts / Weekly Report*) are deleted.

## Commands

```powershell
& $py -m uat_tracker run-daily                  # run the job now (Outlook must be running)
& $py -m uat_tracker run-daily --preview-only   # dry-run without Outlook: previews only
& $py -m uat_tracker run-daily --date 2026-11-06 --preview-only   # pretend it is that day (testing)
& $py -m uat_tracker lights --date 2026-11-06   # colors for every app (nothing saved)
& $py -m uat_tracker emails --date 2026-11-09   # email log
& $py -m uat_tracker jobs                       # job runs, skips and failures
& $py -m uat_tracker holidays list|fetch|add|import|remove
& $py -m uat_tracker upload TP-38 "C:\path\results.pdf"
& $py -m uat_tracker remind B-6 --to someone@company.com [--cc x@company.com] [--note "..."]
& $py -m uat_tracker scheduler print|install|show|remove
```

## Rules implemented (approved 09-Nov-2026)

| Topic | Behaviour |
|---|---|
| Traffic Light | First match wins: Status Complete → Complete; past Planned Finish or %Fail > X → Red; %Exec < %Time or (#Fail ≥ 1 and %Fail ≤ X) → Amber; else Green. X = `fail_threshold` (5%). %Time uses calendar days; Planned Start = Planned Finish → %Time 1 from that day. Blank #Fail and blank Executed = 0. Total TC blank/0 or a planned date blank (or waiting in review) → no color, flag `missing_data`. |
| Color history | A row in `traffic_light_history` + an audit entry (channel `system`) only when the color changes. The Excel Traffic Light column is not changed (write-back is Part 3). |
| Workdays | Mon–Fri minus cached holidays. A trigger date on a non-workday moves to the previous workday. |
| R1 Planned Start | On the start day: one "starts today" email. From the next workday, daily while Actual Start is empty. |
| R2 Planned Finish | On the finish day: one reminder (not sent if Actual Finish is already filled). From the next workday, daily while Actual Finish is empty. Actual Finish > Planned Finish → flag `actual_finish_late`. |
| R3 color change | One email on a change into Amber/Red (first color, Green→Amber/Red, Amber→Red, re-entry after Green). **Red→Amber: no email.** A data gap (no color) in between does not count as a change. A failed send is retried next day while the color holds. |
| R4 close-out | Executed = Total TC and Actual Finish filled (also when later than Planned Finish). R4A: daily "set Status to Complete" (the app never changes Status). R4B: daily "upload results" until TC/Result Complete is Complete or a file is uploaded. |
| Status Complete | Stops every rule for that item. |
| Recipients | To = Report By Email; CC = CC Email. Empty Report By Email → not sent, logged as skipped, flagged. |
| Dedupe | One email per rule + item + day (separately for dry-run and live, so dry-run rows never block live sends). |
| Going live | When the first live run happens, colors recorded during dry-run count as already notified; only new changes email. |
| Job safety | A day that ran OK is not re-run (`--force` re-runs; dedupe still applies). Missed workdays are logged as skipped, never backfilled. A scheduled start after 09:30 is skipped. Outlook not running → day skipped. Re-import fails → nothing sent, PMO warned. |
| Holidays | BOT API, fetched in Monday's run for this year + next, and whenever the current year has no data. Fetch fails → cache kept, PMO warned. No data for the current year → nothing sent, PMO warned. PMO warnings go to `pmo_email` in both modes. |
| Uploads | `<data_dir>\attachments\<item-id>\`, types pdf xlsx xls csv docx doc pptx png jpg zip msg, 25 MB max, hash-verified, never deleted. Sets TC/Result Complete = Complete + audit. The value waits in `pending_writeback` so re-import does not undo it; Part 3 writes it to Excel. |
| Manual reminder | UAT Plan / Defect List items only, recipient given with `--to` (those sheets hold names, not addresses). Respects dry-run. |
| Logs | Every sent / dry-run / skipped / failed email → `email_log` + audit log. |

---

# Part 3: 2-way sync, backups, Conflict Center

The app and the workbook now share one dataset. Excel changes are read in automatically. App changes (edits, review fixes, conflict decisions, uploads, the calculated Traffic Light) are written to Excel.

## How writes work
- **Through Excel itself (COM).** Images, the sensitivity label, formulas, formatting and the other 7 sheets stay exactly as they are. Only changed cells in the 3 sheets are touched, and formula cells are never overwritten.
- **Workbook open in Excel on this PC.** The app writes into that open copy if Excel lets it in. Otherwise it **waits until you close the workbook**: a OneDrive file open in Excel isn't locked, and saving underneath it would fight your session.
  - With **AutoSave on** (approved 09-Nov-2026) the app writes even when Excel shows unsaved changes: AutoSave uploads them anyway, and formulas like TODAY() keep the file "unsaved" after opening. With AutoSave off it waits for a save. Tested live on the real file: a change appeared in the open window in about 3 s.
  - The file is read with full sharing (read, write, delete), because Excel with AutoSave refuses a plain read.
  - If Excel refuses a write for now (a cell being typed in, a dialog open), the write is retried; nothing is lost.
- **Batching.** Changes made within 10 s are saved together. A backup copy goes to `<data_dir>\backups\` before every write; backups are never deleted.
- **Hidden `App ID` column.** Added right after the last header of each sheet, so rows stay bound when people insert or sort. In Test Progress the Traffic Light legend moves one column right.
- **Date format.** One date becomes a real Excel date shown `DD-Mon-YYYY`. A rescheduled date becomes text `13-Nov-2026 => 20-Nov-2026` (last = current). Bad parts nobody has fixed yet stay as typed.

## Commands

```powershell
& $py -m uat_tracker sync init-ids             # first time: add the hidden App ID column
& $py -m uat_tracker sync watch                # keep syncing (Ctrl+C to stop); Part 4 runs this inside the app
& $py -m uat_tracker sync once                 # read Excel, then write queued changes now
& $py -m uat_tracker sync status | log | unblock
& $py -m uat_tracker edit tp TP-38 planned_finish 2026-11-20 --reason "..." --root-cause "Waiting other system data"
& $py -m uat_tracker review                    # PMO review queue, then:
& $py -m uat_tracker resolve-review 17 2026-09-24 --reason "typo"   # empty value = drop that part of the cell
& $py -m uat_tracker conflicts list
& $py -m uat_tracker conflicts resolve 1 excel|app|new [--value "..."] --reason "..."
& $py -m uat_tracker backup list | now | restore 3 --yes
```

`run-daily` now also writes the calculated Traffic Light and any queued changes to Excel at the end (not with `--preview-only`).

## Rules implemented (approved 09-Nov-2026)

| Topic | Behaviour |
|---|---|
| Excel → app | A file watcher waits for 5 s of quiet, with a 60 s check as a fallback, then runs the Part 1 import with channel `excel_sync`. The editor is "Edited in Excel (last saved by <name>)", taken from the file's last-saved-by field as a hint. An unchanged file is skipped. |
| Conflict | The same cell changed in Excel and in the app since the last sync, even to the same value. The app keeps its value, Excel keeps its value, and nothing is written for that cell until the PMO chooses Excel / app / new value (reason required). The PMO is emailed. |
| Excel edit Actual Finish > Planned Finish | Conflict. The app keeps the old Actual Finish until resolved. |
| App edit rules | A planned date that moves needs a reason (root cause from the 5 standard ones). Test Progress: Actual Finish after Planned Finish is refused unless Planned Finish moves in the same edit, and Status Complete needs Actual Finish. Standard statuses only; Report By Email is one valid address; CC Email is `;`-separated valid addresses. A field with an open conflict can't be edited until the conflict is resolved. Dates (edits and review fixes) must be inside the project date window. |
| Restore | Refused while the workbook is open in Excel on this PC. |
| Locked / busy | Retry at 1, 2, 5, 10, 15 and 30 min. Then status = error and one PMO email per episode. The queue stays, and the next file change tries again. |
| OneDrive conflict copy | `UAT Planning-<PC>.xlsx` stops all writes and emails the PMO. `sync unblock` resumes after you've dealt with it. |
| Review queue | The PMO's value is written back clean. Other unresolved parts of the same cell stay as typed. |
| Traffic Light | The calculated color is written to Excel's Traffic Light column. Missing data leaves the cell alone. Manual values are overwritten; the audit log keeps the old ones. |
| Audit | App edits (`app`), Excel changes (`excel_sync`), each write to Excel ("Written to Excel"), conflicts and their resolution, and restores (`restore`). |
| Restore | Takes a backup of the current file first, replaces the workbook, then reads it back in. |

---

# Part 4: Desktop app (PySide6)

Open it with the **Planning Progress** shortcut on the Desktop, or from a terminal in this folder: `.\run.ps1` (app, synced with Excel), `.\run.ps1 -NoSync` (app, no writes to Excel), `.\run.ps1 <command>` for any other command (e.g. `.\run.ps1 sync status`). `run.cmd` does the same where PowerShell scripts are blocked. While the app is open it watches the Excel file: Excel edits appear on the open screen without a restart, and app edits are written to Excel within about 10 s. When the workbook is open in Excel on this PC, writes wait until it is closed. Closing the app stops the watching; the 08:30 job runs through Task Scheduler as before.

| Screen | What it does |
|---|---|
| Dashboard | 4 KPI cards, follow-up list (Overdue / Today / No target / Red / Amber / Missing data) with **Send reminder** for UAT Plan and Defect items, incomplete apps, due within 2 weeks, missing Report By Email warning. "Generate Weekly Report" is disabled until Part 6. |
| UAT Plan | Import from Excel, + Add task (the row is inserted after the last task of its group), filter chips, owner filter, groups A–E (fully complete groups collapsed), detail panel with linked Test Progress, reschedule timeline, linked defects, edit history, Update. |
| Update dialog | Status, dates (picker only), remark; a moved planned date shows "moved A → B (reschedule #N)" and needs a reason + root cause; Test Progress: Complete needs Actual Finish, Actual Finish after Planned Finish is blocked. |
| Test Progress | Summary, filters, progress bars, Pass+Fail mismatch tag, "Not filled" Actual Finish, calculated Traffic Light; Update results (numbers, dates, Report By Email, CC Email); Upload result file. |
| Defect List | Board (drag a card, or "Move to", = status change, logged) / Table; + Add Defect (owner + fix target date required). |
| Edit History | Search, page, editor, date range, type; old value struck through → new; Restore (saved as a new change); Export to Excel. |
| Conflict Center | Conflicts tab: Excel value vs app value with edit times; Use Excel / Use app / Enter new value, reason required; neither side is overwritten until chosen. PMO review queue tab: the cell as typed, the problem part, one-click suggestions (day/month swapped, both readings of an ambiguous date, typed year fixed, month-name typo, remove this part, clear the time cell, closest status) or pick/enter the value; the clean value is written back to Excel. Suggestions are only values inside the project date window; the PMO always chooses. |
| Settings | Excel file, Fail X%, daily email time (+ apply to Windows scheduler), dry-run (asks before turning it OFF), PMO email, BOT holidays status + fetch, backups + restore, email log, email templates; Claude API key (Part 6) and users/roles (Part 5) are placeholders. Saved in the app database; they also apply to the CLI and the scheduled job. |

Rules decided in this part (approved 09-Nov-2026): Send reminder only for UAT Plan / Defect items; the recipient email is remembered per owner name; Open tasks = UAT Plan not Complete; Overdue = finish date before the As-of date and not Complete; Due today counts UAT Plan and Test Progress; the As-of date only changes what is shown; a UAT Plan task is linked to a Test Progress app when the app name appears as a word in the task name, and defects to that app by System/App.

Font: IBM Plex Sans Thai when installed, otherwise Tahoma.

---

# Part 5 (v2): User profile and roles

No login, no OTP, no email verification (approved 10-Nov-2026; replaces the earlier OTP login).

## Using the app
1. **First start on a PC:** a form asks for First name, Last name and Email. A wrong email format can't be saved.
2. The profile is saved for your Windows user (`%LOCALAPPDATA%\PlanningProgress\profile.json`), so the app opens straight in next time.
3. **The very first profile ever** becomes the **Admin**, after a confirmation. This happens once only. Later new emails start as **QA**; the Admin changes roles in Settings.
4. The top bar shows "First Last · Role" and a **Switch user** button that opens the same form.
5. If you type an email that already exists, its saved name is used. A deactivated email can't be used.

**Command line:** commands that change data run as the profile saved on this PC by the app (open the app once first). These are `edit`, `resolve-review`, `conflicts resolve`, `backup restore`, `upload`, `remind`, `holidays fetch|add|import|remove`, `scheduler install|remove` and `sync init-ids|unblock`. Read-only commands, `import`, `sync once|watch` and `run-daily` need no profile.

All email (daily alerts, reminders) is sent by Outlook on this PC, as in Part 2.

## Rules implemented (approved 10-Nov-2026)

| Topic | Behaviour |
|---|---|
| Roles | **Admin**: everything + users. **PMO** and **User BU/PM**: edit all, Conflict Center + review queue, restore, reminders, Settings (not users). **QA**: change only their own rows, upload only to their own Test Progress rows, add defects; everything else read-only. |
| Own row (QA) | Test Progress: Report By Email = profile email. UAT Plan Responsible / Defect Assigned To: one of the user's **Names in Excel** (case and extra spaces ignored). "First Last" from the form is added automatically; the Admin adds others. A cell that lists several people (new line, `/`, `,`, `;`) counts for each of them. A defect created by a QA user but assigned to someone else is not theirs. |
| Enforcement | Checked in the services (app and CLI). The role is read from the database on every action, so a change applies at once. The UI hides what a role can't do, and rows that aren't yours open read-only. A refused action shows a message and is logged (`denied`). |
| Users (Settings, Admin only) | First name, last name, email, role, active, last used, names in Excel. Add, edit, deactivate / reactivate. Never deleted. The last active Admin can't be demoted or deactivated. **Import from Test Progress** lists Report By Email / CC Email addresses that aren't users yet (role QA by default). Nothing is saved until you tick and Save. |
| Last used | Updated when the app opens with that profile or on Switch user. |
| Audit | Every audit and email log entry shows the editor as **"First Last <email>"**. Logged: profile created, switch user, role change, user added, deactivated / reactivated, names changed, denied. These have type "Users / profile" in Edit History. Older entries keep what they had: the log can't be changed. |

**Known limit (accepted for Phase 1, single user):** there is no verification. Anyone using the PC can type any email and gets that email's role, including the Admin's. Lock Windows (Win+L) when you leave the PC. Revisit for Phase 2 (multi-user).

---

# Part 6: AI weekly report

Every Monday at 09:00 the app sends a weekly UAT report. The tracker computes every number and table; Claude (Anthropic API) writes only the executive summary, top risks and suggested follow-ups, in Thai and English.

**Before going live:** the report sends project data (application / task / defect names, dates, reasons, numbers; no people's names or email addresses) to Anthropic. Confirm that company policy allows it. Without an API key the report still goes out, just without the AI part.

## Setup
1. `pip install -e ".[test]"` again (adds the `anthropic` package).
2. Settings: paste the **Claude API key** (kept only in Windows Credential Manager as `PlanningProgress/anthropic`), and check the **Claude model**. Which model is used: Settings > Claude model if filled in, else `[weekly_ai] model` in `config.toml`, else `claude-opus-5-5`. Leave the Settings field blank to follow `config.toml`. An unknown or retired model name does not stop the report: it is made without the AI part and says why. `[weekly_ai] server_fallback` (default `true`) lets the API switch model when one is overloaded; `[weekly_ai.prices]` sets the cost shown per report (model without a price: cost `-`). From a terminal: `.\run.ps1 ai-key set` / `ai-key status` / `ai-key clear`.
3. Settings: **Apply to Windows scheduler** now installs both tasks: *Planning Progress - Daily Alerts - <project>* and *Planning Progress - Weekly Report - <project>* (09:00 Mon-Fri; the job decides the day).
4. Add **Report By Email** / **CC Email** to Test Progress, set the **PMO email** and load the **bank holidays** (see Part 2). Until then the report goes to the PMO only.

## Use
- Dashboard: **Generate Weekly Report** → preview → **Send** / **Cancel** (Admin, PMO, User BU/PM).
- Left menu: **Weekly Reports**: every report kept, viewable as it was sent.
- Terminal: `.\run.ps1 run-weekly [--force] [--date 2026-11-16] [--preview-only] [--no-ai]`, `.\run.ps1 weekly list`, `.\run.ps1 weekly show 3`.

## Rules implemented (approved 10-Nov-2026)

| Topic | Behaviour |
|---|---|
| Day | Monday 09:00. Monday is a bank holiday → the first workday of that week. Once per week. PC or Outlook off → skipped and logged, never sent later. No holiday data for the year → not sent, PMO warned. |
| Data | Excel is re-read first. Overall Executed/Total %, Pass, Fail; Red/Amber apps with the reason (overdue / fail rate / behind time); overdue UAT Plan tasks and Test Progress apps; open defects, past fix target, no fix target; due in the next 2 weeks; reschedules in the last 7 days (reason when changed in the app, "not given" when changed in Excel); open conflicts and review-queue count. Same definitions as the Dashboard. |
| AI | Model from Settings, effort medium, JSON schema output, refusal fallback `"default"` (Opus / Sonnet / Fable models). Timeout 60 s, 2 retries. Only the computed data is sent: no people, emails, files or paths. |
| Number check | Every number in the AI text must appear in the data sent. Otherwise the AI part is dropped. Names under "related" that aren't in the data are removed. |
| AI unavailable | No key, API down, refusal, cut-off, bad JSON or an unknown number → the report goes out **without** the AI part (with a note), and the PMO is warned. |
| Recipients | All Report By Email + CC Email in Test Progress, deduped. Invalid addresses are skipped and flagged (`invalid_report_email`). None at all → the PMO only, plus a warning. Dry-run → the PMO only (the real To is shown in the banner). |
| Format | One HTML email, Thai first then English, no attachment. Tables and numbers rendered by code. |
| Kept | `weekly_report` table: date, automatic/manual, who, recipients, invalid addresses, HTML, data, AI on/off, model, tokens, cost (USD), status (preview / sent / dry_run / failed / skipped / cancelled). Also in the email log and the audit log. |
| Cost | About 9,000 input + 1,500-2,500 output tokens per run: ≈ US$0.07 with `claude-opus-5-5` ($4 / $20 per million tokens). Logged per run. |

# Multi-project add-on (approved 10-Nov-2026)

The app is called **Planning Progress** everywhere (window, shortcut, scheduled tasks, email subjects
`[Planning Progress]`, logs). The internal Python package keeps the name `uat_tracker`.

## Layout
```
%LOCALAPPDATA%\PlanningProgress\
  app.db              projects, bank holidays, shared settings (PMO email, language, Claude model)
  profile.json        who uses this PC (email, name, last project)
  projects\p001\      tracker.db + attachments, backups, snapshots, previews of ONE project
  clear-project.log   "Clear project data" log;  clear-backups\  zip made before each clear
  migration-backup\   backups taken by the one-time migration
  logs\app.log        the app and CLI commands (errors with full details)
  logs\jobs.log       scheduled runs (daily alerts, weekly report: --scheduled)
```
**Logs**: each file rotates at 5 MB, 5 old files kept (`app.log.1` ... `app.log.5`). Every line names the open
project (`[p002:Test]`). Send `logs\` when reporting a problem. The Claude API key is never written there.

**Databases** use SQLite WAL mode, so the app and a scheduled job can use the same project at the same time;
you will see `tracker.db-wal` / `tracker.db-shm` next to `tracker.db` while it is open. Copy a database only with
the app closed (or use the backups the app makes); copying `tracker.db` alone while it is open can miss the last
changes.

**Long actions** (Import from Excel, re-reading Excel after a conflict decision, restoring a workbook backup,
Relink, Clear project data) run in the background behind a "please wait" box; the window stays responsive. If
the sync is writing the workbook at that moment: "Sync running, try again in a moment".

**Excel busy** (a cell being edited, a dialog open in Excel): each Excel call is retried for up to ~30 s
(5 s when switching project), then the change waits in the normal retry queue.
1 project = 1 Excel file. Users + roles, Fail X%, dry-run, daily send time, date window, weekly reports, sync,
conflicts, history and attachments are per project. Shared: Claude API key (Credential Manager), holidays,
PMO email, Outlook account, language.

## Using it
- **Project drop-down** under the app name: every page shows only the selected project; the last one is
  remembered. A user sees only the projects where they have a role.
- **+ New project** (Admin / PMO of any project; anyone before the first project): Project Name, then pick the
  folder. File = `Planning Progress <name>.xlsx`; `/ \ : * ? " < > |` become `-`. Blocked if that file exists
  in the folder or another project uses that file name. The file has the 3 sheets with the headers / formats of
  the old workbook (template in code: `resources/project_template.json`, rebuilt with `tools/build_template.py`)
  plus Report By Email / CC Email. The creator is the project's Admin.
- **Import Excel**: an existing .xlsx with the 3 sheets and required headers (missing ones are listed). Never
  renamed; other sheets never touched. Name = file name without `.xlsx` and without `Planning Progress `.
- **+ Add task** has "+ New group…" (next free letter). **+ Add Application** (Test Progress) adds a row after
  the last application with the `%Progress` formula. Rows are written to that project's file only.
- **New email** in the profile form: QA in each project whose Test Progress lists it (Report By Email / CC Email);
  otherwise "ask the Admin".
- **Settings**: Project card (file + Relink, date window, Archive, archived list), "This project only",
  "Shared" (Admin / PMO), Apply to Windows scheduler (one task pair per project, `--project <id>`).
- **File moved / renamed outside the app**: banner + **Relink** (the new file needs the 3 sheets and >= 80 % of
  the project's App IDs).
- **Danger zone → Clear project data** (TEMPORARY, Admin / PMO): see `services/clear_project.py`. **Off by
  default**; turn on for testing with `[features] enable_clear_project = true`. Restore:
  `.\run.ps1 project restore-clear <project> <zip>`.
  - The **audit log is never cleared**: one entry "project data cleared by <user>, backup <zip>" is added
    (and "project data restored from <zip> ..." after a restore).
  - Feature on + live mode (`dry_run = false` for the open project): red warning at the top of Settings and a
    warning in the log each time the project is opened.

### Removing the Clear project feature
Everything is isolated; delete these (line numbers as of the end of Hardening phase 1, 11-Nov-2026; search for
the names if lines moved):

| File | What to delete |
|---|---|
| `src/uat_tracker/services/clear_project.py` | the whole file |
| `src/uat_tracker/services/projects.py` | lines 112-113 (`warn_if_live` import + call in `open_project`) |
| `src/uat_tracker/ui/screens/settings.py` | `clear_project` in the import on line 15; lines 56-57 (live warning); lines 164-165 (danger zone); methods `_live_warning` (line 242) and `_danger_zone` (line 248, up to `_browse`) |
| `src/uat_tracker/ui/projects.py` | `DANGER_QSS` (line 20) and class `ClearProjectDialog` (line 163, up to `class NoProjectWindow`) |
| `src/uat_tracker/cli_project.py` | the `restore-clear` sub-command: parser line 36, branch from line 111, `ClearError` in lines 119-120; the word in the help/docstrings (lines 1, 18) and `cli.py` line 13 |
| `src/uat_tracker/ui/i18n.py` | all `danger.*` keys (lines 476-501) and `work.clearing` |
| `src/uat_tracker/config.py` | field `enable_clear_project` (line 76) and its loader line 201 |
| `config.example.toml` | the `[features]` section |
| `tests/test_projects.py` | `test_clear_project_*`, `test_import_from_excel_after_clear_unpauses`, imports on lines 23-24 |
| `tests/test_ui.py` | `test_settings_red_warning_only_when_clear_feature_on_in_live_mode` |

The "reading Excel is paused" state (`S_PAUSED`, the `proj.paused` banner) can stay: it only switches on after a clear.

## Known limits
- **Who is using the app is not verified.** The name / email in the profile form is a claimed identity: anyone
  on the PC can "Switch user" and act with that person's role. Roles and the audit log rely on people using
  their own name. Real login (Windows / AD identity) is planned for phase 2.
- RAMC UAT's Test Progress has a *Grand Total* row: a new application is inserted above it, but the existing
  `SUM(D2:D52)` formulas do not grow to include the new row (Excel does not extend a range for a row inserted
  right below it). Fix the totals once in Excel if you add applications from the app.

## Commands
```powershell
.\run.ps1 project list | create "<name>" --folder <dir> | import <xlsx> | relink <project> <xlsx>
.\run.ps1 project archive|unarchive <project>
.\run.ps1 project migrate [--yes]          # one-time: UATTracker data -> project "RAMC UAT" (dry run without --yes)
.\run.ps1 --project "RAMC UAT" summary     # any other command works on one project
```
