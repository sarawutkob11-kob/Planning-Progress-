# Build: Progress Planning (UAT Tracker) desktop app

## 0. Rules
- Role: senior Python dev, Windows desktop app.
- Anything unspecified → STOP, ask user. Never decide alone (open items: §14).
- Step 1: reply with architecture, libs, folder layout, open questions. Wait approval before code.
- Then milestones (§13) one by one; each ends with summary + how to test + wait confirm.
- Never touch source Excel without backup. Never delete data. Test on copy first.

## 1. Context
- User = PMO, project RAMC UAT. Tracks task status, test results, defects across teams.
- Source: `UAT Planning.xlsx`, OneDrive (company), synced locally (visible in File Explorer).
- Current pain: mixed BE/CE years + swapped d/m, reschedules typed as `=>` in one cell, inconsistent status spelling, missing data, manual Traffic Light, no edit trail.
- Goal: desktop app sharing ONE dataset with Excel (2-way real-time sync), auto status calc, email alerts, full edit history.

## 2. Scope
- Phase 1 (build now): PMO only, own PC. Others still edit Excel. App data stored on PMO PC only.
- Phase 2 (design for, don't build): all users (QA, PM, User BU) install app, shared data via OneDrive. Storage TBD at phase 2.
- Storage layer location must be config-switchable without touching business logic.

## 3. Hard constraints
- Python, Windows desktop, runs on PMO PC.
- NO Azure App Registration → NO Microsoft Graph, NO MSAL.
- Excel read/write via local OneDrive-synced file (OneDrive uploads itself).
- Email via running Outlook desktop, PMO account (e.g. pywin32 COM).
- Excel writes: preserve formatting, formulas, all other sheets. Touch only 3 sheets, only changed cells.
- Scheduled jobs (08:30 workdays, Mon 09:00) run only if PC + Outlook on. PC off at that time → skip run, no backfill, log skip.
- UI framework: propose, user picks (§14).

## 4. Login
1. Read email of Outlook account open on PC.
2. Send OTP to it via Outlook.
3. User enters OTP → in.
4. Check role from in-app user table (§11).
OTP expiry/attempts/session length: ask (§14).

## 5. Source: 3 sheets
Map columns by header name, never position (user may add columns). Never touch other sheets (NLeadAccount, Main Test Acc, Payment Txn...).

**UAT Plan**: Task no, Sub task, Group, Task No., Task name, Note, File Path, Responsible (Mandatory), Status, Start date, Start time, Finish date, Finish time, Remark.
- Groups A–E (group header = row with long text in Task no), e.g. `A: SETUP for UAT Testing`, `E: Downstream Systems / Report`.
- Task numbers duplicate in source → app makes unique internal ID (e.g. `E-23`) bound stably to Excel row (binding method: propose, ask).

**Test Progress** (drives all emails): No., Application, Report By, **Report By Email** (new, after Report By, ONE address), **CC Email** (new, multiple, `;`-separated), Total TC, No. of TC Execution, %Progress, #Pass, #Fail, Planned Start, Planned Finish, Actual Start, Actual Finish, Traffic Light, TC/Result Complete, Status, Remark.
- Last row Grand Total = sum, not an app.

**Defect List**: No., Date, System/App., Issues/detail, Status, Severity, Assigned To, Fix target date, Data test, Raised by, Remark.

## 6. Parsing rules (import + every Excel read)
1. All dates CE. Year > 2500 → −543 (`9/14/2569` → 14-Sep-2026).
2. Multi-date cell → LAST TYPED value = current, regardless earlier/later. Prior values = reschedule history. Handles `=>`, `->`, or space-separated w/o arrow.
   - `10/09/2026 => 19/09/2026 => 15/09/2026` → 15-Sep-2026
   - `20/10/2026 25/9/2026 => 17/09/2026 14/9/2026` → 14-Sep-2026
3. Non-date text in date cell (`(BD 30/09/26)`, `(actual start 25/09/26)`) → move to note.
4. Ambiguous (`10/8/2026`) or typo (`24/09/206`, `29/09/2027`, `1/4/1900`, `06-Now-2026`) → NO guessing. Put in "PMO review" queue, show raw, PMO picks/enters.
5. Write back to Excel as `DD-Mon-YYYY`.
6. Reschedule history stored BOTH: in app + in Excel same cell as `6-Nov-2026 => 13-Nov-2026` (last = current per rule 2).
7. Normalize status: `Not Start`, `In progress`, `Blocked`, `Complete`/`Completed` (complete/comlete/Inprogress → standard). Unknown → PMO review queue.

## 7. 2-way sync
- Real-time: file watcher on Excel. Excel change → load into app. App change → write Excel immediately.
- Only 3 sheets.
- Backup Excel on PMO PC before every write.
- File open/locked → queue + retry, show sync status in UI (retry policy: propose, ask).
- Conflict = same cell changed in both since last sync → PMO chooses. Notify PMO in app (Conflict Center) + email. Until resolved, overwrite neither side.
- Excel edit making Actual Finish > Planned Finish → conflict to PMO (app + email).
- Every sync change → Audit Log (§10). Unknown Excel editor → "Edited in Excel".

## 8. Email rules (Test Progress only)

**8.1 Common**
- From Outlook, PMO account.
- 08:30, workdays only (no Sat/Sun, no bank holidays).
- Bank holidays: fetch from Bank of Thailand (BOT) official financial-institution holiday source. Tell user source + fetch frequency before building. Fetch fail → in-app warning (fallback: ask).
- 1 email per item (per app per topic). No digest.
- To = Report By Email. CC = CC Email (empty → no CC). Report By Email empty → don't send, flag "missing email" on Dashboard.
- Trigger date on holiday → send previous workday.
- PC off at 08:30 → skip day.
- Log every sent/skipped/failed email (time, recipient, rule, item).

**8.2 Rule 1 Planned Start**
- Send once on Planned Start date: task starts today.
- From Planned Start on, Actual Start empty → daily (workdays) until filled.

**8.3 Rule 2 Planned/Actual Finish**
- On Planned Finish date: remind to fill Actual Finish.
- From Planned Finish on, Actual Finish empty → daily until filled.
- Actual Finish ≤ Planned Finish = valid.
- Actual Finish > Planned Finish in app = blocked until Planned Finish rescheduled + reason (saved to history). If done in Excel → conflict to PMO (§7).

**8.4 Rule 3 Traffic Light** (auto, no manual entry)
`%Exec = Executed/Total TC`, `%Time = (today − Planned Start)/(Planned Finish − Planned Start)`, `%Fail = #Fail/Executed`.
First match wins:
1. Status Complete → Complete
2. Past Planned Finish and not Complete, OR %Fail > X → Red
3. %Exec < %Time, OR #Fail ≥ 1 and %Fail ≤ X → Amber
4. Else → Green
X default 5% (UAT pass ≥ 95%), editable in Settings.
Email once per color change INTO Amber/Red: Green→Amber, Green→Red, Amber→Red, and re-entry after returning Green. Red→Amber: ask (§14).

**8.5 Rule 4 Close-out**
- Condition: Executed = Total TC (100%) AND Actual Finish filled (≤ Planned Finish).
- Status not Complete → email owner daily to change Status. System must NOT change Status itself.
- TC/Result Complete not Complete → email daily to upload test-result docs.
- Upload = attach file in app. On success → TC/Result Complete = Complete automatically + write to Excel.

**8.6 UAT Plan + Defect List**
- No auto email.
- Manual "Send reminder" button per item on Dashboard (phase 1), via PMO Outlook.

**8.7 AI Weekly Report**
- Auto every Monday 09:00 (Monday holiday: ask §14).
- To: all Report By Email + CC Email in Test Progress, deduped.
- Claude API, key entered in Settings, stored encrypted, never in logs/Excel.
- Content: overall %exec, Pass/Fail, Amber/Red items, overdue tasks, open + overdue defects, due in next 2 weeks, last week's reschedules + reasons.
- PMO can also generate from Dashboard; preview in app before auto send.

## 9. Screens (per approved UX/UI design)

**9.1 Shell + style**
- Left nav (dark navy `#14213D`): Dashboard, UAT Plan, Test Progress, Defect List, Edit History, + Conflict Center, Settings.
- Top bar: project name, search, As-of date, sync status, TH/EN toggle.
- Bg `#F4F6F9`; white cards, border `#E1E6EE`, radius 12px; primary `#1F4FBF`; Amber text `#7A4E00` bg `#FFF1D6`; Red text `#A3260F` bg `#FDE7E3`; Green/Complete `#1E6B3E` bg `#E3F2E8`.
- Font IBM Plex Sans Thai (fallback Tahoma). Controls ≥ 44px high. Dates always DD-Mon-YYYY.

**9.2 Dashboard (PMO)**
- 4 KPI cards: TC executed (e.g. 91% · 847/926), open tasks, due today, open defects (overdue / no target).
- "Follow up today": tag (Overdue, No target, Today, Amber, Missing data) + item + detail + owner + "Send reminder".
- Incomplete apps: progress bars.
- Due within 2 weeks: by date.
- "Generate Weekly Report" button.
- Missing Report By Email warnings.

**9.3 UAT Plan**
- Buttons: Import from Excel, + Add task.
- Filter chips: All, In progress, Not Start, No status, Due in 7 days, Rescheduled ≥2. + Owner filter.
- Table grouped A–E, collapsible. Cols: ID, Task, Owner, Status, Planned Finish (first), Current Finish (latest), reschedule count (≥2 red).
- Right detail panel on click: Note, Status, Owner, Current Start/Finish, linked Test Progress, reschedule timeline (date + change + reason), linked defects, link "View full edit history", "Update" button.

**9.4 Update dialog**
- Status dropdown; Start/Finish via date picker ONLY; Remark.
- Finish changed → warning "Finish moved A → B (reschedule #N)" + REQUIRED reason + root cause pick (Dev not done, Waiting other system data, Defect found, Environment not ready, Waiting user check).
- Completed → REQUIRED Actual Finish.
- Note: all edits auto-logged. Buttons Cancel / Save.

**9.5 Test Progress**
- Summary: Executed, Pass, Fail, Amber count.
- Filters: Needs follow-up, Incomplete data, Complete, All.
- Table: App, reporter, progress bar + Executed/Total, Pass/Fail (+ "Pass+Fail mismatch" tag when Pass+Fail ≠ Executed), Planned Finish, Actual Finish (empty → red "Not filled"), Traffic Light.
- Update results form + test-result file attach (§8.5).

**9.6 Defect List**
- Toggle Board/Table; + Add Defect.
- Board: Open → Assigned → Fixed → Retest → Close/Rejected, with counts.
- Card: #no · app, Severity tag (1–2 red), detail, owner, days open, Overdue / No target tag.
- Add form: REQUIRED individual owner + Fix target date.

**9.7 Edit History (Audit Log)**
- Filters: search, page (UAT Plan/Test Progress/Defect), editor, date range, type (Reschedule, Status change, Data edit, Import).
- Table: date, editor, type, item + page, field, old (strikethrough) → new, reason, "Restore" button.
- Export to Excel.

**9.8 Conflict Center** (added)
- List: item, field, Excel value, app value, edit time each side, choose buttons.
- PMO review queue (§6 unparseable dates, unknown status).
- Show mockup to user before building.

**9.9 Settings** (added)
- Excel path, Fail X%, Claude API key, email time (08:30) + report time (Mon 09:00), BOT holiday fetch status, users + roles, email templates.

## 10. Audit Log
- Log EVERY change from every channel (app, Excel sync, import, restore, system).
- Fields: time, editor, channel, sheet, item ID, field, old, new, reason.
- Immutable. Restore = new entry, never delete.
- Phase 1: audit log, conflicts, attachments, email log stored on PMO PC only.

## 11. Roles
| Role | Rights |
|---|---|
| Admin | All + user mgmt |
| PMO | Edit all, resolve conflicts, restore, send reminders |
| QA | Edit own tasks/results, upload files, add defects |
| User BU / PM | Same as PMO |
Phase 1 single user, but build roles for phase 2.

## 12. Acceptance (real data)
1. `20/10/2026 25/9/2026 => 17/09/2026 14/9/2026` (UAT Plan, CCMS run retrieve data from EDW) → Start 14-Sep-2026, history has 3 prior values.
2. `9/14/2569` → 14-Sep-2026.
3. `24/09/206`, `29/09/2027`, `1/4/1900`, `06-Now-2026` → PMO review queue, not auto-converted.
4. Test Progress ATM: Executed 4, Pass 2, Fail 0 → "Pass+Fail mismatch".
5. DMS-RAMC (126/186, Fail 2, plan 22-Oct → 13-Nov-2026) on 6-Nov-2026 → Amber, 1 email on entering Amber.
6. Edit Excel → app updates with no action; edit app → Excel updates; formatting + other sheets intact.
7. Same cell edited both sides → Conflict Center + PMO email; neither overwritten.
8. Planned Finish on Saturday → reminder sent Friday (or earlier workday if Friday bank holiday).
9. Upload result file → TC/Result Complete = Complete in app + Excel; upload reminders stop.
10. All above appear in Audit Log.

## 13. Milestones
1. Project skeleton + read 3 sheets + parsing (§6) + PMO review queue + unit tests §12.1–4.
2. Local DB + Audit Log + login (Outlook + OTP).
3. Screens: Dashboard, UAT Plan, Update dialog, Test Progress, Defects, Edit History, Settings + TH/EN.
4. Real-time 2-way sync + backup + Conflict Center.
5. Traffic Light + 4 email rules + BOT holidays + manual reminder button + email log.
6. Weekly Report (Claude API).
7. Full §12 tests + short install/user guide.

## 14. Ask user before doing
1. UI framework (offer 2–3 with pros/cons).
2. Stable Excel row ↔ app ID binding on row insert/delete (e.g. add ID column to Excel?).
3. Email language + template per rule (draft for approval).
4. OTP expiry, max attempts, re-login frequency.
5. BOT holiday source + behavior when fetch fails.
6. Retry policy for locked Excel.
7. Red→Amber: send email?
8. Monday holiday: Weekly Report next workday / skip / send anyway?
9. Confirm X = 5%.
10. Anything else unspecified found during work.
