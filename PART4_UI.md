# Progress Planning — Part 4/6: Desktop UI

## Rules
- Parts 1–3 done. FIRST read existing code, DB schema, tests. UI calls existing functions (import, parsing, Traffic Light, emails, attachments, sync, conflicts, backup). No business logic in UI. Mismatch → ask.
- Build ONLY this part. Not yet: login/roles (Part 5), AI report (Part 6) → leave placeholders (e.g. "Generate Weekly Report" disabled).
- UI framework not chosen yet → propose 2–3 (pros/cons), ask. Chosen → use it.
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with screen list, navigation, components, open questions. Show 1 screen first (Dashboard) for approval, then rest.

## Global
- Language toggle TH/EN in top bar, all labels both languages, remember choice.
- Dates always `DD-Mon-YYYY`. Date input = date picker ONLY (no free typing).
- Controls ≥ 44px high, keyboard navigable, visible focus.
- Style: left nav navy `#14213D`; bg `#F4F6F9`; white cards, border `#E1E6EE`, radius 12px; primary `#1F4FBF`; Amber text `#7A4E00` bg `#FFF1D6`; Red text `#A3260F` bg `#FDE7E3`; Green/Complete `#1E6B3E` bg `#E3F2E8`; font IBM Plex Sans Thai (fallback Tahoma).
- Never color alone: color tags always carry text (Red/Amber/Overdue...).
- Every edit goes through existing functions → audit log + Excel write-back (Part 3).

## Shell
- Left nav: Dashboard, UAT Plan, Test Progress, Defect List, Edit History, Conflict Center (badge count), Settings.
- Top bar: project name, search, As-of date, sync status (synced / pending N / error), dry-run indicator, TH/EN.

## Screens
**1 Dashboard (PMO)**
- 4 KPI cards: TC executed (e.g. 91% · 847/926), open tasks, due today, open defects (overdue / no target).
- "Follow up today" list: tag (Overdue, No target, Today, Amber, Missing data) + item + detail + owner + "Send reminder" (Part 2 manual reminder).
- Incomplete apps: progress bars.
- Due within 2 weeks: list by date.
- Missing Report By Email warnings.
- "Generate Weekly Report" (disabled until Part 6).

**2 UAT Plan**
- Buttons: Import from Excel, + Add task.
- Filter chips: All, In progress, Not Start, No status, Due in 7 days, Rescheduled ≥2; owner filter.
- Table grouped A–E, collapsible (fully-complete groups collapsed by default). Cols: ID, Task, Owner, Status, Planned Finish (first), Current Finish (latest), reschedule count (≥2 red).
- Right detail panel on row click: Note, Status, Owner, Current Start/Finish, linked Test Progress (bar), reschedule timeline (date + change + reason), linked defects, "View full edit history" (→ Edit History filtered), "Update".

**3 Update dialog (tasks + Test Progress)**
- Status dropdown, Start/Finish date pickers, Remark.
- Finish changed → warning "Finish moved A → B (reschedule #N)" + REQUIRED reason + root cause (Dev not done, Waiting other system data, Defect found, Environment not ready, Waiting user check).
- Completed → REQUIRED Actual Finish.
- Actual Finish > Planned Finish → block save until Planned Finish rescheduled + reason.
- Note "All edits are logged". Cancel / Save.

**4 Test Progress**
- Summary: Executed/Total, Pass, Fail, Amber count.
- Filters: Needs follow-up, Incomplete data, Complete, All.
- Table: App, reporter, progress bar + Executed/Total, Pass/Fail (+ "Pass+Fail mismatch" tag), Planned Finish, Actual Finish (empty → red "Not filled"), Traffic Light (auto, read-only).
- Update results form: Executed, Pass, Fail, Actual Start/Finish, Report By Email, CC Email (validate format).
- Upload test-result file (Part 2 store) → TC/Result Complete auto Complete.

**5 Defect List**
- Toggle Board / Table; + Add Defect.
- Board: Open → Assigned → Fixed → Retest → Close/Rejected, counts per column; move card = status change (audit logged).
- Card: #no · app, Severity tag (1–2 red), detail, owner, days open, Overdue / No target tag.
- Add/edit form: REQUIRED individual owner + Fix target date.

**6 Edit History**
- Filters: search, page, editor, date range, type (Reschedule, Status change, Data edit, Import).
- Table: date, editor, type, item + page, field, old (strikethrough) → new, reason, "Restore" (creates new entry).
- Export to Excel.

**7 Conflict Center**
- Tab Conflicts: item, field, Excel value, app value, edit time each side, buttons Use Excel / Use app / Enter new + reason.
- Tab PMO review queue (Part 1): raw value, sheet, item, suggestion field → PMO enters/chooses value.
- Show mockup before building.

**8 Settings**
- Excel file path, Fail X%, email time (08:30), report time (Mon 09:00), dry-run ON/OFF (confirm dialog before turning OFF), BOT holiday status + last fetch, backup folder + restore, email log viewer, email templates, Claude API key field (placeholder until Part 6), users/roles (placeholder until Part 5).

## Acceptance
1. All 8 screens reachable from nav; TH/EN switches every label.
2. Date fields reject typed text; only picker.
3. Reschedule without reason → cannot save.
4. Completed without Actual Finish → cannot save.
5. Actual Finish > Planned Finish → blocked until reschedule + reason.
6. Edit in UI → Excel updates (Part 3) + Edit History row appears.
7. Edit in Excel → UI refreshes without restart.
8. Conflict resolved in UI → both sides updated + logged.
9. Upload file → TC/Result Complete shows Complete.
10. "Send reminder" respects dry-run.
11. Dashboard numbers match DB (test with real file: 847/926 executed on 6-Nov-2026 data).

## Ask before doing
1. UI framework (if not chosen).
2. Dashboard mockup approval before other screens.
3. Conflict Center mockup.
4. Anything else unspecified.
