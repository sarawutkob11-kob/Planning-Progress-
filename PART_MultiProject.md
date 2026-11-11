# Planning Progress — Add-on: Multi-project

## Rules
- Parts 1–4 done. FIRST read existing code, DB schema, tests, settings. Reuse. Mismatch with this prompt → ask, don't rewrite.
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with design, schema/migration plan, affected files, open questions. Wait approval before code.
- Backup DB + every Excel file before migration/rename. Never delete data. Test on copies first.

## A. App name
- Use "Planning Progress" everywhere: window title, menu, Weekly Report email subject/header, logs, docs.
- Replace all "Progress Planning" / "UAT Tracker" in code + UI text.

## B. Project selector
- Dropdown at top of left menu, under app name.
- Selected project filters ALL pages: Dashboard, UAT Plan, Test Progress, Defect, History (+ Conflict Center, Audit Log).
- No cross-project overview page.
- Remember last selected project on restart.

## C. 1 project = 1 Excel file
- File name: `Planning Progress <Project Name>.xlsx` (prefix always "Planning Progress ").
- Add task / Application / Defect via forms → write to that project's own file only.
- Sync, conflict, history, attachments, Traffic Light, alerts: all scoped per project.

## D. Create project (form)
- Field: Project Name only.
- Ask folder every time (folder picker).
- Invalid filename chars `/ \ : * ? " < > |` → auto-replace with `-` in file name (e.g. `A/B` → `Planning Progress A-B.xlsx`). Project display name keeps original text. Show final file name before create.
- Duplicate check uses final file name (after replace): exists in chosen folder, or project already in app → block, show message.
- New file = 3 sheets only: UAT Plan, Test Progress, Defect List. Headers + formatting (col widths, styles, validation, any ID column added in Part 1) copied from `UAT Planning.xlsx`. No data rows.
- Store template definition in code/resource, not by reading the live source file each time.

## E. Existing project
- Current `UAT Planning.xlsx` → project "RAMC UAT", file renamed to `Planning Progress RAMC UAT.xlsx`.
- File is on OneDrive, others may use it → before rename: backup, show user old/new path, wait confirm.
- Migrate existing DB data → project RAMC UAT (add project_id etc.). Verify counts before/after.

## F. Import existing Excel as project
- Pick .xlsx file → validate 3 main sheets + required headers. Missing → show what's missing, don't import.
- Never rename imported file. Keep file name as-is.
- Project name = file name minus `.xlsx` (and minus "Planning Progress " prefix if present).
- Other sheets in file: ignore, never modify.
- Same duplicate rules as D.

## G. Per project (separate) vs shared
Per project:
- Users + Roles (one user can be in some projects only; role may differ per project)
- Fail X% threshold (default 5%)
- Weekly Report (one report per project, own recipients)
- Email send time + dry-run switch
Shared (confirmed by user): Claude API key, bank holiday cache, Outlook/PMO account. Other app-level settings not listed → ask.

## H. Tests
- Create project → file name, 3 sheets, headers/format match template, empty.
- Invalid char / duplicate name → blocked.
- Switch project → every page shows only that project's data.
- Form add → writes only to selected project's file.
- Import valid / invalid file.
- RAMC UAT migration: row counts + Traffic Light results same as before.
- Per-project X%, dry-run, send time, roles isolated (change in A doesn't affect B).

## I. Temporary "Clear project data" button
- Temporary, for testing. Put behind config flag `ENABLE_CLEAR_PROJECT` (default on now), easy to remove later. Keep code isolated (one module + one button).
- Location: Settings page, separate "Danger zone" section, red button. PMO role only.
- Clears app-side data of SELECTED project only: tasks, apps, defects, Traffic Light history, sync/conflict, audit log, alert/email log, cached attachments + history on PMO PC.
- Other projects: untouched.
- Keeps: project itself (name, file link, per-project settings: users/roles, X%, Weekly Report, send time, dry-run), Excel file (never touched), shared settings.
- Before clear: auto backup that project's DB data + attachment folder to timestamped zip, show path.
- Confirm: dialog shows project name + what will be cleared/kept, user must type the project name to enable button.
- After clear: project stays selected, pages show empty state. Log action (user, time, project) in a separate file, since project audit log is cleared.
- Test: clear project A → A empty, project B data unchanged, A's Excel file unchanged (hash before/after), A's settings kept, backup zip restores correctly.

## Open questions (ask in Step 1)
1. User with no role in selected project → hide project from dropdown, or show read-only?
2. Project delete/archive needed?
3. Excel file moved/renamed outside app → how to relink?
4. Restore from Clear-project backup: need a restore button, or manual only?
5. Anything else unspecified found while reading code.
