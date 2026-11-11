# Progress Planning — Part 2/6: Traffic Light + email alerts

## Rules
- Part 1 (core data layer) done. FIRST read existing code, DB schema, tests. Reuse them. Mismatch with this prompt → ask, don't rewrite Part 1.
- Build ONLY this part. Not yet: 2-way sync (Part 3), UI (Part 4), login (Part 5), AI report (Part 6).
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with design, libs, schema changes, open questions. Wait approval before code.
- Excel still read-only this part. Before each scheduled run, re-import Excel (Part 1 import) to get latest data.

## Constraints
- No Microsoft Graph / MSAL. Email via running Outlook desktop, PMO account (e.g. pywin32 COM).
- Runs on PMO PC. Job fires only if PC + Outlook on. PC off at run time → skip, no backfill, log skip.
- **Dry-run mode REQUIRED and ON by default**: emails go to PMO only (or preview file), never real recipients. Switch to live only when user says so.

## A. Traffic Light (Test Progress, auto, no manual)
`%Exec = Executed/Total TC`, `%Time = (today − Planned Start)/(Planned Finish − Planned Start)`, `%Fail = #Fail/Executed`.
First match wins:
1. Status Complete → Complete
2. Past Planned Finish and not Complete, OR %Fail > X → Red
3. %Exec < %Time, OR #Fail ≥ 1 and %Fail ≤ X → Amber
4. Else → Green
- X default 5%, configurable.
- Store color history per app (for change detection + audit log).
- Missing dates/TC → no color, flag "missing data".

## B. Working days
- Workday = Mon–Fri minus bank holidays.
- Bank holidays: Bank of Thailand (BOT) official financial-institution holidays. Find source, tell user source + fetch frequency before building. Cache locally. Fetch fail → warn (fallback: ask).
- Rule trigger date on non-workday → previous workday.

## C. Email common
- Only Test Progress drives auto emails.
- Run 08:30 workdays.
- 1 email per item per rule. No digest.
- To = Report By Email (one). CC = CC Email (`;` list, empty → none). Report By Email empty → don't send, flag "missing email".
- Email log: time, rule, app, to, cc, result (sent/skipped/failed/dry-run), error.
- Each send → audit log entry.
- Dedupe: same rule + app + day → max 1 send.

## D. Rules
**R1 Planned Start**
- On Planned Start date: send once "task starts today".
- From Planned Start on, Actual Start empty → daily until filled.

**R2 Planned/Actual Finish**
- On Planned Finish date: remind to fill Actual Finish.
- From Planned Finish on, Actual Finish empty → daily until filled.
- Actual Finish ≤ Planned Finish = valid. Actual Finish > Planned Finish = flag for PMO (enforcement in UI = Part 4, conflict = Part 3).

**R3 Traffic Light change**
- Email once on change INTO Amber/Red: Green→Amber, Green→Red, Amber→Red, re-entry after Green.
- Red→Amber: ask user.

**R4 Close-out**
- Condition: Executed = Total TC AND Actual Finish filled (≤ Planned Finish).
- Status not Complete → daily email to change Status. Never auto-change Status.
- TC/Result Complete not Complete → daily email to upload test-result docs.
- Build attachment store now (save file on PMO PC, link to app). On upload → TC/Result Complete = Complete in DB + audit log. Excel write-back in Part 3. Upload button in Part 4 (CLI/function OK now).

## E. Manual reminder (UAT Plan + Defect List)
- No auto email for these sheets.
- Build function to send reminder for 1 item via Outlook (respects dry-run). Button wired in Part 4.

## F. Scheduler
- Daily 08:30 workdays job: re-import → compute colors → evaluate R1–R4 → send.
- Must survive app restart (state in DB).
- Propose: in-app scheduler vs Windows Task Scheduler. Ask.

## Acceptance
1. DMS-RAMC (126/186, Fail 2, plan 22-Oct → 13-Nov-2026), today 6-Nov-2026 → Amber; first run sends 1 R3 email, second run same color → none.
2. Fail 10/100, X 5% → Red.
3. Planned Finish Saturday → R2 reminder Friday; Friday bank holiday → Thursday.
4. Actual Start empty 3 workdays after Planned Start → 3 emails, none on weekend.
5. Executed = Total, Actual Finish filled, Status not Complete → daily R4 email; Status set Complete → stops.
6. Upload file → TC/Result Complete = Complete, upload reminders stop.
7. Report By Email empty → no send, flagged.
8. Dry-run on → nothing reaches real recipients.
9. Every send/skip in email log + audit log. Unit tests for all.

## Ask before doing
1. Email language + template per rule (draft for approval).
2. Red→Amber email?
3. Scheduler approach.
4. BOT source + fetch fail behavior.
5. Attachment storage folder + allowed file types/size.
6. Anything else unspecified.
