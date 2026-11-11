# Progress Planning — Part 1/6: Core data layer

## Rules
- Role: senior Python dev, Windows desktop app.
- This is part 1 of 6. Build ONLY this part. Design data model so later parts plug in: 2-way Excel sync, email alerts, desktop UI, login/roles, AI weekly report.
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with architecture, libs, folder layout, DB schema, open questions. Wait approval before code.
- Never touch source Excel. Work on a copy. Part 1 = read only (no write-back yet).

## Context
- User = PMO, project RAMC UAT. Source `UAT Planning.xlsx`, OneDrive (company), synced locally.
- Phase 1: PMO only, own PC; app data on PMO PC. Phase 2 (later): multi-user via OneDrive → storage location must be config-switchable.
- No Azure App Registration → no Microsoft Graph / MSAL ever.

## Scope
Read 3 sheets → normalize → store in local DB + immutable audit log + PMO review queue. CLI or minimal script to run import + print summary is enough. No UI yet.

## Source sheets
Map columns by header name, never position. Never read/touch other sheets.

**UAT Plan**: Task no, Sub task, Group, Task No., Task name, Note, File Path, Responsible (Mandatory), Status, Start date, Start time, Finish date, Finish time, Remark.
- Groups A–E; group header = row with long text in Task no (e.g. `A: SETUP for UAT Testing`).
- Task numbers duplicate → unique internal ID (e.g. `E-23`) bound stably to Excel row (binding method: propose, ask).

**Test Progress**: No., Application, Report By, Report By Email (new, ONE address), CC Email (new, `;`-separated), Total TC, No. of TC Execution, %Progress, #Pass, #Fail, Planned Start, Planned Finish, Actual Start, Actual Finish, Traffic Light, TC/Result Complete, Status, Remark.
- Last row Grand Total = sum, skip. New email columns may not exist yet → handle missing.

**Defect List**: No., Date, System/App., Issues/detail, Status, Severity, Assigned To, Fix target date, Data test, Raised by, Remark.

## Parsing rules
1. Dates CE. Year > 2500 → −543 (`9/14/2569` → 14-Sep-2026).
2. Multi-date cell → LAST TYPED value = current, regardless earlier/later. Prior values = reschedule history (ordered). Handles `=>`, `->`, space-separated w/o arrow.
   - `10/09/2026 => 19/09/2026 => 15/09/2026` → 15-Sep-2026
   - `20/10/2026 25/9/2026 => 17/09/2026 14/9/2026` → 14-Sep-2026
3. Non-date text in date cell (`(BD 30/09/26)`, `(actual start 25/09/26)`) → note field.
4. Ambiguous (`10/8/2026`) or typo (`24/09/206`, `29/09/2027`, `1/4/1900`, `06-Now-2026`) → NO guessing → PMO review queue with raw value.
5. Display/store format DD-Mon-YYYY.
6. Status normalize: `Not Start`, `In progress`, `Blocked`, `Complete`/`Completed` (complete/comlete/Inprogress → standard). Unknown → review queue.
7. Flag data issues: Pass+Fail ≠ Executed, missing owner, missing Report By Email, Complete without Actual Finish.

## Audit log
- Every change (import included): time, editor, channel, sheet, item ID, field, old, new, reason.
- Immutable. Restore later = new entry.

## Acceptance
1. CCMS row example (rule 2 second case) → Start 14-Sep-2026, history 3 prior values.
2. `9/14/2569` → 14-Sep-2026.
3. Typo values (rule 4) → review queue, not converted.
4. Test Progress ATM (Executed 4, Pass 2, Fail 0) → flagged mismatch.
5. Re-import same file → no duplicate rows, no new audit entries.
6. Unit tests for all above.

## Ask before doing
1. Row ↔ ID binding on row insert/delete (add ID column to Excel?).
2. DB choice (propose).
3. Anything else unspecified.
