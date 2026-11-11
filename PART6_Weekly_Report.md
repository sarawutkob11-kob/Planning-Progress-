# Progress Planning — Part 6/6: AI Weekly Report

## Rules
- Parts 1–5 done. FIRST read existing code, DB schema, UI, tests. Reuse (data, Traffic Light, workdays, Outlook sender, dry-run, scheduler, audit/email log, roles). Mismatch → ask.
- Build ONLY this part.
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with design, prompt draft for the AI, sample report, open questions. Wait approval before code.

## A. Schedule
- Auto every Monday 09:00 (Monday not workday: ask).
- Same scheduler as Part 2. PC/Outlook off → skip, log, no backfill.
- Before generating: re-import Excel (latest data).
- Manual: "Generate Weekly Report" on Dashboard (enable placeholder from Part 4) → preview → Send / Cancel. Roles: PMO, User BU/PM, Admin.

## B. Recipients
- All Report By Email + CC Email in Test Progress, deduped, invalid skipped + flagged.
- Send from Outlook, PMO account. Respect dry-run (→ PMO only).
- Log in email log + audit log.

## C. Content (data computed by code, AI writes summary only)
Code computes numbers; AI must not invent or recalc numbers.
1. Overall: Executed/Total %, Pass, Fail.
2. Amber/Red apps + reason (time / fail rate / overdue).
3. Overdue tasks (UAT Plan + Test Progress).
4. Open defects; overdue fix target; no fix target.
5. Due in next 2 weeks.
6. Reschedules last 7 days + reasons (from history).
7. Open conflicts + review-queue count.
- AI part: short executive summary + top risks + suggested follow-ups, based only on provided data.
- Numbers/tables in email rendered by code, not AI.

## D. Claude API
- Key entered in Settings (field from Part 4), stored encrypted (propose method, e.g. Windows Credential Manager/DPAPI). Never in logs, Excel, DB plain text.
- Model name configurable in Settings (propose default).
- Send only needed fields (aggregates + item names/dates/reasons). No attachments, no file paths.
- Timeout + retry; API fail → behavior: ask (e.g. send report without AI summary + flag, or skip + alert PMO).
- Log tokens/cost per run (if available).

## E. Storage
- Save every report (date, recipients, content, AI on/off, status) on PMO PC; view past reports in app.

## Acceptance
1. Monday 09:00 workday → report generated + sent (dry-run → PMO only).
2. Monday holiday → follows rule user chose.
3. Recipients = deduped Report By Email + CC Email.
4. Numbers in email match DB exactly (test with real file).
5. AI text never contains numbers not in input (test check).
6. API key not found in any log/DB/Excel plain text.
7. API down → chosen fallback works.
8. Manual generate → preview → send works; QA cannot send.
9. Past reports viewable.
10. Unit tests for all.

## Ask before doing
1. Monday holiday rule.
2. Report language (TH / EN / both) + email format (HTML body / attachment).
3. API fail behavior.
4. Default model + key storage method.
5. Anything else unspecified.

## Tell user
Report sends project data to an external AI service (Anthropic). Confirm company policy allows it before going live.
