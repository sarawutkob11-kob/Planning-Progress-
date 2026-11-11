# Progress Planning — Part 5/6 (v2): User profile (no OTP) + roles

## Rules
- Parts 1–4 done. FIRST read existing code, DB schema, UI, tests. Reuse. Mismatch → ask.
- Build ONLY this part. Not yet: AI report (Part 6).
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with design, schema changes, screen changes, open questions. Wait approval before code.
- Replaces previous Part 5 (OTP login). No OTP, no email verification.

## A. Identify user (no login)
1. First app start on a PC → form: First name, Last name, Email (required, validate format).
2. Save as local profile; app opens straight in next time (no prompt).
3. Top bar: show name + role; "Switch user" → same form.
4. Email = user key for roles, audit log, "own" items.
5. First profile ever → Admin (confirm dialog). Later new emails → default role: ask.

## B. Roles
| Role | Rights |
|---|---|
| Admin | All + user mgmt |
| PMO | Edit all, resolve conflicts + review queue, restore, send reminders, Settings (except user mgmt) |
| QA | Edit own tasks + own Test Progress rows, upload files, add defects; read others |
| User BU / PM | Same as PMO |
- Enforce in service layer (not UI only). UI hides/disables what role can't do.
- "Own" = match profile email with: Test Progress → Report By Email; UAT Plan → Responsible; Defect → Assigned To (name→email mapping for last two: ask).
- Denied action → message + audit log.

## C. User management (Settings, Admin only)
- List: first name, last name, email, role, active, last used.
- Add / edit role / deactivate (no delete).
- Optional import from Test Progress Report By Email + CC Email (Admin reviews first).

## D. Audit
- Editor in every audit/email log = "First Last <email>".
- Log: profile created, switch user, role change, user add/deactivate.

## Acceptance
1. First start → form; next start → no form.
2. Invalid email → cannot save.
3. First profile → Admin, once only.
4. QA cannot edit others' rows (service layer test).
5. User BU/PM = PMO rights.
6. Only Admin sees user management.
7. Edit History shows "First Last <email>".
8. Switch user changes editor in new log entries.
9. Unit tests for all.

## Ask before doing
1. Default role for new emails.
2. Name → email mapping for UAT Plan Responsible + Defect Assigned To.
3. Anything else unspecified.

## Known limit (tell user)
No verification → anyone can type any email and get that email's role. OK for Phase 1 (single user). Phase 2 multi-user → revisit.
