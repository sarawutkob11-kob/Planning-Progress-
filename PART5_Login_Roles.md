# Progress Planning — Part 5/6: Login (Outlook + OTP) + roles

## Rules
- Parts 1–4 done. FIRST read existing code, DB schema, UI, tests. Reuse. Mismatch → ask.
- Build ONLY this part. Not yet: AI report (Part 6).
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with design, schema changes, screen changes, open questions. Wait approval before code.

## Constraints
- No Microsoft Graph / MSAL / SSO. Identity = Outlook desktop account + OTP email.
- Phase 1: single user (PMO) on own PC. Build roles fully for Phase 2 (multi-user via OneDrive).

## A. Login flow
1. App start → read email of Outlook account(s) open on PC. Multiple accounts → user picks one.
2. Email not in user table → access denied (except first-run bootstrap).
3. Send OTP to that email via Outlook (not subject to dry-run).
4. User enters OTP → session starts.
- OTP: random, single-use, stored hashed, expiry + max attempts + lockout (values: ask).
- Session timeout / re-login frequency: ask.
- Outlook closed → clear message, no login.
- Logout button in top bar; show current user + role.
- First run: no users yet → current Outlook email becomes Admin (confirm dialog).

## B. Roles
| Role | Rights |
|---|---|
| Admin | All + user mgmt |
| PMO | Edit all, resolve conflicts + review queue, restore, send reminders, Settings (except user mgmt) |
| QA | Edit own tasks + own Test Progress rows, upload files, add defects; read others |
| User BU / PM | Same as PMO |
- Enforce in service layer (not UI only). UI hides/disables what role can't do.
- "Own" = match logged-in email with: Test Progress → Report By Email; UAT Plan → Responsible (name only today → name→email mapping: ask); Defect → Assigned To (ask same).
- Denied action → message + audit log entry.

## C. User management (Settings, Admin only)
- List users: email, display name, role, active, last login.
- Add / edit role / deactivate (no delete; audit kept).
- Import users from Test Progress Report By Email + CC Email list (optional button, Admin reviews before saving).

## D. Audit
- Editor in every audit/email log entry = logged-in email (replace placeholder from earlier parts).
- Log: login success/fail, lockout, logout, role change, user add/deactivate.

## Acceptance
1. Known email + correct OTP → in; wrong OTP → error; over max attempts → locked.
2. Expired OTP rejected; used OTP rejected.
3. Unknown email → denied.
4. First run → Admin bootstrap works once only.
5. QA cannot edit others' rows (service layer test, not just UI).
6. User BU/PM can do everything PMO can.
7. Only Admin sees user management.
8. All edits show real user email in Edit History.
9. Outlook closed → clear error.
10. Unit tests for all.

## Ask before doing
1. OTP length, expiry, max attempts, lockout time.
2. Session timeout / re-login frequency.
3. Name → email mapping for UAT Plan Responsible + Defect Assigned To.
4. Multiple Outlook accounts handling.
5. Anything else unspecified.

## Known limit (tell user, don't solve alone)
OTP goes to the same Outlook open on the PC → proves mailbox access, not a separate factor. Anyone using the unlocked PC can log in. Acceptable for Phase 1? Ask.
