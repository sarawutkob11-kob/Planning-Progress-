-- Part 5 v2 (approved 10-Nov-2026): user profile without OTP. Users get first / last name; "last login"
-- becomes "last used". One-time codes, lockout counters, sessions and SMTP settings are no longer used.

ALTER TABLE app_user ADD COLUMN first_name TEXT;
ALTER TABLE app_user ADD COLUMN last_name TEXT;
ALTER TABLE app_user RENAME COLUMN last_login_at TO last_used_at;

DROP TABLE IF EXISTS otp;
DROP TABLE IF EXISTS login_state;
DROP TABLE IF EXISTS session;

DELETE FROM settings WHERE key LIKE 'mail.smtp.%' OR key IN ('auth.bootstrapped', 'auth.last_account');
