-- Per-project result files folder (25-Nov-2026). NULL or '' = default <data_dir>/attachments of this PC.
-- Set in Settings by an Admin/PMO of the project (services/projects.py set_attachment_dir).
-- 001_app.sql and 002_project_delete.sql are left as they were.

ALTER TABLE project ADD COLUMN attachment_dir TEXT;
