-- SharePoint link of the per-project result files folder (25-Nov-2026). NULL = folder chosen on this PC (or
-- default). When set, attachment_dir holds the local folder OneDrive syncs for that link
-- (services/sharepoint_link.py); uploads still use attachment_dir. 001-003 are left as they were.

ALTER TABLE project ADD COLUMN attachment_url TEXT;
