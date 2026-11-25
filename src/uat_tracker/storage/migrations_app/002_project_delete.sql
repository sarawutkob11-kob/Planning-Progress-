-- Delete project (approved 25-Nov-2026): policy change from "archived, never deleted".
-- A project can now be deleted for good, but only after it has been archived.
-- The service (services/delete_project.py) writes a backup zip first; this trigger is the
-- second guard so a plain DELETE on an active project (archived = 0) is still refused.
-- 001_app.sql is left as it was; this file replaces its trigger.

DROP TRIGGER project_no_delete;
CREATE TRIGGER project_no_delete BEFORE DELETE ON project
WHEN OLD.archived = 0
BEGIN SELECT RAISE(ABORT, 'archive the project before deleting it'); END;
