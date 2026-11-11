-- Part 4: reminder recipients remembered per owner name (approved), new rows queued for Excel.

CREATE TABLE contact (
    name_key TEXT PRIMARY KEY,         -- owner name, case-folded, spaces collapsed
    name TEXT NOT NULL, email TEXT NOT NULL, updated_at TEXT NOT NULL
);
