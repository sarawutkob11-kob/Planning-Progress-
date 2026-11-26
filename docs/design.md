# Design — uat_tracker

> รอบถัดไปให้เพิ่ม section ใหม่ต่อท้าย ห้ามเขียนทับ section เดิม

---

# 2026-11-25 — หน้าตาใหม่ของเมล Weekly Report และชื่อสถานะ "มีความเสี่ยงล่าช้า"

อ้างอิง: `docs/spec.md` section 2026-11-25 (R1-R16, AC-1 ถึง AC-23)

## ภาพรวมแนวทาง
1. **ชื่อสถานะ** แก้ที่ตาราง label จุดเดียว (`LABEL_TH` / `LABEL_EN` ใน `traffic_light.py`) ทุกที่ที่อ่านตารางนี้ (i18n, alerts, weekly, `daily_job` ที่เขียน Excel) เปลี่ยนตามเอง ส่วนที่เป็นข้อความตายตัว (template JSON, README, คอมเมนต์) แก้ตรงตัว
2. **ข้อมูลรายงาน** `weekly_data.build_data` ยังเป็นที่เดียวที่คำนวณตัวเลข เพิ่มเฉพาะข้อมูลที่หน้าใหม่ต้องใช้ (จำนวนวันที่เลยกำหนดของแอปล่าช้า, รายการแอปข้อมูลไม่ครบ, ช่วง 7 วัน)
3. **หน้าเมล** เขียน `weekly_report.render` ใหม่เป็น HTML แบบ table + inline style ตามแบบที่อนุมัติ แยกเป็นฟังก์ชันย่อยทีละส่วน เหตุผลภาษาไทยสร้างใน `weekly_report` จากตัวเลขใน data (ไม่แตะ `Light.reason`)
4. **เรื่องภายใน** `render` ไม่รับข้อความ AI error / notes อีกต่อไป เรื่องผู้รับไปแสดงในหน้า preview และแจ้ง PMO ผ่าน `alert_pmo`

เหตุผล: ใช้กลไกเดิมให้มากที่สุด (label เดียว, `alert_pmo`, `_queue_traffic_lights` ที่เขียนทับคำเก่าอยู่แล้ว) ไฟล์ที่แก้มีน้อย และ logic ทุกส่วนทดสอบด้วย pytest ได้โดยไม่ต้องใช้ Qt

## ไฟล์ที่แก้

| ไฟล์ | สิ่งที่เปลี่ยน |
|---|---|
| `services/traffic_light.py` | `LABEL_EN[AMBER] = "At risk of delay"`, `LABEL_TH[AMBER] = "มีความเสี่ยงล่าช้า"`; เพิ่ม `missing_fields(rec)` และให้ `compute` ใช้ (ข้อความ `reason` เหมือนเดิมทุกตัวอักษร) |
| `ui/i18n.py` | แก้คอมเมนต์; เพิ่ม key `wk.invalid` |
| `alerts/templates.py` | ไม่แก้โค้ด (`COLOR_TH` อ่านจาก label) |
| `services/daily_job.py` | แก้ docstring ของ `_queue_traffic_lights` เท่านั้น |
| `resources/project_template.json` | dropdown, conditional format และค่าตัวอย่าง "At risk" -> "At risk of delay" |
| `README.md` | ส่วนสถานะ Traffic Light, R3, Dashboard และ Weekly report ให้ตรงคำใหม่และช่วง 7 วัน |
| `services/weekly_data.py` | `days_ahead` default 7; key `due_next_2_weeks` -> `due_next_7_days`; ใน `delayed_or_at_risk` เพิ่ม `days_late` (กรณี overdue); เพิ่ม key `missing_data: list[{"id","app","fields"}]`; คง `review_queue` ไว้ใน data (ไม่แสดงในเมล) |
| `services/weekly_report.py` | `render(data, ai)` เขียนใหม่; `_reason_th`; ลบ `_KIND`, `_h`, `_table`, `_ai_html` เดิม; `generate` ไม่ส่ง notes เข้า render; `run_weekly` เพิ่ม `alert_pmo` เรื่องอีเมลไม่ถูกต้อง; เพิ่ม `preview_notes(r)` |
| `ui/weekly.py` | `WeeklyPreviewDialog` แสดงบรรทัดจาก `preview_notes(r)` ใต้บรรทัด AI |
| `tests/test_status_labels.py`, `tests/test_weekly.py` | แก้ assertion ที่อ้างคำ/หน้าเดิม และเพิ่ม test ตาม AC |

## Interface สำคัญ

```python
# services/traffic_light.py
def missing_fields(rec: dict) -> list[str]:
    """Blank required fields: 'Total TC', 'Planned Start', 'Planned Finish' (this order)."""

# services/weekly_report.py
def render(data: dict, ai: dict | None) -> tuple[str, str]:
    """(subject, html). Lays out numbers from build_data only; no calculation here."""

def _reason_th(a: dict, threshold: str) -> str:
    """Thai one-line reason for one row of data['delayed_or_at_risk'] (R8)."""

def preview_notes(r: dict) -> list[str]:
    """Lines for the preview dialog: invalid recipients (from r['invalid_recipients'])."""
```

### ส่วนของ `render` (ลำดับตามแบบ)
`_header`, `_summary_line`, `_ai_block` (เมื่อ `ai` ไม่ใช่ None), `_kpis`, `_status_bar`, `_watch_list`, `_overdue_plan`, `_defects`, `_due_week`, `_other`, `_footer` แต่ละตัวคืน `<tr>...</tr>` ของตารางหลักกว้าง 640px หัวข้อใช้ `_heading(th, en)` ตัวเดียว (บรรทัดไทยตัวหนา + บรรทัดอังกฤษเล็กสีเทา)

### `_reason_th` แยกกรณีจากข้อมูลที่มีอยู่
| เงื่อนไข (`reason_kind`, สถานะ, `fail`) | ข้อความ |
|---|---|
| `overdue` | เลยวันจบตามแผน ({planned_finish}) มา {days_late} วัน และยังไม่เสร็จ |
| `fail_rate` + ล่าช้า | Fail {pct_fail} เกินเกณฑ์ {threshold} |
| `fail_rate` + มีความเสี่ยงล่าช้า | มี Fail {fail} เคส ({pct_fail}) แต่ยังไม่เกินเกณฑ์ {threshold} |
| `time` + `fail == 0` | ทดสอบได้ {pct_exec} แต่เวลาผ่านไปแล้ว {pct_time} |
| `time` + `fail >= 1` | ทดสอบได้ {pct_exec} แต่เวลาผ่านไปแล้ว {pct_time} · มี Fail {fail} เคส ({pct_fail}) แต่ยังไม่เกินเกณฑ์ {threshold} |

แอปข้อมูลไม่ครบมาจาก `data["missing_data"]`: "ยังไม่ได้กรอก " + ", ".join(fields)

## Data model / schema
ไม่เปลี่ยน schema ของ DB ค่า internal (Red/Amber/Green) คงเดิม `weekly_report.data_json` ของรายงานใหม่มี key ใหม่ตามตาราง รายงานเก่าไม่แปลง

## ความเสี่ยงและผลกระทบ
- **ข้อมูลที่ส่งให้ AI เปลี่ยนเล็กน้อย** `weekly_ai` ส่ง `data` ทั้งก้อน key ที่เปลี่ยนชื่อหรือเพิ่ม และคำ "At risk of delay" จะไปถึง AI ด้วย prompt ไม่ได้อ้างชื่อ key (ตรวจแล้ว) จึงไม่แก้ prompt
- **ไฟล์ Excel เดิม** ได้ "At risk of delay" แต่ dropdown และ conditional format เดิมยังเป็น "At risk" (ผู้ใช้รับทราบแล้วตาม spec)
- **คอลัมน์ Traffic Light ใน Excel** คำยาวขึ้น อาจตัดคำในไฟล์เดิม ไม่แก้ความกว้างคอลัมน์ (นอก spec)
- **Test ฝั่ง Qt** การแก้ `ui/weekly.py` เป็นแค่เพิ่ม label จาก `preview_notes` logic ทดสอบที่ `preview_notes` ส่วน `tests/test_ui.py` ต้องรันบน Windows ที่มี PySide6
- **Outlook** HTML ใช้ table + inline style ควรส่ง dry-run จริงหนึ่งครั้งหลัง test ผ่าน เพื่อดูหน้าตาใน Outlook

---

# 2026-11-25 — ลบโปรเจกต์ (Delete project)

อ้างอิง: `docs/spec.md` section "2026-11-25 — ลบโปรเจกต์ (Delete project)" (D1-D12, AC-1 ถึง AC-33) ผู้ใช้อนุมัติ A1-A8 ตามที่เขียน

## ภาพรวมแนวทาง
1. **service ใหม่ไฟล์เดียว** `services/delete_project.py` (แยกจาก `projects.py` เพราะมีงานไฟล์/zip/rollback แบบเดียวกับ `clear_project.py`) ใช้ `ProjectError` เดิมเป็น error ทุกกรณี ทำให้ CLI เดิม (`except ProjectError -> "NOT DONE: ..." exit 2`) และ UI (`error(self, friendly(exc))`) รองรับโดยไม่ต้องเพิ่มกลไก
2. **ลำดับ fail safely** ตาม A4: ตรวจ -> zip (เขียนเป็น `.part` แล้ว `os.replace` เมื่อเสร็จ) -> rename โฟลเดอร์เป็น `projects/.deleting-<id>-<stamp>` -> transaction ของ `app.db` (ลบแถว + event "deleted") -> `rmtree` โฟลเดอร์ชั่วคราว จุดที่แตะไฟล์ระบบผ่าน helper ระดับ module `_rename` / `_rmtree` เพื่อให้ test monkeypatch ได้ตรงจุด
3. **กันชั้นที่สองใน DB** migration `002_project_delete.sql` แทน trigger เดิมด้วย trigger ที่ abort เฉพาะแถว `archived = 0`
4. **id ไม่ใช้ซ้ำ** `AppDb.next_project_id()` นับ max จาก `project.id` รวมกับ `project_event.project_id`
5. **UI** ปุ่มแดงใต้ตาราง Archived ใน `_project_card` + `DeleteProjectDialog` (โครงเดียวกับ `ClearProjectDialog`) รันใน worker ที่เปิด `AppDb` ของตัวเอง (แบบ `relink_flow` เพราะ sqlite connection ข้าม thread ไม่ได้)
6. **CLI** `project delete <project> [--confirm-name NAME]` ใช้ service เดียวกัน

เหตุผล: ใช้ `require_manager`, `ProjectError`, `DANGER_QSS`, `run_in_worker`, รูปแบบ backup ของ `clear_project._backup` ที่มีอยู่แล้ว ไม่แตะ `clear_project.py` (feature ชั่วคราวที่จะถูกถอดออก จึงไม่ใช้ i18n key `danger.*` ร่วมกัน) ไม่เปลี่ยน schema ตาราง

## ไฟล์ที่สร้าง / แก้

| ไฟล์ | สิ่งที่เปลี่ยน |
|---|---|
| `storage/migrations_app/002_project_delete.sql` (ใหม่) | `DROP TRIGGER project_no_delete;` แล้วสร้างใหม่ชื่อเดิม `BEFORE DELETE ON project WHEN OLD.archived = 0 BEGIN SELECT RAISE(ABORT, 'archive the project before deleting it'); END;` + คอมเมนต์อธิบายว่าเปลี่ยนนโยบาย "never deleted" (ห้ามแก้ `001_app.sql`) |
| `storage/app_db.py` | เพิ่ม `delete_project(project_id)`; แก้ `next_project_id()` ให้รวม id จาก `project_event` |
| `services/delete_project.py` (ใหม่) | `delete_project(...)`, `DeleteResult`, `DeleteLocked`, `what_is_deleted(...)`, helper `_rename`, `_rmtree`, `_write_backup` |
| `services/projects.py` | แก้ docstring บรรทัด "Archive only ... No delete." เท่านั้น (D12) |
| `ui/projects.py` | เพิ่ม `DeleteProjectDialog`; แก้ docstring ของ module |
| `ui/screens/settings.py` | `_project_card`: ปุ่ม `self.del_btn` ข้างปุ่ม Restore selected; เพิ่ม `_delete_archived()` |
| `ui/i18n.py` | key ใหม่ `proj.delete*`, `work.deleting` (EN + TH) |
| `cli_project.py` | subcommand `delete`, docstring / help ของ `project` |
| `README.md` | ส่วน Settings (บรรทัด ~291), คำสั่ง `project` (~332-333) และย่อหน้าสั้นเรื่องลบโปรเจกต์ + `delete-backups` |
| `tests/test_delete_project.py` (ใหม่, QA) | test ของ service / migration / CLI / UI ตาม AC |

## Interface สำคัญ

```python
# storage/app_db.py
class AppDb:
    def delete_project(self, project_id: str) -> None:
        """DELETE FROM project WHERE id=?. Caller wraps in transaction(); the trigger aborts unarchived rows."""

    def next_project_id(self) -> str:
        """p001, p002 ... one more than the highest id ever used (project table + project_event)."""


# services/delete_project.py
BACKUP_DIR_NAME = "delete-backups"          # <root>/delete-backups/<id>_<yyyymmdd-HHMMSS>.zip
TEMP_PREFIX = ".deleting-"                   # <root>/projects/.deleting-<id>-<stamp>/
DB_FILES = ("tracker.db", "tracker.db-wal", "tracker.db-shm")

class DeleteLocked(ProjectError):
    """Folder could not be renamed (file in use). Nothing deleted; backup_zip is kept."""
    backup_zip: Path

@dataclass(frozen=True)
class DeleteResult:
    project_id: str
    name: str
    backup_zip: Path
    leftover: Path | None                    # temp folder that could not be fully removed (D5.3)

def project_dir(base: Config, project_id: str) -> Path:
    """<root>/projects/<id>"""

def what_is_deleted(base: Config, project: dict) -> dict[str, str]:
    """For CLI / dialog text: {"folder": ..., "excel": ..., "backup_dir": ...}. Reads nothing."""

def delete_project(base: Config, appdb: AppDb, person: Person, project_id: str, *,
                   typed_name: str) -> DeleteResult:
    """Delete one ARCHIVED project for good (D1-D9). Raises ProjectError (DeleteLocked when the folder is in use)."""

def _write_backup(base: Config, project: dict, folder: Path | None, person: Person, stamp: str) -> Path:
    """Zip written as <name>.zip.part then os.replace -> <name>.zip. Any error: .part and temp db removed, re-raised."""

def _rename(src: Path, dst: Path) -> None:    # os.rename; monkeypatch point for AC-12
def _rmtree(path: Path) -> None:              # shutil.rmtree; monkeypatch point for AC-14


# ui/projects.py
class DeleteProjectDialog(QDialog):
    def __init__(self, base: Config, person: Person, project: dict, parent=None) -> None: ...
    # attrs for tests: self.typed (QLineEdit), self.go (QPushButton), self.result: DeleteResult | None


# ui/screens/settings.py
class SettingsScreen:
    def _delete_archived(self) -> None: ...   # self.del_btn.clicked
```

### ลำดับใน `delete_project` (ทุกข้อก่อนข้อ 6 ไม่เปลี่ยนอะไรเลย)
1. `require_manager(base, appdb, person)` -> ProjectError "Not allowed..." (AC-2)
2. `project = appdb.get_project(project_id)`; `None` -> `ProjectError(f"Unknown project {project_id}")` (AC-5)
3. `project["archived"] != 1` -> `ProjectError("Archive the project before deleting it")` (AC-3)
4. `typed_name.strip() != project["name"].strip()` -> `ProjectError("The typed name does not match the project name")` (AC-4, ตรงตัว ไม่ casefold)
5. `folder = project_dir(...)`; ถ้า `Path(excel_path).resolve()` อยู่ใต้ `folder.resolve()` -> `ProjectError` บอกให้ย้ายไฟล์ Excel ออก / Relink ก่อน (**เลือก abort** ตาม AC-7) ไม่สร้าง zip
6. `stamp = now("%Y%m%d-%H%M%S")`; `zpath = _write_backup(...)` ถ้า `OSError` / `sqlite3.Error` -> `ProjectError("Backup failed, nothing deleted: ...") from exc` (AC-10)
   - เนื้อ zip: ถ้ามี `folder/tracker.db` -> copy ด้วย `sqlite3.backup` ไปไฟล์ชั่วคราวใน `delete-backups/` แล้วเขียนเป็น `tracker.db` (ปิดทั้งสอง connection ใน `finally`); ทุกไฟล์อื่นใน `folder.rglob("*")` ตาม path สัมพัทธ์ ยกเว้น `DB_FILES` ที่ชั้นบนสุด; `manifest.json` = ทุก field ของแถว project (`PROJECT_COLS`) ที่ระดับบนสุด + `deleted_at`, `deleted_by` (`person.editor`), `data_dir`
   - ไม่มี `folder` -> zip มีแค่ `manifest.json` (AC-15)
7. ถ้ามี `folder`: `tmp = folder.parent / f"{TEMP_PREFIX}{id}-{stamp}"`; `_rename(folder, tmp)` ถ้า `OSError` -> `raise DeleteLocked(...)` ข้อความ "files of this project are in use: close other app windows / Excel / Explorer and try again. Backup kept: <zip>" (AC-12)
8. `with appdb.transaction(): appdb.delete_project(id); appdb.log_event(action="deleted", detail=f"name={name}; excel={excel_path}; backup={zpath}")` ถ้า error ใดๆ -> `_rename(tmp, folder)` (ถ้า rename กลับไม่ได้ `log.error` และใส่ path ชั่วคราวในข้อความ) แล้ว `raise ProjectError("Could not update app.db, nothing deleted: ...") from exc` (AC-13)
9. `_rmtree(tmp)` ถ้า `OSError` -> `log.warning("... folder left: %s", tmp)` และ `leftover = tmp if tmp.exists() else None` ไม่ raise (AC-14)
10. `log.info(...)` แล้วคืน `DeleteResult`

ข้อควรระวัง: service ไม่เปิดไฟล์ Excel ไม่เรียก `validate_workbook` / `copy_shared` / `open_project` (AC-6) connection ไปที่ tracker.db ทั้งหมด (`_with_repo` ใน `require_manager`, ตัว backup) ต้องปิดก่อนข้อ 7

### UI
- `_project_card` (ภายใน `if self.manager:` และ `if arch:` เดิม): วาง `ub` (Restore selected) และ `self.del_btn = button(t("proj.delete"))` + `setStyleSheet(DANGER_QSS)` ใน `QHBoxLayout` เดียวกัน (AC-23)
- `_delete_archived`: `row = self.arch_table.currentRow()`; `row < 0` -> return (AC-27); `dlg = DeleteProjectDialog(app.base, app.person, self.arch_rows[row], self)`; ถ้า `dlg.exec()` -> ทำแบบ `_unarchive` (`app.win.retranslate()` ถ้ามี + `self.ctx.refresh_all()`) (AC-25)
- `DeleteProjectDialog`: หัวข้อ `proj.delete_title` (มีชื่อ), รายการลบ `proj.delete_list` (ใส่ path โฟลเดอร์จาก `what_is_deleted`), รายการที่ไม่แตะ `proj.delete_kept_list` (มี path Excel และคำว่าไม่ถูกแตะ), `proj.delete_backup` (path `delete-backups`), ช่อง `proj.delete_type`, ปุ่ม `self.go` (`proj.delete_go`, `DANGER_QSS`, disabled จนพิมพ์ชื่อตรงด้วยกติกา `strip()` เดียวกับ service) (AC-24)
  - `_go`: `run_in_worker(self, t("work.deleting"), job, done, failed, name="delete-project")` โดย `job` เปิด `projects.open_app_db(base)` เอง แล้ว `close()` ใน `finally`
  - `done(res)`: `self.result = res`; `info(self, t("proj.delete_done", p=res.backup_zip) + ("\n" + t("proj.delete_left", p=res.leftover) if res.leftover else ""))`; `self.accept()`
  - `failed(exc)`: `self.go.setEnabled(True)`; `error(self, friendly(exc))` dialog ไม่ปิด รายการ Archived ไม่ถูก refresh (AC-26)
- i18n key ใหม่ (EN + TH ไม่ว่าง, AC-28): `proj.delete`, `proj.delete_title`, `proj.delete_will`, `proj.delete_list`, `proj.delete_kept`, `proj.delete_kept_list`, `proj.delete_backup`, `proj.delete_type`, `proj.delete_go`, `proj.delete_done`, `proj.delete_left`, `work.deleting` ข้อความ error จาก service เป็นภาษาอังกฤษเหมือน `ProjectError` อื่นในระบบ

### CLI (`cli_project.py`)
- `d = ps.add_parser("delete", help="delete an ARCHIVED project for good (backup zip first; Excel file not touched)")`, `d.add_argument("project")`, `d.add_argument("--confirm-name", default=None)`
- ใช้ flow เดิม: ไม่มี profile -> 3, `find_project` ไม่เจอ -> "Unknown project" 2
- พิมพ์สิ่งที่จะถูกลบจาก `what_is_deleted` ก่อนเสมอ; ไม่มี `--confirm-name` -> พิมพ์บรรทัดสุดท้าย `NOT DONE: type the project name: --confirm-name "<name>"` return 2 (ไม่เรียก service)
- มี `--confirm-name` -> `delete_project(..., typed_name=args.confirm_name)` ชื่อผิด / ยังไม่ Archive / ไม่ใช่ manager -> `ProjectError` -> except เดิมพิมพ์ `NOT DONE: ...` return 2
- สำเร็จ: `Deleted <id> '<name>'. Backup: <zip>` (+ `Folder left, delete it yourself: <path>` ถ้ามี) return 0
- แก้ docstring module และ help ของ `project` ให้มี `delete`

## Data model / schema
- `app.db` `user_version` 1 -> 2 ด้วย `002_project_delete.sql` (trigger เท่านั้น ไม่เปลี่ยนตาราง) runner `_migrate` เดิมรองรับไฟล์ถัดไปอยู่แล้ว
- `project_event` ได้ action ใหม่ `"deleted"` แถวเก่าไม่ถูกแตะ (ไม่มี FK)
- ไฟล์ใหม่บนดิสก์: `<root>/delete-backups/<id>_<stamp>.zip` (ไม่ลบอัตโนมัติ) และอาจมี `<root>/projects/.deleting-<id>-<stamp>/` ค้างถ้าลบไม่หมด

## ความเสี่ยงและผลกระทบ
- **`require_manager` / `can_manage` เปิด tracker.db ของทุกโปรเจกต์** ถ้าโปรเจกต์ที่จะลบไม่มีโฟลเดอร์ (AC-15) `open_repository` จะสร้าง `projects/<id>/tracker.db` เปล่าขึ้นมาใหม่ระหว่างตรวจสิทธิ์ ผลคือ zip มี tracker.db เปล่า และโฟลเดอร์ถูก rename/ลบตามปกติ ผลลัพธ์ยังผ่าน AC-15 แต่ test ไม่ควร assert ว่า zip มี **แค่** `manifest.json` ถ้าไม่ได้ตรวจก่อน (ดู spec gap ในรายงาน) ไม่แก้ `_with_repo` เพราะนอก spec
- **Windows file lock** การ rename ล้มเหลวเมื่อมีโปรแกรมใดถือไฟล์ในโฟลเดอร์ (ตั้งใจ ใช้เป็นตัวตรวจ) แต่ OneDrive / antivirus อาจถือไฟล์ชั่วคราวทำให้ต้องกดซ้ำ ข้อความ error บอกให้ลองใหม่
- **sqlite connection ข้าม thread** worker ต้องเปิด `AppDb` ของตัวเอง ห้ามใช้ `app.appdb` ของ UI thread
- **test เดิม (AC-33)** `tests/test_projects.py::test_archive_hides_and_restores` ยิง `DELETE FROM project` ทั้งตารางขณะที่มีแถว `archived = 0` -> trigger ใหม่ยัง abort ด้วย `IntegrityError` test จึงยังผ่าน แค่คอมเมนต์ `# never deleted` ล้าสมัย `tests/test_auth.py:224` เป็น trigger ของ `app_user` ไม่เกี่ยว
- **ผู้ลบเสียสิทธิ์ manager** ถ้าโปรเจกต์ที่ลบเป็นที่เดียวที่เป็น Admin/PMO หลัง `refresh_all` ปุ่มต่างๆ จะหาย (พฤติกรรมตาม spec)
- **zip ใหญ่** ถ้า `backups/` ในโฟลเดอร์มี workbook backup มาก การลบใช้เวลานาน (รันใน worker แล้ว)
- **`[sync] backup_dir` ภายนอก** ไม่ถูกแตะ (A8) เพราะ service ทำงานเฉพาะ `projects/<id>/`

---

# 2026-11-25 — ดาวน์โหลดฟอร์ม Excel Progress เปล่า (Download blank form)

อ้างอิง: `docs/spec.md` section "2026-11-25 — ดาวน์โหลดฟอร์ม Excel Progress เปล่า" (D1-D8, AC-1 ถึง AC-17)

ผู้ใช้อนุมัติตามข้อเสนอทุกข้อ (C1-C5): Q1 ปุ่มอยู่ top bar หน้า Test Progress; Q2 ป้องกันเฉพาะไฟล์ Excel ของโปรเจกต์ปัจจุบัน; Q3 ชื่อเริ่มต้น `Planning Progress Form.xlsx`; Q4 ไม่เปิดไฟล์/โฟลเดอร์ แสดงแค่ข้อความพร้อม path; Q5 คง App ID ที่ซ่อนและ Report By Email / CC Email ไว้ (ผลเหมือน `create_workbook()`) **รอบนี้ผู้ใช้สั่งไม่เขียน test (ข้าม qa-tester)**

## ภาพรวมแนวทาง
1. **แยกส่วนสร้าง workbook** ใน `services/project_template.py` ออกจาก `create_workbook()` เป็น `_build_workbook() -> openpyxl.Workbook` (ย้ายโค้ดบรรทัด `tpl = template()` ถึงจบ loop sheet มาทั้งก้อน ไม่แก้ logic) `create_workbook()` เรียก helper นี้ ส่วนตรวจ `FileExistsError`, `mkdir`, temp `.creating.xlsx` และ `replace` คงเดิมทุกบรรทัด (AC-4)
2. **ฟังก์ชันใหม่ในไฟล์เดียวกัน** `export_blank_form(dest, project_excel)` ใช้ `_build_workbook()` ตัวเดียวกัน จึงได้ไฟล์เหมือน `create_workbook()` ทุกประการโดยอัตโนมัติ (AC-1, AC-2, AC-3, Q5) ไม่มี template แยก
3. **เขียนแบบปลอดภัย** save ลง temp ในโฟลเดอร์ปลายทาง (`tempfile.mkstemp(dir=dest.parent, prefix=".~form-", suffix=".xlsx")` ชื่อไม่ชนกัน) แล้ว `os.replace(tmp, dest)` ถ้า error ใดๆ ลบ temp ใน `except` แล้ว raise ต่อ (AC-5, AC-6)
4. **ปฏิเสธไฟล์ของโปรเจกต์** เทียบ path ก่อนสร้างอะไรทั้งสิ้น ด้วย `os.path.normcase(str(Path(p).resolve()))` (resolve จัดการ relative / `..`, normcase ทำให้ไม่สนตัวพิมพ์บน Windows) ตรงกัน -> raise `FormIsProjectFile` (AC-7)
5. **UI** ปุ่มใหม่ใน top bar หน้า Test Progress ก่อนปุ่ม Upload ใช้ `QFileDialog.getSaveFileName` แบบหน้า History และ `info` / `error` จาก `ui/dialogs.py` ไม่ใช้ worker thread (D7)

เหตุผล: ใช้ template และโค้ดสร้าง workbook ชุดเดียว (ไม่มีโอกาสหน้าตาสองไฟล์เพี้ยนกัน) service อยู่ที่เดียวกับ `create_workbook` ไม่แตะ DB / COM / auth (D5, D6) แก้ 3 ไฟล์ source เท่านั้น

## ไฟล์ที่แก้

| ไฟล์ | สิ่งที่เปลี่ยน |
|---|---|
| `services/project_template.py` | เพิ่ม `_build_workbook()`, `FormIsProjectFile`, `_same_file()`, `export_blank_form()`; `create_workbook()` เรียก `_build_workbook()` (พฤติกรรมเดิม); เพิ่มบรรทัดใน docstring module |
| `ui/screens/test_progress.py` | `self.form_btn` ใน top bar (ก่อน `self.up_btn`), ตั้งข้อความใน `refresh()`, method `_download_form()` |
| `ui/i18n.py` | key `tp.form`, `tp.form_saved`, `tp.form_is_project`, `tp.form_failed` (EN, TH) ตามตาราง D8 ของ spec |

## Interface สำคัญ

```python
# services/project_template.py
FORM_FILENAME = "Planning Progress Form.xlsx"     # default name in the save dialog (Q3)


class FormIsProjectFile(ValueError):
    """The chosen path is the current project's Excel file (D4). Nothing was written."""


def _build_workbook() -> openpyxl.Workbook:
    """Blank workbook from template(): 3 sheets, headers, styles, validations, cond formats, legend. Not saved."""


def _same_file(a: Path, b: Path) -> bool:
    """True when a and b point to the same path: resolve() + os.path.normcase (case-insensitive on Windows)."""


def create_workbook(path: Path) -> Path:
    """Unchanged behaviour: FileExistsError if path exists; temp '<stem>.creating.xlsx' then replace."""


def export_blank_form(dest: Path, project_excel: Path | None) -> Path:
    """Save a blank form at dest, overwriting an existing file (the save dialog already asked). Returns dest.

    Raises FormIsProjectFile when dest is project_excel (checked first; nothing written, no temp file).
    Raises OSError (e.g. PermissionError) when it cannot be saved; the old dest is untouched, the temp is removed.
    """
```

### ลำดับใน `export_blank_form`
1. `dest = Path(dest)`; ถ้า `project_excel` ไม่ใช่ None และ `_same_file(dest, project_excel)` -> `raise FormIsProjectFile(str(dest))`
2. `wb = _build_workbook()`
3. `fd, name = tempfile.mkstemp(dir=dest.parent, prefix=".~form-", suffix=".xlsx")`; `os.close(fd)`; `tmp = Path(name)` (ถ้า `mkstemp` เองล้มเหลวเพราะโฟลเดอร์เขียนไม่ได้ -> `OSError` ส่งออกไป ไม่มี temp ค้าง)
4. `try: wb.save(tmp); os.replace(tmp, dest)` / `except BaseException: tmp.unlink(missing_ok=True)` (ห่อด้วย `try` กัน error ตอนลบ) แล้ว `raise`
5. `return dest`

ไม่สร้างโฟลเดอร์ปลายทาง (dialog เลือกโฟลเดอร์ที่มีอยู่แล้ว) ไม่เติมนามสกุลใน service (UI ทำ)

### UI (`ui/screens/test_progress.py`)
- `__init__`: `self.form_btn = button("")`; `self.form_btn.clicked.connect(self._download_form)`; `top.addWidget(self.form_btn)` ก่อน `top.addWidget(self.up_btn)` ไม่ซ่อนตาม role (D6, AC-9)
- `refresh()`: `self.form_btn.setText(t("tp.form"))`
- `_download_form()`:
  1. `path, _ = QFileDialog.getSaveFileName(self, t("tp.form"), FORM_FILENAME, "Excel (*.xlsx)")`; `if not path: return` (AC-10, AC-13)
  2. `if not path.lower().endswith(".xlsx"): path += ".xlsx"` (AC-12)
  3. `try: out = export_blank_form(Path(path), self.ctx.cfg.excel_path)`
     - `except FormIsProjectFile: error(self, t("tp.form_is_project")); return` (AC-14)
     - `except OSError: error(self, t("tp.form_failed", p=path)); return` (AC-15)
  4. `info(self, t("tp.form_saved", p=str(out)))` ไม่เปิดไฟล์ / โฟลเดอร์ (AC-11, Q4)

  หมายเหตุ: ถ้าเติม `.xlsx` เอง Qt จะไม่ได้ถามยืนยันเขียนทับสำหรับชื่อใหม่นั้น (ยอมรับตาม spec D7 ที่ให้เติมนามสกุล)

## Data model / schema
ไม่เปลี่ยน ไม่แตะ tracker.db / app.db / config (D5, AC-8)

## ความเสี่ยงและผลกระทบ
- **refactor `create_workbook`** เป็นการย้ายโค้ดล้วน ถ้าย้ายพลาดจะกระทบการสร้างโปรเจกต์ใหม่ (`services/projects.py`) ต้องคงส่วน `FileExistsError` / `.creating.xlsx` / `replace` ไว้ใน `create_workbook` ตามเดิม
- **ไม่มี test รอบนี้** (ผู้ใช้สั่ง) AC ทั้งหมดไม่ได้ถูกยืนยันด้วย pytest ใหม่ และ AC-17 ไม่ได้ถูกตรวจโดย qa-tester แนะนำให้ developer รัน `pytest -q` ชุดเดิมหลัง T29 อย่างน้อยหนึ่งครั้ง (เพื่อยืนยันว่า refactor ไม่ทำ test เดิมพัง) และให้ผู้ใช้ลองกดปุ่มเองหนึ่งครั้ง
- **OneDrive / antivirus** อาจล็อกไฟล์ temp ชั่วขณะทำให้ `os.replace` ล้มเหลวเป็น `PermissionError` -> ผู้ใช้เห็น `tp.form_failed` และลองใหม่ได้ ไฟล์เดิมไม่เสีย
- **ป้องกันเฉพาะโปรเจกต์ปัจจุบัน** (Q2) ถ้าผู้ใช้เลือกไฟล์ Excel ของโปรเจกต์อื่น จะถูกเขียนทับได้ (หลังยืนยันใน dialog) ตามที่ผู้ใช้ตัดสิน
- **`.ctx.cfg.excel_path`** ถ้าหน้า Test Progress เปิดได้ แปลว่ามี cfg ของโปรเจกต์อยู่แล้ว ไม่ต้องจัดการกรณี None ใน UI (service รับ None ได้อยู่แล้ว)

---

# 2026-11-26 — นำเข้า / เปลี่ยนไฟล์ Excel ของโปรเจกต์จากลิงก์ SharePoint

อ้างอิง: `docs/spec.md` section "2026-11-26 — นำเข้า / เปลี่ยนไฟล์ Excel ของโปรเจกต์จากลิงก์ SharePoint" (E-1 ถึง E-15, AC-1 ถึง AC-38) และ "การตัดสินใจของผู้ใช้ (26-Nov-2026)" ท้าย section: **Q1 ไม่มีปุ่มแยก** ปุ่ม `proj.import` เดิมถามก่อนว่า "เลือกไฟล์…" หรือ "วางลิงก์ SharePoint…"; Q2 ลิงก์อยู่ใน event detail เท่านั้น; Q3/Q4 banner และ CLI นอกขอบเขต; Q5 สิทธิ์ `require_manager`; Q6 มีหน้าต่างยืนยัน; E2, E6 ตามข้อเสนอ

## ภาพรวมแนวทาง
1. **resolver ของไฟล์ต่อยอดของเดิม ไม่แก้ของเดิม** ฟังก์ชันใหม่ `resolve_file_link(link, libraries=None)` ใน `services/sharepoint_link.py` ทำการตรวจเฉพาะของไฟล์ก่อน (Office Online ที่ไม่มี path -> `link_sharing`, `.`/`..`, ต้องลงท้าย `.xlsx`) แล้ว **เรียก `resolve_link(...)` เดิม** เพื่อจับคู่ library (prefix ยาวที่สุด + `MountPoint` + ส่วนที่เหลือ) error `AttachmentDirError` จาก `link_location` / `resolve_link` ถูกแปลงเป็น `ExcelLinkError` code เดิม ทำให้ `resolve_link`, `link_location`, `resolve_attachment_link`, `set_attachment_link` ไม่ถูกแตะเลย (AC-12)
2. **exception ใหม่แยกจากไฟล์แนบ** `ExcelLinkError(ProjectError)` ใน `services/projects.py` (โครงเดียวกับ `AttachmentDirError`: `code`, `path`) **ไม่** subclass `AttachmentDirError` เพื่อไม่ให้ `except AttachmentDirError` ในหน้า Settings แปลเป็นข้อความ "โฟลเดอร์" ผิด UI แปลด้วย `t(f"proj.xl_{code}", p=path)` (E-15) เป็น subclass ของ `ProjectError` จึงไม่มีที่ไหนที่ catch `ProjectError` อยู่แล้วต้องแก้
3. **service ใช้ทางเดิม** แยกตัวกลางของ `import_project` / `relink` ออกเป็น `_import_file` / `_relink_file` (ย้ายโค้ดล้วน) ที่รับ `link_url: str | None` ฟังก์ชันเดิมเรียกด้วย `None` (detail และพฤติกรรม `OSError` เหมือนเดิมทุกตัวอักษร AC-27) ฟังก์ชันใหม่ `import_project_link` / `relink_link` = ตรวจสิทธิ์ (+ หาโปรเจกต์) -> `resolve_excel_link` -> ตัวกลางเดียวกันด้วย `link_url` ทุกการตรวจ (workbook, ชื่อไฟล์ซ้ำ, ชื่อโปรเจกต์, App ID 80%) จึงเป็นโค้ดชุดเดียว
4. **อ่านไฟล์ไม่ได้ (E-8)** เฉพาะเมื่อ `link_url` ไม่ใช่ None: `OSError` จาก `validate_workbook` / `app_ids_in` (ทั้งคู่ผ่าน `copy_shared`) -> `ExcelLinkError("link_unreadable", path)` ห่อเฉพาะสองจุดนี้ (ไม่ห่อการเขียน app.db / tracker.db เพื่อไม่ให้แปลผิดว่า "ไม่มีอะไรถูกบันทึก")
5. **UI** ปุ่ม `proj.import` เดิม (sidebar + `NoProjectWindow` ทั้งคู่เรียก `AppController.import_project`) -> `import_choice()` (QMessageBox 2 ปุ่ม + Cancel) -> `import_project_flow` เดิม หรือ `import_link_flow` ใหม่ ส่วนอ่านไฟล์ครั้งแรกใน worker แยกออกจาก `import_project_flow` เป็น `_read_new_project` ให้ทั้งสอง flow ใช้ หน้า Settings เพิ่มปุ่ม `proj.relink_link` ข้างปุ่ม Relink -> `relink_link_flow` (โครงเดียวกับ `relink_flow` + getText + confirm)

เหตุผล: กลไกแปลงลิงก์ / ตรวจไฟล์ / สร้างโปรเจกต์ / Relink มีอยู่ครบแล้ว งานรอบนี้คือ "ด่านหน้า" สำหรับไฟล์ + ส่งต่อเข้าทางเดิม ไม่มี dependency ใหม่ ไม่เปลี่ยน schema ทุก logic ทดสอบได้ด้วย pytest โดย inject / monkeypatch รายการ `SyncedLibrary` (ไม่อ่าน registry) และ `tmp_path` เป็น MountPoint

## ไฟล์ที่สร้าง / แก้

| ไฟล์ | สิ่งที่เปลี่ยน |
|---|---|
| `services/sharepoint_link.py` | เพิ่ม `_OFFICE_PAGES`, `resolve_file_link()`; บรรทัดใน docstring module (ลิงก์ไฟล์) ไม่แก้ฟังก์ชันเดิม |
| `services/projects.py` | เพิ่ม `ExcelLinkError`, `resolve_excel_link()`, `validate_link_file()`, `import_project_link()`, `relink_link()`, helper `_validate()`, `_ids()`, `_source()`; แยก `_import_file()` / `_relink_file()` ออกจาก `import_project` / `relink` (พฤติกรรมเดิม); บรรทัดใน docstring module |
| `ui/projects.py` | เพิ่ม `import_choice()`, `_import_choice_box()`, `link_error_text()`, `import_link_flow()`, `relink_link_flow()`; แยก `_read_new_project()` ออกจาก `import_project_flow` (พฤติกรรมเดิม); import `QInputDialog`, `QLineEdit` (มีแล้ว), `QMessageBox` |
| `ui/app.py` | `AppController.import_project()` เรียก `import_choice()` ก่อน |
| `ui/screens/settings.py` | `_project_card`: ปุ่ม `self.relink_link_btn` ข้างปุ่ม Relink; method `_relink_link()` |
| `ui/i18n.py` | key ตามตาราง E-15 + `proj.import_how`, `proj.import_file` (EN + TH) |
| `tests/test_sharepoint_link.py` (ใหม่, QA) | resolver ของไฟล์ + characterization ของ `resolve_link` / `link_location` เดิม (AC-12) |
| `tests/test_excel_link.py` (ใหม่, QA) | service (a)/(b), UI, i18n |

## Interface สำคัญ

```python
# services/projects.py
class ExcelLinkError(ProjectError):
    """A SharePoint link that can't be used as a project's Excel file. `code` picks the UI text proj.xl_<code>.
    Codes: link_invalid, link_sharing, link_not_xlsx, link_not_synced, link_not_found, link_unreadable.
    `path` = cleaned link (first four) or the local path (link_not_xlsx for a folder, link_not_found,
    link_unreadable)."""

    def __init__(self, code: str, path: Path | str, message: str) -> None: ...
    code: str
    path: str


def resolve_excel_link(link: str, libraries: list[SyncedLibrary] | None = None) -> tuple[str, Path]:
    """-> (cleaned link, local .xlsx path that exists). resolve_file_link + exists / not a folder (E-7).
    Reads no file content. Raises ExcelLinkError."""


def validate_link_file(base: Config, path: Path) -> list[str]:
    """validate_workbook(path, base.header_overrides); OSError -> ExcelLinkError('link_unreadable', path)."""


def import_project(base: Config, appdb: AppDb, person: Person, *, path: Path) -> dict:      # unchanged behaviour
def relink(base: Config, appdb: AppDb, person: Person, project_id: str, new_path: Path) -> str:  # unchanged


def import_project_link(base: Config, appdb: AppDb, person: Person, *, link: str) -> dict:
    """(a) E-9: require_manager -> resolve_excel_link -> _import_file(..., link_url=url). Raises ProjectError /
    ExcelLinkError; nothing saved on error."""


def relink_link(base: Config, appdb: AppDb, person: Person, project_id: str, link: str) -> str:
    """(b) E-10: require_manager -> project exists -> resolve_excel_link -> _relink_file(..., link_url=url).
    Returns the same summary as relink()."""


def _import_file(base: Config, appdb: AppDb, person: Person, path: Path, link_url: str | None) -> dict:
    """Body of import_project after require_manager (moved code). detail = f"existing file {path}" + _source(link_url)."""


def _relink_file(base: Config, appdb: AppDb, person: Person, project: dict, new_path: Path,
                 link_url: str | None) -> str:
    """Body of relink after the project lookup (moved code).
    detail = f"{old} -> {new_path} ({summary})" + _source(link_url)."""


def _validate(base: Config, path: Path, link_url: str | None) -> list[str]:
    """validate_workbook; OSError re-raised as is when link_url is None, else ExcelLinkError('link_unreadable')."""


def _ids(path: Path, link_url: str | None) -> set[str]:
    """app_ids_in; same OSError rule as _validate."""


def _source(link_url: str | None) -> str:
    """'' or f' (SharePoint: {link_url})'."""


# services/sharepoint_link.py
_OFFICE_PAGES = ("doc.aspx", "wopiframe.aspx")

def resolve_file_link(link: str, libraries: list[SyncedLibrary] | None = None) -> Path:
    """Local path of a SharePoint .xlsx link (existence NOT checked). Raises ExcelLinkError
    (link_invalid / link_sharing / link_not_xlsx / link_not_synced)."""


# ui/projects.py
def _import_choice_box(parent) -> tuple[QMessageBox, QPushButton, QPushButton]:
    """Box text proj.import_how; buttons proj.import_file, proj.import_link (+ Cancel). Not shown."""

def import_choice(parent) -> str | None:
    """'file' / 'link' / None (Cancel / closed)."""

def link_error_text(exc: ExcelLinkError) -> str:
    """t(f'proj.xl_{exc.code}', p=exc.path)."""

def _read_new_project(parent, base, person, project: dict, on_done=None) -> None:
    """First read of a just-created project in a worker, then info(proj.imported) + on_done(project) (moved code)."""

def import_project_flow(parent, base, appdb, person, on_done=None) -> dict | None:    # unchanged behaviour
def import_link_flow(parent, base, appdb, person, on_done=None) -> dict | None: ...
def relink_link_flow(parent, base, appdb, person, project: dict, on_done=None): ...


# ui/screens/settings.py
class SettingsScreen:
    def _relink_link(self) -> None: ...      # self.relink_link_btn.clicked
```

### ลำดับใน `resolve_file_link` (ทุกข้อไม่แตะดิสก์)
1. `text = clean_link(link)`
2. `host, segs = link_location(text)` — `AttachmentDirError` (`link_invalid` / `link_sharing`) -> `raise ExcelLinkError(exc.code, text, <ข้อความไฟล์>) from exc` (AC-5 `/:x:/s/`, `/:x:/g/`; AC-8)
3. ถ้า segment สุดท้ายของ `urlsplit(text).path` (หลัง `unquote`, casefold) อยู่ใน `_OFFICE_PAGES` -> `ExcelLinkError("link_sharing", text, ...)` (AC-5 `Doc.aspx` / `WopiFrame.aspx`; ตรวจจาก path ของ URL ไม่ใช่ `segs` เพราะ `segs` อาจมาจาก `?id=`; ไม่อ่าน `file=` ตาม E-3)
4. ถ้ามี segment ที่ `strip()` เป็น `.` หรือ `..` -> `ExcelLinkError("link_invalid", text, ...)` (AC-11; ตรวจก่อนข้อ 5 ตามลำดับของ `resolve_link`)
5. `if not segs or not segs[-1].casefold().endswith(".xlsx")` -> `ExcelLinkError("link_not_xlsx", text, ...)` **ก่อน**หา library (AC-6, AC-7; ลิงก์โฟลเดอร์ `/:f:/r/...`, หน้า library ที่ `link_location` ตัด `Forms/*.aspx` ทิ้ง, `.xls`/`.xlsm`/`.csv` เข้าข้อนี้)
6. `return resolve_link(text, libraries)` — `AttachmentDirError("link_not_synced")` -> `ExcelLinkError("link_not_synced", text, ...)` (AC-1 ถึง AC-4, AC-9) (code อื่นจาก `resolve_link` ไม่ควรเกิดหลังข้อ 2-4 ถ้าเกิดให้แปลง code เดิมเหมือนกัน)

`ExcelLinkError` import จาก `.projects` (บรรทัด import เดิมมี `AttachmentDirError` อยู่แล้ว ไม่เกิด import วน เพราะ `projects.py` import `sharepoint_link` แบบ lazy ในฟังก์ชัน)

### ลำดับใน `resolve_excel_link`
1. `from .sharepoint_link import clean_link, resolve_file_link` (lazy แบบ `resolve_attachment_link`)
2. `url = clean_link(link)`; `path = resolve_file_link(url, libraries)`
3. `if not path.exists()` -> `ExcelLinkError("link_not_found", path, ...)` (AC-10 ทั้งกรณีโฟลเดอร์แม่มี / ไม่มี; ไม่ `mkdir` ไม่สร้างอะไร)
4. `if path.is_dir()` -> `ExcelLinkError("link_not_xlsx", path, ...)` (E-7 path ที่เป็นโฟลเดอร์)
5. `return url, path`

### ลำดับใน `import_project_link` / `relink_link`
- (a) `require_manager` (AC-15 ก่อนแปลงลิงก์) -> `url, path = resolve_excel_link(link)` -> `return _import_file(base, appdb, person, path, url)`
- (b) `require_manager` -> `project = appdb.get_project(project_id)` (`None` -> `ProjectError(f"Unknown project {project_id}")` ข้อความเดิม) -> `resolve_excel_link(link)` -> `return _relink_file(base, appdb, person, project, path, url)`
- `_import_file`: `_validate(...)` -> problems -> `ProjectError("This file can't be imported:\n- ...")` (ข้อความเดิม) -> ชื่อไฟล์ซ้ำ (ข้อความเดิม) -> `name_from_file` -> `_new_project(..., "imported", f"existing file {path}" + _source(link_url))`
- `_relink_file`: `_validate` -> problems "can't be linked" -> ชื่อซ้ำ -> `known` / `_ids(new_path, link_url)` / 80% (ข้อความเดิม) -> transaction เดิม detail `f"{project['excel_path']} -> {new_path} ({summary})" + _source(link_url)`
- `excel_path` ที่บันทึก = `str(path)` จาก resolver (ไม่ resolve symlink, ไม่ copy / rename ไฟล์) (AC-13, AC-20, AC-21)

### UI
- **ตัวเลือกวิธีนำเข้า (Q1)** `AppController.import_project`: `how = import_choice(self.win)`; `"file"` -> `import_project_flow(...)` เดิม; `"link"` -> `import_link_flow(self.win, self.base, self.appdb, self.person, on_done=lambda p: self.open_project(p["id"]))`; `None` -> ไม่ทำอะไร ปุ่มใน sidebar / `NoProjectWindow` ไม่แก้ (เงื่อนไขแสดงตาม `can_manage` เดิม) ดังนั้น AC-28 ตีความเป็น: ปุ่ม `proj.import` แสดงเมื่อ manager (เดิม) และกล่องที่ `_import_choice_box` สร้างมีปุ่ม `t("proj.import_link")`
- `_import_choice_box(parent)`: `QMessageBox(QMessageBox.Question, t("app"), t("proj.import_how"), parent=parent)`; `f = box.addButton(t("proj.import_file"), QMessageBox.AcceptRole)`; `k = box.addButton(t("proj.import_link"), QMessageBox.AcceptRole)`; `box.addButton(QMessageBox.Cancel)`; `import_choice` -> `box.exec()` แล้วดู `box.clickedButton()`
- `import_link_flow`:
  1. `link, ok = QInputDialog.getText(parent, t("proj.xl_link_title"), t("proj.xl_link_prompt"))`; `not ok or not link.strip()` -> `return None` (AC-30)
  2. `try: url, path = projects.resolve_excel_link(link); problems = projects.validate_link_file(base, path)` / `except projects.ExcelLinkError as exc: error(parent, link_error_text(exc)); return None` (AC-31); `problems` -> `error(parent, "\n".join(problems))` return None
  3. `confirm(parent, t("proj.import_link_q", u=url, p=path, n=projects.name_from_file(path)))` ไม่ -> return None (AC-29, AC-30)
  4. `p = projects.import_project_link(base, appdb, person, link=url)`; `ExcelLinkError` -> `link_error_text`; `ProjectError` อื่น -> `error(parent, str(exc))`; return None
  5. `_read_new_project(parent, base, person, p, on_done)`; `return p` (AC-29 `info(proj.imported)` + `open_project`)
- `relink_link_flow` (โครงเดียวกับ `relink_flow`):
  1. `getText` ชื่อเรื่องและ prompt เดียวกับข้างบน ยกเลิก/ว่าง -> `return None` ไม่เรียก service (AC-34)
  2. `url, path = projects.resolve_excel_link(link)` ใน UI thread (ไม่อ่านเนื้อไฟล์) `ExcelLinkError` -> `error(link_error_text)` return None
  3. `confirm(parent, t("proj.relink_link_q", u=url, p=path, n=project["name"]))` ไม่ -> return None
  4. `job`: `db = projects.open_app_db(base)` -> `projects.relink_link(base, db, person, project["id"], url)` -> `db.close()` ใน `finally`
  5. `done(summary)`: `info(parent, t("proj.relinked", s=summary))` + `on_done()`; `failed(exc)`: `error(parent, link_error_text(exc) if isinstance(exc, projects.ExcelLinkError) else friendly(exc))` (AC-34 ข้อความ ProjectError ที่ไม่มี code = `str(exc)`)
  6. `run_in_worker(parent, t("work.relinking"), job, done, failed, name="relink-link")`
- Settings `_project_card`: หลัง `row.addWidget(rl)` เพิ่ม `self.relink_link_btn = button(t("proj.relink_link"))`, `clicked -> self._relink_link`, `setVisible(self.manager)` (AC-32 เงื่อนไขเดียวกับ `rl`); `_relink_link` = `_relink` แต่เรียก `relink_link_flow(..., on_done=lambda: app.open_project(pid))` (AC-33)

### i18n (E-15)
key ทั้ง 12 ตัวในตาราง E-15 ของ spec (`proj.import_link`, `proj.relink_link`, `proj.xl_link_title`, `proj.xl_link_prompt`, `proj.import_link_q`, `proj.relink_link_q`, `proj.xl_link_invalid`, `proj.xl_link_sharing`, `proj.xl_link_not_xlsx`, `proj.xl_link_not_synced`, `proj.xl_link_not_found`, `proj.xl_link_unreadable`) และเพิ่มสำหรับ Q1:

| key | EN | TH |
|---|---|---|
| `proj.import_how` | How do you want to import the Excel file? | ต้องการนำเข้าไฟล์ Excel แบบไหน? |
| `proj.import_file` | Choose file… | เลือกไฟล์… |

`proj.import_link` ใช้เป็นข้อความปุ่มที่สองในกล่องเลือก (ข้อความตามตาราง spec) code ทุกตัวของ `ExcelLinkError` ต้องมี key `proj.xl_<code>` (AC-35) ไม่แตะ `proj.att_link_*` (AC-37) หมายเหตุ: `proj.xl_link_not_synced` (TH) อ้าง "นำเข้าไฟล์ Excel" ซึ่งตอนนี้คือปุ่มเดียวกันที่มีตัวเลือก "เลือกไฟล์…" ความหมายยังถูก

## Data model / schema
ไม่เปลี่ยน schema (`app.db` / `tracker.db`) ไม่มี column ลิงก์ (Q2) ลิงก์ที่ clean แล้วอยู่ใน `project_event.detail` ของ action `imported` / `relinked` เท่านั้น `excel_path` เป็น path ในเครื่องธรรมดา

## ความเสี่ยงและผลกระทบ
- **refactor `import_project` / `relink`** ย้ายโค้ดล้วน ถ้าพลาดกระทบการนำเข้า / Relink แบบเดิม ต้องคงข้อความ `ProjectError` ทุกข้อความ, ลำดับการตรวจ และ detail เดิม (`link_url=None` -> `_source` คืน `""`) และให้ `OSError` แบบเดิมยังหลุดออกไปเหมือนเดิมเมื่อไม่ใช่ลิงก์ test เดิมใน `tests/test_projects.py` ต้องผ่านโดยไม่แก้ (AC-27)
- **refactor `import_project_flow`** การย้ายส่วน worker ไป `_read_new_project` ต้องคง `name="import-project"`, ข้อความ `work.reading` และ `on_done(p)`
- **ตรวจไฟล์สองรอบตอนนำเข้า** UI เรียก `validate_link_file` ก่อน confirm แล้ว service ตรวจอีกรอบ (เหมือน flow เลือกไฟล์เดิม) ไฟล์ใน OneDrive ที่เป็น placeholder ถูก copy สองครั้ง ช้าขึ้นเล็กน้อย การตรวจรอบแรกอยู่บน UI thread เหมือนเดิม (ไม่ใช้ worker เพื่อไม่ขยาย scope)
- **`link_unreadable` แคบ** แปลงเฉพาะ `OSError` จาก `validate_workbook` / `app_ids_in` (รวม error ของ `tempfile` ในสองฟังก์ชันนี้ ซึ่งจะถูกรายงานว่า "อ่านไฟล์จาก OneDrive ไม่ได้" ด้วย ยอมรับได้) `OSError` ตอนสร้าง tracker.db ของโปรเจกต์ใหม่ (หลังบันทึกแถวใน app.db แล้ว) ยังหลุดออกไปแบบเดิม
- **ไม่มี test เดิมของ `sharepoint_link.py`** T31 ต้องเขียน characterization test ของ `resolve_link` / `link_location` (รวม `/:f:/r/`, `Forms/AllItems.aspx` ไม่มี `?id=`, `?RootFolder=`) ก่อนหรือพร้อมเพิ่ม `resolve_file_link` เพื่อยืนยัน AC-12
- **AC-10 "ไม่สร้างไฟล์/โฟลเดอร์"** `require_manager` เปิด tracker.db ของโปรเจกต์ที่มีอยู่ (`_with_repo`) ซึ่งอาจสร้างไฟล์ใต้ `<root>/projects/` ได้ test ควร assert ว่าไม่มีอะไรใหม่ใต้ MountPoint (`tmp_path` ของ OneDrive) ไม่ใช่ทั้ง `tmp_path`
- **ตัวพิมพ์ใน path** Windows ไม่สนตัวพิมพ์ `exists()` จึงผ่านแม้ลิงก์ต่างตัวพิมพ์จากโฟลเดอร์จริง `excel_path` ที่บันทึกจะเป็นตัวพิมพ์ตามลิงก์ (ไม่ใช่ตามดิสก์) ไม่มีผลต่อการใช้งานบน Windows
- **ชื่อไฟล์ / โฟลเดอร์ที่ OneDrive เปลี่ยนเอง** (เช่นอักขระที่ Windows ไม่รับ) จะได้ `link_not_found` ผู้ใช้ใช้ปุ่มเลือกไฟล์แทน (นอกขอบเขต)
- **กล่องเลือกวิธีนำเข้า** เพิ่มหนึ่งคลิกให้ผู้ใช้ที่เลือกไฟล์แบบเดิม (ตามที่ผู้ใช้ตัดสิน Q1) ไม่มี test เดิมของ flow นี้ที่ต้องแก้ (ตรวจแล้ว `tests/` ไม่มีการเรียก `import_project_flow` / `imp_proj_btn`)
- **README** spec ไม่ได้สั่งให้แก้ README รอบนี้จึงไม่มี task README (หัวหน้าตัดสินถ้าต้องการ)
