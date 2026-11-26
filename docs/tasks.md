# Tasks — งานที่ยังค้าง

สถานะ: `[ ]` ยังไม่เริ่ม, `[~]` กำลังทำ, `[x]` เสร็จและ test ผ่าน, `[!]` ติดปัญหา

## ลบโปรเจกต์ (Delete project) (25-Nov-2026) — spec/design section "2026-11-25 — ลบโปรเจกต์"

> AC ในหัวข้อนี้หมายถึง AC ของ spec section "ลบโปรเจกต์" (ไม่ใช่ AC-1..23 ของ Weekly Report ด้านบน)

- [~] T20: migration 002 + AppDb (delete_project, next_project_id ไม่ใช้ id ซ้ำ)
  - ไฟล์: storage/migrations_app/002_project_delete.sql (ใหม่), storage/app_db.py, tests/test_delete_project.py (ใหม่)
  - ทำอะไร: trigger `project_no_delete` ใหม่ที่ abort เฉพาะ `OLD.archived = 0` (ไม่แก้ 001_app.sql); `AppDb.delete_project(id)`; `next_project_id()` นับ max จาก `project.id` รวม `project_event.project_id`; test: app.db เวอร์ชัน 1 ที่มีข้อมูล -> เปิดแล้ว user_version = 2 ข้อมูลครบ, DELETE ตรงๆ แถว archived=0 abort / archived=1 ลบได้, id ที่เหลือแค่ใน project_event ไม่ถูกใช้ซ้ำ
  - AC ที่ครอบคลุม: AC-16, AC-17, AC-19 (ส่วน AppDb)
  - ⚠️ 26-Nov: test AC-16 4 ข้อ fail (`assert 4 == 2`) เพราะ test ตรวจ `user_version == 2` ตายตัว แต่ตอนนี้มี migration 003/004 แล้ว version จริงเป็น 4

- [ ] T28: รัน test ทั้งหมด และรายงาน test เดิมที่แก้
  - ไฟล์: tests/test_projects.py (เฉพาะคอมเมนต์ `# never deleted` บรรทัด ~463 ถ้าจะแก้)
  - ทำอะไร: `pytest -q` ทั้งหมด; ยืนยันว่า `test_archive_hides_and_restores` ยังผ่าน (DELETE ทั้งตารางชนแถว archived=0 -> IntegrityError) ห้ามลบ test; รายงานหัวหน้าว่าแก้ test ใด
  - AC ที่ครอบคลุม: AC-33

## นำเข้า / เปลี่ยนไฟล์ Excel จากลิงก์ SharePoint (26-Nov-2026) — spec/design section "2026-11-26 — นำเข้า / เปลี่ยนไฟล์ Excel ของโปรเจกต์จากลิงก์ SharePoint"

> AC ในหัวข้อนี้หมายถึง AC ของ spec section 2026-11-26 (AC-1..38) test ทุกข้อใช้ `tmp_path` เป็น MountPoint, รายการ `SyncedLibrary` ปลอม (ส่งเข้า `resolve_file_link(..., libraries)` หรือ monkeypatch `uat_tracker.services.sharepoint_link.synced_libraries`) ห้ามอ่าน registry จริง ห้ามใช้ Excel / COM workbook สร้างด้วย `build_workbook` / `_wb` แบบ `tests/test_projects.py` ไฟล์ปลอม `L` / `F` ตามตัวอย่างใน spec

- [ ] T38: รัน test ทั้งหมด และรายงาน test เดิมที่แก้
  - ไฟล์: (ไม่มี source — แก้เฉพาะถ้า test ไม่ผ่าน)
  - ทำอะไร: `pytest -q` ทั้งหมด (test ฝั่ง Qt ต้องรันบน Windows ที่มี PySide6); ยืนยันว่า test เดิมของ projects / settings / ui ผ่านโดยไม่แก้; ห้ามลบ test เดิม ถ้าต้องแก้ต้องรายงานหัวหน้าว่าแก้ test ใดเพราะอะไร; แยกรายงาน failure ที่มีอยู่ก่อนรอบนี้ (ดูหมายเหตุ T29 / T30) ออกจาก failure ใหม่
  - AC ที่ครอบคลุม: AC-38
