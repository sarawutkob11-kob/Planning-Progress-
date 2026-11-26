# Spec — uat_tracker

> หมายเหตุ: ไฟล์นี้สร้างครั้งแรกในรอบนี้ งานรอบก่อน (T1–T12) ใช้ "ต้องการ" ใน `docs/tasks.md` เป็นเกณฑ์ และผลตรวจอยู่ใน `docs/review.md` รอบถัดไปให้เพิ่ม section ใหม่ต่อท้าย ห้ามเขียนทับ section เดิม

---

# 2026-11-24 — Sync bug fixes B1–B5

> ปรับปรุง 2026-11-24: ใส่คำตอบของผู้ใช้สำหรับ Q1–Q6 แล้ว (ดู "การตัดสินใจของผู้ใช้" ท้าย section)

## เป้าหมาย
แก้บั๊ก 5 ข้อในระบบ sync 2 ทาง (app <-> Excel ผ่าน COM) ของ `src/uat_tracker/` ที่ทำให้:
- ค่าที่ผู้ใช้แก้ในแอประหว่างที่ sync กำลังเขียนลง Excel ไม่ถูกเขียนลง Excel อีกเลย (DB กับ Excel ไม่ตรงกันถาวร) — B1
- เปลี่ยนช่วงวันที่ (date window) หรือ Excel path ใน Settings แล้ว sync ยังใช้ค่าเก่า ทำให้วันที่ถูกล้างและเกิด audit "Edited in Excel" ปลอม หรือยังอ่าน/เขียนไฟล์เดิม — B2
- sync thread ตายเงียบเมื่อ DB ถูก lock แต่ chip ยังแสดง "Synced" — B3
- ชีต/หัวคอลัมน์ใน Excel หาย ทำให้สร้าง backup ใหม่ทุก 1 วินาทีไม่สิ้นสุด และไม่มีสถานะ retry — B4
- สลับโปรเจกต์หรือปิดแอประหว่าง sync รอบยาว ทำให้แอป crash (qFatal) และอาจทิ้ง EXCEL.EXE ที่ซ่อนอยู่ค้างถือไฟล์ — B5

ผู้ได้ประโยชน์: PMO และผู้ใช้ทุกคนที่แก้ข้อมูลในแอปแล้วคาดหวังว่าค่าจะไปถึง Excel ที่แชร์บน OneDrive รวมถึงผู้ดูแลระบบที่ต้องเชื่อสถานะ sync ที่แสดงบนจอได้

## ลำดับการแก้ (บังคับ)
B1 -> B2 -> B3 -> B4 -> B5 (B5 ทำเป็นข้อสุดท้าย)

## ขอบเขต (In scope)
- **B1** (เดิม #1, High) — `services/sync.py` `_flush` :298-305, `_retry` :365-379, `_fail` :355-363 และ `storage/sqlite_repo.py` `set_pending_state` (รวม Protocol ใน `storage/repository.py`): เพิ่ม version guard ให้การเปลี่ยนสถานะ pending มีผลเฉพาะกับแถวที่ยังเป็นเวอร์ชันเดียวกับตอนที่อ่านมา
- **B2** (เดิม #2, High) — `ui/app.py:116`, `ui/main_window.py:189-192`, `ui/screens/settings.py:66-75, 273-302`, `services/importer.py:113`, `config.py:228-229`: sync ต้องใช้ date window **และ Excel path** ล่าสุดหลังเปลี่ยนใน Settings (ใช้กลไกเดียวกันทั้งสองค่า)
- **B3** (เดิม #3, High) — `services/sync.py:460-463` (`watch`), `ui/main_window.py` `SyncThread.run` และ `_status` :256: watch loop ต้องไม่ตายเมื่อ log ลง DB ไม่ได้ มี exponential backoff และ chip ต้องแสดงสถานะ error เมื่อ sync thread ตายด้วย exception
- **B4** (เดิม #5, Medium) — `services/sync.py:268-274`, `excel/writer.py:312-313`, `:620-624`: `HeaderError` ระหว่างเขียนต้องเข้าสถานะ retry แบบเดียวกับ `WriterError` และสร้าง workbook backup ได้ไม่เกินหนึ่งไฟล์ต่อ error episode
- **B5** (เดิม #4, High) — `ui/main_window.py:51-53`, `ui/app.py:121-136, 150-153, 176-180`, `excel/writer.py` `com_call`: ห้ามทิ้ง reference ของ QThread ที่ยังทำงานอยู่ ปิดแอปต้องแสดงหน้าต่าง "กรุณารอ" และหลัง N วินาทีมีปุ่ม "Force close" และส่งสัญญาณหยุดเข้าไปใน `com_call` เพื่อให้การรอ Excel busy จบได้เร็ว

## นอกขอบเขต (Out of scope / Non-goals)
- ข้อค้นพบอื่นจาก bug hunt (#6–#14)
- การออกแบบ locking ใหม่ของ `BEGIN IMMEDIATE` (รวมถึงประเด็นที่ `BEGIN IMMEDIATE` lock `app.db` ที่ ATTACH ไว้ด้วย) ยกเว้นกรณีที่จำเป็นต้องแก้เล็กน้อยเพื่อให้ AC ของ B3 ผ่านจริง
- พฤติกรรมของ OneDrive / AutoSave
- การเปลี่ยนนโยบาย "backup ไม่ถูกลบ" (ยังคงไม่ลบ backup)
- การ restart sync thread อัตโนมัติเมื่อ thread ตาย (ผู้ใช้ยืนยันว่านอกขอบเขต B3 แค่แสดงสถานะ error)
- `SyncService.init_ids` (`sync.py:344-348`) ที่จับแค่ `WriterError` ทำให้ `HeaderError` หลุดออกไปได้ (ผู้ใช้ยืนยันว่านอกขอบเขต)
- การเปลี่ยนไฟล์ Excel ของโปรเจกต์ในโหมด multi-project ซึ่งทำผ่าน Relink (`relink_flow` -> `open_project`) อยู่แล้ว ไม่ได้ผ่านช่อง Excel path ใน Settings
- ฟีเจอร์ UI ใหม่นอกเหนือจากสถานะ chip (B3) และหน้าต่าง "กรุณารอ" / ปุ่ม "Force close" ตอนปิดแอป (B5)

## User stories
- **US-1 (B1)** ในฐานะผู้ใช้ที่แก้ข้อมูลในแอป ฉันต้องการให้ค่าล่าสุดที่ฉันแก้ถูกเขียนลง Excel เสมอ แม้ฉันจะแก้ระหว่างที่ sync กำลังเขียนค่าก่อนหน้า เพื่อให้ Excel ที่ทีมใช้ร่วมกันตรงกับแอป
- **US-2 (B2)** ในฐานะ PMO ฉันต้องการให้การเปลี่ยนช่วงวันที่หรือ Excel path ใน Settings มีผลกับ sync ทันที เพื่อไม่ให้วันที่ถูกล้าง ไม่มีประวัติ "Edited in Excel" ที่ไม่ได้เกิดขึ้นจริง และ sync ทำงานกับไฟล์ที่ฉันเลือก
- **US-3 (B3)** ในฐานะผู้ใช้ ฉันต้องการให้ sync ทำงานต่อได้เมื่อ DB ถูก lock ชั่วคราว และถ้า sync หยุดทำงานจริง ฉันต้องเห็นบนจอเป็นสถานะ error พร้อมเหตุผล เพื่อไม่เข้าใจผิดว่าข้อมูลถึง Excel แล้ว
- **US-4 (B4)** ในฐานะผู้ดูแลระบบ ฉันต้องการให้ปัญหาชีต/หัวคอลัมน์ใน Excel ถูกจัดการเป็น retry ตามรอบปกติ และไม่สร้างไฟล์ backup ซ้ำ เพื่อไม่ให้ดิสก์เต็มและได้รับการแจ้งเตือนเมื่อแก้ไม่ได้
- **US-5 (B5)** ในฐานะผู้ใช้ ฉันต้องการสลับโปรเจกต์หรือปิดแอปได้โดยแอปไม่ crash ไม่ทิ้ง Excel ที่ซ่อนอยู่ค้างถือไฟล์ และถ้าต้องรอนาน ฉันเลือกบังคับปิดได้โดยรู้ความเสี่ยง

## Acceptance criteria

ทุกข้อทดสอบด้วย pytest ได้ โดยใช้ fake opener / fake COM / fake repo หรือ connection ปลอม (ไม่ต้องใช้ Excel จริง) สำหรับ test ฝั่ง Qt ใช้ `QApplication` แบบ offscreen ตามแนวของ `tests/test_ui.py`

### B1 — version guard ของ pending_writeback

- **AC-1** (แก้ระหว่างเขียน -> ไม่ fail)
  Given รายการ field X มี pending ค่า A ถึงคิว (`_due`) และ fake opener ที่ระหว่าง `write_workbook` เรียก `repo.upsert_pending` ของ X เป็นค่า B **โดยใช้ `created_at` เท่ากับของเดิม** (จำลองการแก้ภายในวินาทีเดียวกัน)
  When `flush(force=True)` ทำงานจนจบ (รวม read-back)
  Then แถว pending ของ X มี `value` = B, `status='pending'`, `attempts=0`, `last_error` เป็น NULL และ **ไม่ใช่** `'failed'` และไม่มีข้อความ "Written, but Excel did not read back the same value" ของ X

- **AC-2** (ค่าใหม่ถูกเขียนในรอบถัดไป)
  Given สถานะหลัง AC-1
  When เรียก `flush(force=True)` อีกครั้ง
  Then fake workbook ได้รับค่า B ที่ cell ของ X และหลัง read-back แถว pending ของ X ถูกล้างออกจากคิว (ไม่อยู่ใน `all_pending()`)

- **AC-3** (`_retry` ไม่ทับแถวที่ถูกแก้)
  Given X มี pending ค่า A ถึงคิว และ fake opener แก้ X เป็นค่า B (upsert) แล้ว raise `WorkbookBusy`
  When `flush(force=True)`
  Then ผลเป็น `retry`, แถวของ X มี `value` = B, `attempts=0`, `next_attempt_at` NULL, `first_failed_at` NULL, `last_error` NULL และอยู่ใน `_due()` ทันที ส่วนรายการอื่นในรอบเดียวกันที่ไม่ถูกแก้ได้ `attempts=1` และ `next_attempt_at` ตามตาราง `retry_minutes` ตามเดิม

- **AC-4** (`_fail` ไม่ทับแถวที่ถูกแก้)
  Given X ถึงคิว และระหว่างรอบ flush X ถูก upsert เป็นค่าใหม่ก่อนที่ sync จะเรียก `_fail` กับ X (ทั้งกรณี fail ก่อนเขียน เช่น "Item is no longer in the Excel file" และกรณี `report.failed` จาก writer)
  When `flush(force=True)`
  Then แถวของ X ยังเป็น `status='pending'` ค่าใหม่ ไม่ถูกตั้งเป็น `'failed'`

- **AC-5** (repository level)
  Given แถว pending ที่อ่านมาเป็น snapshot S แล้วถูก `upsert_pending` เปลี่ยนค่า (created_at เดิมหรือใหม่ก็ได้)
  When เรียก `set_pending_state(..., <version guard จาก S>)`
  Then แถวไม่ถูกเปลี่ยนเลย (ทุกคอลัมน์เหมือนหลัง upsert) และถ้าแถวไม่ถูกแก้ระหว่างนั้น การเรียกเดียวกันอัปเดตได้ตามปกติ

- **AC-6** (ทุกจุดใน sync ส่ง guard)
  Given repo spy ที่บันทึกทุกการเรียก `set_pending_state`
  When รัน flush ครบทุกเส้นทาง (written + read-back fail, `_fail`, `_retry`, reset 'writing' ที่ค้าง, `retry_now`)
  Then ทุกการเรียกจาก `SyncService` ส่ง version guard มาด้วย

- **AC-7** (ไม่ถดถอย)
  Given แถวที่ไม่ถูกแก้ระหว่างเขียน
  When flush
  Then พฤติกรรมเหมือนเดิมทุกกรณี (written -> ล้างจากคิว, read-back ไม่ตรง -> failed, retry ตามตาราง) และ test เดิมทั้งหมด (548) ยังผ่าน

### B2 — date window และ Excel path ล่าสุดหลังเปลี่ยนใน Settings

- **AC-8** (sync ใช้ window ใหม่)
  Given โปรเจกต์ที่ window เดิม W1 และ sync (SyncThread / SyncService) ถูกสร้างก่อนเปลี่ยน, รายการที่มีวันที่อยู่ใน W2 แต่อยู่นอก W1 ถูก import ด้วย W2 แล้ว
  When เปลี่ยน window เป็น W2 ผ่าน Settings (`projects.set_window` + บันทึก) แล้ว sync อ่าน Excel อีกครั้ง (`pull(force=True)` หรือรอบถัดไปของ thread)
  Then วันที่ของรายการนั้นยังอยู่ (ไม่ถูกตั้งเป็น NULL) และไม่มี audit ใหม่ที่ `editor` ขึ้นต้นด้วย "Edited in Excel" สำหรับ field วันที่เหล่านั้น

- **AC-9** (window ที่ sync ใช้ = ค่าใน app.db)
  Given window ของโปรเจกต์ใน `app.db` ถูกเปลี่ยนเป็น W2
  When sync ทำ import รอบถัดไป
  Then `DateWindow` ที่ `run_import` ได้รับ คือ W2 (ตรวจด้วย spy / monkeypatch ของ `run_import` หรือ `DateWindow`)

- **AC-10** (Excel path ใหม่: อ่านและเขียนไฟล์ใหม่)
  Given โหมด single-file (ไม่มี project, ช่อง Excel path แสดงใน Settings) sync ทำงานกับไฟล์ P1 และมีไฟล์ P2 อยู่จริง
  When เปลี่ยน Excel path เป็น P2 ใน Settings แล้วบันทึก จากนั้น sync ทำ `pull` และ `flush` รอบถัดไป
  Then `run_import` อ่านจาก P2, fake opener ถูกเรียกด้วย P2 (ไม่ใช่ P1), backup ถูกสร้างจาก P2 และไฟล์ P1 ไม่ถูกเปิดเขียนหรือแก้ไขหลังการเปลี่ยน

- **AC-11** (Excel path ใหม่: watcher ดูไฟล์ใหม่)
  Given สถานการณ์เดียวกับ AC-10 และ Observer ของ watchdog ถูกแทนด้วย fake ที่บันทึก path ที่ `schedule`
  When เปลี่ยน path เป็น P2 แล้วบันทึก
  Then การ watch ที่ทำงานอยู่หลังการเปลี่ยน schedule โฟลเดอร์ของ P2 และ event ของไฟล์ P2 ทำให้เกิด pull ส่วน event ของ P1 ไม่ทำให้เกิด pull

- **AC-12** (ไม่เปลี่ยนค่า -> ไม่มีผลข้างเคียง)
  Given บันทึก Settings โดยไม่เปลี่ยนทั้ง window และ Excel path
  When `_save`
  Then ไม่มีการ restart sync thread เพิ่ม (ถ้าเลือกแนวทาง restart) และพฤติกรรมเดิมของ Settings (PMO email ล่าสุดตาม T5, dry run ฯลฯ) ยังทำงานเหมือนเดิม

### B3 — watch loop ไม่ตายเงียบ, backoff และ chip สะท้อนสถานะจริง

- **AC-13** (log ลง DB ไม่ได้ -> loop ยังอยู่)
  Given `watch()` ที่ใช้ Observer ปลอม (monkeypatch watchdog), `pull`/`flush` raise และ `repo.log_sync` raise `sqlite3.OperationalError("database is locked")`
  When loop ทำงานอย่างน้อย 3 รอบ แล้ว set `stop`
  Then `watch()` ไม่ raise, ทำงานครบทุกรอบจนถูก stop และข้อผิดพลาดถูกบันทึกลง file log (`logging`, ตรวจด้วย `caplog`) อย่างน้อยรอบละครั้ง

- **AC-14** (`on_event` raise -> loop ยังอยู่)
  Given `on_event` ที่ raise ทุกครั้ง
  When loop ทำงานหลายรอบ
  Then `watch()` ไม่ raise จนกว่าจะ stop

- **AC-15** (exponential backoff สูงสุด 60 วินาที)
  Given `stop` ปลอมที่บันทึกค่า timeout ที่ส่งให้ `stop.wait` และรอบของ loop ที่ error ติดกัน 8 รอบ
  When loop ทำงาน
  Then ค่าที่ส่งให้ `stop.wait` หลัง error รอบที่ 1–8 ตามลำดับคือ 1, 2, 4, 8, 16, 32, 60, 60 วินาที (สูตร `min(2 ** (k - 1), 60)` เมื่อ k = จำนวน error ติดกัน) เพดาน 60 วินาทีเป็นค่าคงที่ที่ตั้งชื่อไว้ในโค้ด

- **AC-16** (reset backoff หลังรอบที่สำเร็จ)
  Given error ติดกัน 4 รอบ (รอ 8 วินาที) แล้วตามด้วยรอบที่สำเร็จ แล้ว error อีกครั้ง
  When loop ทำงาน
  Then หลังรอบที่สำเร็จรอ 1 วินาที (ค่าปกติ) และ error ครั้งถัดไปเริ่มนับใหม่ที่ 1 วินาที ส่วน `stop.set()` ยังทำให้ loop จบได้ภายในระยะรอปัจจุบัน

- **AC-17** (thread ตายด้วย exception -> chip error สีแดง + tooltip เหตุผล)
  Given `SyncThread` ที่ `run()` จบด้วย exception ข้อความ M (เช่น `SqliteRepository` หรือ `watch` raise)
  When `MainWindow._status()` ทำงานหลัง thread จบ
  Then `sync_chip` แสดงข้อความ `t("sync.error")` ด้วยสีของสถานะ error (`color:#A3260F; background:#FDE7E3`) และ tooltip มีข้อความ M, ข้อผิดพลาดถูกบันทึกลง file log และไม่มีการเริ่ม SyncThread ใหม่อัตโนมัติ

- **AC-18** (thread จบโดยไม่ได้สั่งหยุด)
  Given `SyncThread` ที่ไม่ได้ถูก `stop()` แต่ไม่ทำงานแล้ว (`isRunning()` เป็น False หลังเคย start) โดยไม่มี exception
  When `_status()`
  Then chip ไม่แสดง `sync.synced`

- **AC-19** (ไม่ถดถอย)
  Given sync thread ทำงานปกติ หรือ `sync_thread is None`
  When `_status()`
  Then chip แสดงเหมือนเดิม (synced / pending / paused / blocked / error / off)

### B4 — HeaderError ระหว่างเขียน และ backup หนึ่งไฟล์ต่อ episode

นิยาม **error episode**: ช่วงของการลองเขียนที่ล้มเหลวติดกัน (ทุกสาเหตุที่เข้า `_retry`) นับจากครั้งแรกที่ล้มเหลวจนถึงการเขียนที่สำเร็จครั้งถัดไป การเขียนที่สำเร็จเป็นจุดจบของ episode

- **AC-20** (ชีตหาย -> retry)
  Given fake opener ที่ `sheet(name)` raise `HeaderError` (จำลอง com_error จาก `Worksheets(name)`) และมีรายการถึงคิว
  When `flush(force=True)`
  Then ไม่มี exception หลุดออกจาก `flush`, ผลเป็น `retry`, ทุกรายการในคิวได้ `attempts=1`, `next_attempt_at` ตามตาราง `retry_minutes` และ `last_error` มีข้อความของ `HeaderError`

- **AC-21** (หัวคอลัมน์หาย -> retry)
  Given workbook ที่หัวคอลัมน์ที่ต้องใช้หายไป (`map_headers` raise `HeaderError` ระหว่าง `apply_writes`)
  When `flush(force=True)`
  Then ผลเหมือน AC-20

- **AC-22** (backup ไม่เกิน 1 ไฟล์ต่อ episode)
  Given สถานการณ์ AC-20 ที่ HeaderError เกิดทุกครั้ง
  When เรียก `flush()` ซ้ำหลายครั้ง ทั้งก่อนถึง `next_attempt_at` และเลื่อนนาฬิกาปลอมผ่านทุกช่วงของตาราง `retry_minutes` จนเกินตาราง
  Then ตลอด episode จำนวนไฟล์ใน `backup_dir` และจำนวนแถว backup ใน DB เพิ่มขึ้นไม่เกิน 1

- **AC-23** (การเขียนที่สำเร็จปิด episode โดยไม่สร้าง backup ที่สอง)
  Given episode ที่มี backup แล้ว 1 ไฟล์ จากนั้นสาเหตุของ error หายไป
  When flush ครั้งถัดไปเขียนสำเร็จ
  Then จำนวน backup ของ episode นั้นยังไม่เกิน 1 และ flush ที่ล้มเหลวครั้งต่อไปหลังจากนั้นเป็น episode ใหม่ ซึ่งสร้าง backup ได้อีก 1 ไฟล์

- **AC-24** (ใช้กับทุกสาเหตุที่เข้า retry)
  Given episode ที่ล้มเหลวด้วย `WorkbookLocked` / `WorkbookBusy` / `RowsMoved` / `WriterError` ซ้ำหลายครั้ง
  When flush ซ้ำผ่านตาราง retry
  Then backup ไม่เกิน 1 ไฟล์ต่อ episode เช่นเดียวกับ AC-22

- **AC-25** (ครบรอบ retry -> error + alert ครั้งเดียว)
  Given HeaderError เกิดทุกครั้งจนครบตาราง `retry_minutes`
  When flush ครั้งที่เกินตาราง
  Then ผลเป็น `error` และส่ง alert ถึง PMO หนึ่งครั้งต่อ episode เหมือน error ประเภทอื่น

- **AC-26** (ไม่เคยบันทึก workbook โดยไม่มี backup)
  Given การเขียนที่สำเร็จ ทั้งแบบไม่มี error มาก่อน และแบบที่ปิด episode
  When flush
  Then มี backup ของ episode/รอบนั้นที่ถูกสร้าง **ก่อน** `book.save()` ถูกเรียกเสมอ (ตรวจด้วยลำดับการเรียกใน fake opener) และการเขียนสำเร็จที่ไม่มี error มาก่อนยังสร้าง backup 1 ไฟล์เหมือนเดิม

### B5 — QThread ไม่ถูกทำลายขณะยังทำงาน และการปิดแอป

- **AC-27** (stop timeout -> เก็บ reference ไว้)
  Given `SyncThread` (หรือ thread ปลอมที่ block จนกว่าจะปล่อย) ที่ยังไม่จบหลัง `stop()` ครบ timeout
  When `AppController._close_current()` / `open_project()` (สลับโปรเจกต์) / `restart_sync()`
  Then object ของ thread นั้นยังถูกอ้างอิงอยู่ใน controller (เช่น retiring list) ตราบที่ `isRunning()` เป็น True และไม่มี qFatal / crash

- **AC-28** (thread จบ -> ปล่อย reference)
  Given thread ใน retiring list
  When thread นั้นจบ (`finished`)
  Then ถูกนำออกจาก retiring list และถูก `deleteLater()` (signal `finished` ต่อกับ `deleteLater`)

- **AC-29** (ปิดแอปตอนไม่มี sync ทำงาน -> ปิดทันที)
  Given ไม่มี thread ที่ยังทำงานอยู่ (ทั้งตัวปัจจุบันและใน retiring list)
  When `AppController.quit()` หรือปิด `MainWindow`
  Then `QApplication.quit()` ถูกเรียกทันทีและไม่มีหน้าต่าง "กรุณารอ"

- **AC-30** (ปิดแอปตอน sync ยังทำงาน -> หน้าต่าง "กรุณารอ")
  Given มี thread ทำงานอยู่ (ตัวปัจจุบันหรือใน retiring list)
  When `AppController.quit()` หรือปิด `MainWindow`
  Then แสดงหน้าต่าง "กรุณารอ" ที่บอกว่ากำลังรอ sync จบ, `QApplication.quit()` **ยังไม่ถูกเรียก** จนกว่าทุก thread จะจบ, เมื่อทุก thread จบหน้าต่างปิดเองและ `QApplication.quit()` ถูกเรียก และระหว่างรอ UI ยังตอบสนอง (ไม่ใช้ `wait()` แบบ block บน UI thread)

- **AC-31** (หลัง N วินาทีมีปุ่ม "Force close" พร้อมคำเตือน)
  Given หน้าต่าง "กรุณารอ" ของ AC-30 และค่าคงที่ที่ตั้งชื่อไว้ในโค้ด (เช่น `QUIT_FORCE_AFTER_S`) มีค่าเริ่มต้น 15 วินาที
  When ยังไม่ครบ N วินาที / ครบ N วินาทีแล้ว (ใน test ใช้ timer ปลอมหรือ monkeypatch ค่าคงที่ให้สั้น)
  Then ก่อนครบ N ปุ่ม "Force close" ไม่แสดง และหลังครบ N ปุ่มแสดงพร้อมข้อความเตือนว่า EXCEL.EXE ที่ซ่อนอยู่อาจยังถือไฟล์ Excel ค้างไว้ ข้อความทั้งสองมีทั้งภาษาไทยและอังกฤษใน i18n

- **AC-32** (กด "Force close")
  Given ปุ่ม "Force close" แสดงอยู่และ thread ยังไม่จบ
  When ผู้ใช้กดปุ่ม
  Then แอปปิดทันทีโดยไม่รอ thread (ผ่านฟังก์ชันออกจากโปรแกรมที่ monkeypatch ได้ใน test) และมี warning ลง file log ว่าบังคับปิดขณะ sync ยังทำงาน

- **AC-33** (stop event ส่งเข้า `com_call`)
  Given `com_call` ที่ fn raise busy error (`RPC_E_CALL_REJECTED`) ทุกครั้ง และ stop event ของ thread นั้นถูก set แล้ว (หรือถูก set ระหว่างรอ)
  When `com_call` ทำงาน
  Then จบด้วย exception ภายในการรอไม่เกิน 1 ช่วง delay (ตรวจด้วย `_sleep` ปลอมว่าไม่ถูกเรียกต่อหลัง stop) แทนการรอจนหมด budget 30 วินาที

- **AC-34** (ยกเลิกกลางคันแล้ว workbook ถูกปิด และคิวไม่หาย)
  Given flush ถูกยกเลิกกลางคันด้วยกลไกของ AC-33
  When `write_workbook` จบ
  Then `book.close()` ถูกเรียก และรายการในคิวยังอยู่ (สถานะ pending หรือ retry ไม่ใช่ failed) เพื่อเขียนต่อเมื่อเปิดโปรเจกต์อีกครั้ง

- **AC-35** (ไม่มีจุดที่ทิ้ง reference)
  Given ทุกเส้นทางที่หยุด sync (`_close_current`, `restart_sync`, `quit`, `MainWindow.closeEvent` กรณี `_owns_sync`, และการ restart ของ B2)
  When thread ยังไม่จบหลัง stop
  Then ไม่มีเส้นทางใดตั้ง reference เป็น `None` โดยไม่ย้ายไปเก็บที่อื่นก่อน

### ทั่วไป
- **AC-36** `pytest -q` ผ่านทั้งหมด: test เดิม 548 ข้อ และ test ใหม่ของ AC-1 ถึง AC-35

## ข้อจำกัด / สมมติฐาน
- Python 3.x, PySide6, SQLite, Excel COM (pywin32) โค้ดตาม PEP 8 และใช้ type hints
- test ห้ามใช้ Excel จริง ใช้ fake opener / fake COM / fake repo ได้ (Observer ของ watchdog ต้อง monkeypatch เพราะยังไม่มี test ของ `watch()` เลย)
- `created_at` ใน `pending_writeback` มีความละเอียดระดับวินาที (`timespec="seconds"`) ดังนั้น version guard ของ B1 **ห้ามพึ่ง `created_at` อย่างเดียว** ต้องจับการแก้ภายในวินาทีเดียวกันได้ (เช่น เทียบ `created_at` + `value` + `excel_value` หรือใช้ตัวนับเวอร์ชัน) กรณีผู้ใช้แก้กลับเป็นค่าเดิมเป๊ะภายในวินาทีเดียวกัน ถือว่าไม่เป็นปัญหา เพราะค่าที่เขียนคือค่าที่ต้องการอยู่แล้ว
- B2 / Excel path: ช่อง Excel path ใน Settings มีเฉพาะโหมด single-file (`settings.py:67`) และ override ถูกใช้เฉพาะเมื่อ `project_id is None` (`config.py:228`) ในโหมดนี้ sync thread อาจเป็นของ `MainWindow` (`_owns_sync`) ไม่ใช่ของ `AppController` กลไกที่เลือกต้องครอบทั้งสองแบบ ข้อความ `set.restart` ("มีผลหลังเปิดแอปใหม่") จะไม่ตรงกับพฤติกรรมใหม่ ให้ปรับหรือเอาออกได้เพราะเป็นส่วนหนึ่งของการแก้นี้
- การเลือกแนวทางทางเทคนิค (เช่น B2 ใช้ restart sync หรือให้ `SyncService` สร้าง cfg/path ใหม่ทุกรอบ, B4 เก็บสถานะ episode ไว้ที่ใด) เป็นหน้าที่ของ architect ตราบที่ผ่าน AC
- ถ้า B2 เลือกแนวทาง restart sync จะเรียก `stop()` ซึ่งมีความเสี่ยงแบบ B5 จนกว่า B5 จะเสร็จ ยอมรับได้ตามลำดับที่กำหนด แต่หลัง B5 เสร็จ เส้นทางนี้ต้องผ่าน AC-27 และ AC-35 ด้วย
- B4: กติกา "backup ไม่เกิน 1 ไฟล์ต่อ episode" ใช้กับทุกสาเหตุที่เข้า retry ไม่ใช่เฉพาะ `HeaderError` ถ้ามี test เดิมที่นับจำนวน backup แล้วขัดกับกติกานี้ ห้ามแก้ test เงียบ ๆ ให้รายงานหัวหน้า/ผู้ใช้ก่อน
- B5: "Force close" เป็นทางเลือกของผู้ใช้ที่ยอมรับความเสี่ยงเอง (EXCEL.EXE ค้าง) ส่วนค่า N = 15 วินาทีปรับได้ที่ค่าคงที่ในโค้ดเท่านั้น (ไม่ต้องมีช่องใน Settings)
- นโยบายเดิมยังคงอยู่: backup ไม่ถูกลบ, ตาราง retry 1/2/5/10/15/30 นาที, email แจ้ง PMO หนึ่งครั้งต่อ episode, lock file กันการเขียนพร้อมกัน

## การตัดสินใจของผู้ใช้ (2026-11-24)
- Q1 (B3) backoff แบบ exponential 1 -> 2 -> 4 ... วินาที สูงสุด 60 วินาที reset หลังรอบที่สำเร็จ -> AC-15, AC-16
- Q2 (B3) thread ตายด้วย exception -> chip สถานะ error สีแดงพร้อม tooltip บอกเหตุผล ไม่ restart อัตโนมัติ -> AC-17
- Q3 (B4) workbook backup ไม่เกิน 1 ไฟล์ต่อ error episode -> AC-22, AC-23, AC-24, AC-26
- Q4 (B5) ปิดแอปขณะ sync ทำงาน -> หน้าต่าง "กรุณารอ" หลัง N = 15 วินาที (ค่าคงที่) มีปุ่ม "Force close" พร้อมคำเตือนเรื่อง EXCEL.EXE -> AC-29 ถึง AC-32
- Q5 (B2) เปลี่ยน Excel path ใน Settings ต้องมีผลกับ sync ด้วย (อยู่ในขอบเขต) -> AC-10, AC-11, AC-12
- Q6 (B4) `init_ids` กับ `HeaderError` -> นอกขอบเขต

## คำถามที่ยังไม่มีคำตอบ
- ไม่มี

---

# 2026-11-25 — หน้าตาใหม่ของเมล Weekly Report และชื่อสถานะ "มีความเสี่ยงล่าช้า"

> แบบที่ผู้ใช้อนุมัติ: Design canvas "Planning Progress – Weekly UAT Report Email" (https://claude.ai/artifact/XxxvEeuHWCk4xKhGRCF4EF) กรอบ `Main` = เมลจริงจากข้อมูล DEMO UAT วันที่ 25-Nov-2026, กรอบ `Appendix` = หน้าตาของแต่ละส่วนตอนที่มีรายการ (ข้อมูลสมมติ)

## เป้าหมาย
เมล Weekly Report ปัจจุบัน (`services/weekly_report.py` `render`) มีตัวหนังสือมาก ทุกจุดเขียนสองภาษา "ไทย / English" ข้อมูลบางอย่างแสดงซ้ำ และมีเรื่องภายในของ PMO ปนอยู่ ทำให้ทีมและผู้บริหารอ่านยาก รอบนี้เปลี่ยนหน้าตาเมลให้สะอาดตา อ่านสถานะได้ในไม่กี่วินาที โดยข้อมูลยังครบ และเปลี่ยนชื่อสถานะ Amber ภาษาไทยเป็น "มีความเสี่ยงล่าช้า" ทั้งในเมลและหน้าจอโปรแกรม

ผู้ได้ประโยชน์: ผู้บริหารและทีมที่ได้รับเมลทุกสัปดาห์ และ PMO ที่ต้องส่งรายงาน

## ขอบเขต (In scope)

### ส่วนที่ 1: ชื่อสถานะ
- **R1** `LABEL_TH[AMBER]` ใน `services/traffic_light.py` เปลี่ยนจาก "มีความเสี่ยง" เป็น "มีความเสี่ยงล่าช้า" และ `LABEL_EN[AMBER]` เปลี่ยนจาก "At risk" เป็น "At risk of delay" ทุกที่ที่อ่านค่านี้จะเปลี่ยนตาม (หน้าจอ `ui/i18n.py` `tag.Amber`, `tp.amber`, เมลแจ้งเตือน `alerts/templates.py` `COLOR_TH`, เมล Weekly Report)
- **R1a Excel** คอลัมน์ Traffic Light ที่ `daily_job._queue_traffic_lights` เขียนจะได้ "At risk of delay" แทน "At risk" ไฟล์ที่ยังมี "At risk" ถูกเขียนทับอัตโนมัติด้วยกลไกเดิม (เทียบค่ากับ `label()`) ใน `resources/project_template.json` แก้ 3 จุด: dropdown `"On track,At risk of delay,Delayed"`, conditional format ที่เทียบ `"At risk"` เป็น `"At risk of delay"` (สีเดิม) และค่าตัวอย่างในแถวข้อมูล
- **R1b เอกสาร** README.md ส่วนที่อธิบายสถานะ (คำอังกฤษและไทย) และ docstring/คอมเมนต์ที่อ้าง "At risk" / "มีความเสี่ยง"

### ส่วนที่ 2: เมล Weekly Report (`services/weekly_report.py`, `services/weekly_data.py`)
- **R2 ภาษา** เนื้อหาเป็นภาษาไทย ภาษาอังกฤษมีเฉพาะชื่อหัวข้อ เป็นบรรทัดเล็กสีเทาใต้หัวข้อไทย ตามตารางนี้
  | หัวข้อไทย | บรรทัดอังกฤษ |
  |---|---|
  | รายงาน UAT ประจำสัปดาห์ · {project} | Weekly UAT Report |
  | สถานะแอป {n} แอป | App status (อยู่บรรทัดเดียวกัน คั่นด้วย ·) |
  | แอปที่ต้องติดตาม | Apps to watch |
  | งาน UAT Plan ที่เลยกำหนด | Overdue UAT Plan tasks |
  | Defect | (ไม่มี เพราะเป็นคำอังกฤษอยู่แล้ว) |
  | เลยวันเป้าหมาย / ไม่มีวันเป้าหมาย | Past fix target / No fix target (อยู่บรรทัดเดียวกัน คั่นด้วย ·) |
  | ครบกำหนดใน 1 สัปดาห์ | Due within 1 week |
  | การเลื่อนวัน {from} ถึง {to} | Reschedules |
  | อื่น ๆ | Other |
  คำที่ทีมใช้อยู่แล้วคงเป็นภาษาอังกฤษได้: Pass, Fail, Defect, Total TC, Planned Finish, UAT Plan, Test Progress, ชื่อแอป, ชื่องาน, สถานะ Defect
- **R3 หัวเมล** `[Planning Progress] รายงาน UAT ประจำสัปดาห์ {project} · {as_of}`
- **R4 ส่วนหัว** "Planning Progress" (เล็ก), ชื่อรายงาน, บรรทัด "Weekly UAT Report", "ข้อมูล ณ {as_of}"
- **R5 บรรทัดสรุป** ในกล่องพื้นเทาอ่อน: "ล่าช้า {n} แอป · มีความเสี่ยงล่าช้า {n} แอป · ข้อมูลไม่ครบ {n} แอป" (สีแดง / ส้ม / เทา) คำนวณจากข้อมูลเดียวกับแถบสถานะ ไม่ใช้ AI
- **R6 ตัวเลขหลัก** 4 ช่อง: ทดสอบแล้ว {pct} พร้อม "{executed} จาก {total} เคส" และแถบ %, Pass, Fail, Defect ที่เปิดอยู่
- **R7 แถบสถานะแอป** แถบสีแบ่งตามสัดส่วน เสร็จ / ล่าช้า / มีความเสี่ยงล่าช้า / ตามแผน / ข้อมูลไม่ครบ พร้อมคำอธิบายสีและจำนวน สถานะที่เป็น 0 ไม่มีช่องในแถบ
- **R8 แอปที่ต้องติดตาม** แสดงแอปสถานะล่าช้า, มีความเสี่ยงล่าช้า และ**ข้อมูลไม่ครบ (ใหม่)** เรียงตามลำดับนั้น แต่ละแอป: จุดสี, ชื่อแอป, ป้ายสถานะ, เหตุผลภาษาไทย 1 บรรทัด, ด้านขวา "{executed}/{total}" (ข้อมูลไม่ครบแสดง "-") และหมายเหตุท้ายส่วน "ตัวเลขด้านขวาคือเคสที่ทดสอบแล้ว / ทั้งหมด" เหตุผลภาษาไทยตามกรณี:
  - เลยวันจบตามแผน: "เลยวันจบตามแผน ({planned_finish}) มา {days_late} วัน และยังไม่เสร็จ"
  - Fail เกินเกณฑ์: "Fail {pct_fail} เกินเกณฑ์ {threshold}"
  - ช้ากว่าเวลา: "ทดสอบได้ {pct_exec} แต่เวลาผ่านไปแล้ว {pct_time}"
  - มี Fail แต่ไม่เกินเกณฑ์: "มี Fail {n} เคส ({pct_fail}) แต่ยังไม่เกินเกณฑ์ {threshold}"
  - ช้ากว่าเวลาและมี Fail ไม่เกินเกณฑ์พร้อมกัน: "ทดสอบได้ {pct_exec} แต่เวลาผ่านไปแล้ว {pct_time} · มี Fail {n} เคส ({pct_fail}) แต่ยังไม่เกินเกณฑ์ {threshold}"
  - ข้อมูลไม่ครบ: "ยังไม่ได้กรอก {fields}" โดย fields คือช่องที่ขาด (Total TC, Planned Start, Planned Finish) คั่นด้วย ", "
- **R9 งาน UAT Plan ที่เลยกำหนด** แสดงเฉพาะงานจาก UAT Plan (งานใน Test Progress ที่เลยกำหนดแสดงใน R8 แล้ว) แต่ละแถว: ID, ชื่องาน, "กำหนด {date}", ด้านขวา "เลยมา {n} วัน" สีแดง ไม่มีรายการแสดง "ไม่มี"
- **R10 Defect** บรรทัดเดียว "เปิดอยู่ {n} · เลยวันเป้าหมาย {n} · ไม่มีวันเป้าหมาย {n}" ถ้ามีรายการ แสดงรายการย่อยใต้หัวเล็ก "เลยวันเป้าหมาย" และ "ไม่มีวันเป้าหมาย" แต่ละแถว: #ID, System, รายละเอียด, บรรทัดเล็ก "{status} · เป้าหมาย {date}" และด้านขวา "เลยมา {n} วัน" (เฉพาะเลยวันเป้าหมาย)
- **R11 ครบกำหนดใน 1 สัปดาห์** ช่วงเปลี่ยนจาก 14 วันเป็น 7 วัน (`days_ahead=7` ของ `build_data`) หัวข้อมีจำนวนงาน จัดกลุ่มตามวันที่ แต่ละงาน: ID, ชื่อ, ป้ายหน้า (UAT Plan / Test Progress) แอปใน Test Progress ที่อยู่ใน R8 มีจุดสีตามสถานะหน้าชื่อ และหมายเหตุท้ายส่วน "จุดสีหน้าชื่อคือแอปที่อยู่ในรายการต้องติดตาม"
- **R12 อื่น ๆ** การเลื่อนวัน {from} ถึง {to} (เช่น "การเลื่อนวัน 18-Nov-2026 ถึง 24-Nov-2026"): ไม่มีแสดง "ไม่มี" ถ้ามี แต่ละแถว: ชื่อ, "{field} {from} → {to}", บรรทัดเล็ก "{date} · {page} · เหตุผล: {reason}" หรือ "แก้ใน Excel ไม่ได้ระบุเหตุผล" และบรรทัด "ข้อขัดแย้งที่ยังเปิด: {n}"
- **R13 ตัดออกจากเมล** ข้อความ "ไม่มีสรุปจาก AI ...", "คิวตรวจของ PMO", "ไม่มีผู้รับใน Report By Email / CC Email ..." และ "ข้ามอีเมลที่ไม่ถูกต้อง ..." ไม่อยู่ในเมลที่ส่ง PMO ได้รับแจ้งแยก: เรื่อง AI และไม่มีผู้รับใช้ `alert_pmo` เดิมในรอบอัตโนมัติ, อีเมลไม่ถูกต้องเพิ่ม `alert_pmo` ใหม่ในรอบอัตโนมัติ (บอกจำนวนและรายการ แบบเดียวกับข้อความเดิม) และในหน้า preview ยังเห็นทั้งสามเรื่อง
- **R14 ท้ายเมล** บรรทัดเดียว "ตัวเลขทั้งหมดคำนวณโดยระบบ Planning Progress ของ PMO ({project}) และส่งอัตโนมัติ" ถ้ามีสรุปจาก AI เพิ่ม "ส่วนสรุปร่างโดย AI (Claude)"
- **R16 สรุปจาก AI** เมื่อ AI ทำงานสำเร็จ วางส่วนสรุปใต้บรรทัดสรุป (R5) ก่อนตัวเลขหลัก (R6) ใช้เฉพาะข้อความภาษาไทยที่ AI เขียน (`summary_th`, `risks` `title_th`/`why_th`, `follow_ups` `action_th`/`why_th`) หัวข้อ "สรุปสำหรับผู้บริหาร (ร่างโดย AI)" / Executive summary, "ความเสี่ยงหลัก" / Top risks, "สิ่งที่ควรติดตาม" / Suggested follow-ups ตามกติกา R2 ข้อความภาษาอังกฤษของ AI ไม่แสดงในเมล เมื่อ AI ไม่สำเร็จ ไม่มีส่วนนี้เลย (R13)
- **R15 HTML สำหรับเมล** ใช้ table + inline style เท่านั้น (ไม่มี flex, grid, `<style>` ภายนอก) เพื่อให้แสดงใน Outlook ได้ ความกว้างเนื้อหา 640px สีตามแบบ: ล่าช้า #DC2626, มีความเสี่ยงล่าช้า #F59E0B (ตัวอักษร #B45309), ตามแผน #16A34A, เสร็จ #2563EB, ข้อมูลไม่ครบ #9CA3AF

## นอกขอบเขต (Out of scope)
- หน้าจอ Dashboard ของโปรแกรม (`services/dashboard.py` `days_ahead=14` ยังเหมือนเดิม) และหน้าจออื่นนอกจากการเปลี่ยนชื่อสถานะตาม R1
- ข้อความ `Light.reason` ภาษาอังกฤษใน `traffic_light.py` ที่บันทึกลงประวัติและ audit ห้ามเปลี่ยน เหตุผลภาษาไทยใน R8 สร้างเฉพาะตอนทำเมล
- `LABEL_EN` ของสถานะอื่น (Delayed / On track / Complete / Missing data) ไม่เปลี่ยน
- การเทียบกับสัปดาห์ก่อน (เช่น % ทดสอบเพิ่มขึ้น) ยังไม่ทำ
- ตาราง/กฎการคำนวณ Traffic Light, ผู้รับเมล, กำหนดการส่ง, เนื้อหาที่ส่งให้ AI (`weekly_ai`)
- เนื้อหาและ prompt ที่ AI ใช้เขียนสรุป

## User stories
- **US-1** ในฐานะผู้บริหาร ฉันต้องการเห็นสถานะรวมของ UAT ในไม่กี่วินาทีที่เปิดเมล เพื่อรู้ว่าโปรเจกต์มีปัญหาหรือไม่
- **US-2** ในฐานะผู้บริหารหรือทีม ฉันต้องการเห็นว่าแอปไหนต้องติดตามและเพราะอะไรเป็นภาษาที่อ่านเข้าใจทันที เพื่อสั่งการหรือแก้ไขได้ตรงจุด
- **US-3** ในฐานะทีม ฉันต้องการรู้งานที่ครบกำหนดในสัปดาห์หน้าและงานไหนมีความเสี่ยงอยู่แล้ว เพื่อวางแผนงานสัปดาห์นี้
- **US-4** ในฐานะ PMO ฉันต้องการให้เรื่องภายใน (AI, คิวตรวจ) ไม่ไปอยู่ในเมลผู้บริหาร แต่ฉันยังได้รับแจ้ง
- **US-5** ในฐานะผู้ใช้โปรแกรม ฉันต้องการให้ชื่อสถานะในหน้าจอกับในเมลตรงกัน คือ "มีความเสี่ยงล่าช้า"

## Acceptance criteria
ทุกข้อทดสอบด้วย pytest โดยเรียก `render` / `build_data` / `generate` กับ repo ทดสอบตามแนวของ `tests/test_weekly.py` และ `tests/test_status_labels.py` ไม่ต้องส่งเมลจริง

### ชื่อสถานะ (R1)
- **AC-1** Given `LABEL_TH` When อ่าน `LABEL_TH[AMBER]` Then ได้ "มีความเสี่ยงล่าช้า" และ `LABEL_EN[AMBER]` ได้ "At risk of delay"
- **AC-2** Given ภาษา UI ไทย When `i18n.t("tag.Amber")` และ `i18n.t("tp.amber")` Then ได้ "มีความเสี่ยงล่าช้า" Given ภาษา UI อังกฤษ Then ได้ "At risk of delay"
- **AC-3a** (Excel) Given แอปสี Amber และค่าใน Excel เป็น "At risk" When `_queue_traffic_lights` Then มีการเขียน "At risk of delay" ลงคอลัมน์ Traffic Light Given ค่าใน Excel เป็น "At risk of delay" อยู่แล้ว Then ไม่เขียนซ้ำ
- **AC-3b** (template) Given `project_template.json` Then dropdown ของ Traffic Light เป็น "On track,At risk of delay,Delayed", conditional format เทียบ "At risk of delay" และไม่มีคำว่า "At risk" ที่ไม่ตามด้วย " of delay" เหลืออยู่
- **AC-3** Given เมลแจ้งเตือนสี Amber When สร้างด้วย `alerts/templates.py` Then มี "มีความเสี่ยงล่าช้า (At risk of delay)"

### เมล Weekly Report
- **AC-4** (หัวเมล) Given project "DEMO UAT" as_of 25-Nov-2026 When `render` Then subject = "[Planning Progress] รายงาน UAT ประจำสัปดาห์ DEMO UAT · 25-Nov-2026"
- **AC-5** (ภาษา) Given ข้อมูลใด ๆ When `render` Then HTML มีบรรทัดอังกฤษตามตาราง R2 และไม่มีข้อความคู่ภาษาเดิม เช่น "Executed", "Overview", "Data as of", "ทดสอบแล้ว / Executed", "Delayed / At risk apps"
- **AC-6** (บรรทัดสรุป) Given ล่าช้า 2, มีความเสี่ยงล่าช้า 3, ข้อมูลไม่ครบ 1 When `render` Then มี "ล่าช้า 2 แอป", "มีความเสี่ยงล่าช้า 3 แอป", "ข้อมูลไม่ครบ 1 แอป"
- **AC-7** (ตัวเลขหลัก) Given executed 260, total 370, pass 255, fail 5, defect เปิด 5 When `render` Then มี "70.3%", "260 จาก 370 เคส", 255, 5 และ "Defect ที่เปิดอยู่"
- **AC-8** (แถบสถานะ) Given จำนวนแต่ละสถานะ When `render` Then คำอธิบายสีมีทั้ง 5 สถานะพร้อมจำนวน และความกว้างของแต่ละช่องในแถบเป็นสัดส่วนของจำนวนแอป สถานะที่เป็น 0 ไม่มีช่อง
- **AC-9** (เหตุผลภาษาไทย) Given แอปแต่ละกรณีของ R8 When `render` Then เหตุผลตรงกับข้อความใน R8 ทุกกรณี (test แยกกรณีละ 1 ข้อ รวม 6 ข้อ)
- **AC-10** (ข้อมูลไม่ครบ) Given แอปที่ไม่มี Total TC When `build_data` และ `render` Then แอปนั้นอยู่ในแอปที่ต้องติดตาม ป้าย "ข้อมูลไม่ครบ" เหตุผล "ยังไม่ได้กรอก Total TC" ด้านขวา "-" และอยู่หลังแอปล่าช้าและมีความเสี่ยงล่าช้า
- **AC-11** (ไม่ซ้ำ) Given แอปใน Test Progress เลย Planned Finish และยังไม่เสร็จ When `render` Then แอปนั้นอยู่ในแอปที่ต้องติดตามพร้อม "มา {n} วัน" และไม่อยู่ในส่วนงาน UAT Plan ที่เลยกำหนด
- **AC-12** (UAT Plan เลยกำหนด) Given งาน UAT Plan ที่เลยกำหนด 1 งาน When `render` Then แสดง ID ชื่อ วันกำหนด และ "เลยมา {n} วัน" Given ไม่มี Then แสดง "ไม่มี"
- **AC-13** (Defect) Given defect เปิด 5 ไม่มีรายการเลยเป้าหมายหรือไม่มีเป้าหมาย When `render` Then มีบรรทัดสรุปและไม่มีรายการย่อย Given มีรายการ Then แสดงรายการใต้หัว "เลยวันเป้าหมาย" / "ไม่มีวันเป้าหมาย" ตาม R10
- **AC-14** (7 วัน) Given งานครบกำหนดวันที่ as_of+7 และ as_of+8 When `build_data` Then มีงาน as_of+7 และไม่มีงาน as_of+8
- **AC-15** (จัดกลุ่มและจุดสี) Given 2 งานวันเดียวกัน และ 1 งานเป็นแอปมีความเสี่ยงล่าช้า When `render` Then วันที่แสดงครั้งเดียวต่อกลุ่ม หัวข้อมีจำนวนงานถูกต้อง และเฉพาะงานของแอปที่อยู่ใน R8 มีจุดสี
- **AC-16** (การเลื่อนวัน) Given ไม่มีการเลื่อนวัน Then แสดง "การเลื่อนวัน {from} ถึง {to}: ไม่มี" (ไม่ใช่ "{from}-{to}") Given มี 1 รายการที่แก้ใน Excel และ 1 รายการที่แก้ในแอปพร้อมเหตุผล Then แสดงตาม R12 ทั้งสองแบบ
- **AC-17** (ตัดเรื่องภายใน) Given ไม่มี Claude API key, คิวตรวจ 1 รายการ, ไม่มีผู้รับ และอีเมลไม่ถูกต้อง 1 รายการ When `generate` Then HTML ที่เก็บไว้ไม่มี "No AI summary", "ไม่มีสรุปจาก AI", "คิวตรวจ" / "review queue", "ไม่มีผู้รับ" และ "ข้ามอีเมลที่ไม่ถูกต้อง"
- **AC-21** (แจ้ง PMO แยก) Given เงื่อนไขเดียวกับ AC-17 When `run_weekly` แบบอัตโนมัติ Then PMO ได้ `alert_pmo` 3 เรื่อง: "Weekly report sent without the AI summary" (เดิม), "Weekly report: no recipients" (เดิม) และเรื่องอีเมลไม่ถูกต้องที่มีจำนวนและรายการ (ใหม่)
- **AC-22** (สรุป AI) Given AI ทำงานสำเร็จ When `render` Then ส่วนสรุปอยู่หลังบรรทัดสรุปและก่อนตัวเลขหลัก มีข้อความ `summary_th`, `title_th`, `action_th` และไม่มี `summary_en`, `title_en`, `action_en` Given AI ไม่สำเร็จ Then ไม่มีหัวข้อ "สรุปสำหรับผู้บริหาร"
- **AC-23** (ท้ายเมล) Given AI สำเร็จ Then ท้ายเมลมี "ส่วนสรุปร่างโดย AI (Claude)" Given ไม่สำเร็จ Then ไม่มี
- **AC-18** (HTML สำหรับเมล) Given HTML จาก `render` Then ไม่มี `display:flex`, `display:grid` และ `<style` และไม่มีเครื่องหมาย — หรือ –
- **AC-19** (ข้อมูลไม่เปลี่ยน) Given ข้อมูลชุดเดียวกัน When เทียบตัวเลขในเมลใหม่กับ `build_data` Then ทุกตัวเลขตรงกัน (ไม่มีการคำนวณใหม่ใน `render`)
- **AC-20** `pytest -q` ผ่านทั้งหมด test ใน `tests/test_status_labels.py` และ `tests/test_weekly.py` ที่ตรวจข้อความเดิม (เช่น "มีความเสี่ยง", "At risk", "แอปที่ล่าช้า / มีความเสี่ยง", "No AI summary") ต้องแก้ให้ตรงกับ spec นี้ ห้ามลบ test ทิ้ง

## ข้อจำกัด / สมมติฐาน
- Python 3.x, PEP 8, type hints, pytest
- `weekly_data.build_data` เป็นที่เดียวที่คำนวณตัวเลข `render` แค่จัดหน้า (หลักการเดิมของ Part 6)
- เปลี่ยนชื่อ key `due_next_2_weeks` เป็นชื่อที่ตรงกับ 7 วันได้ แต่ข้อมูลนี้ถูกส่งให้ AI ด้วย ถ้าเปลี่ยนชื่อ key ต้องตรวจว่า `weekly_ai` ยังทำงานและ test ของ AI ผ่าน
- รายงานเก่าที่เก็บใน `weekly_report` (HTML เดิม) ไม่ต้องแปลงใหม่
- **ผลกับไฟล์ Excel ที่มีอยู่แล้ว** (ผู้ใช้รับทราบ 25-Nov-2026): template ใหม่มีผลเฉพาะโปรเจกต์ที่สร้างใหม่ ไฟล์ Excel เดิมยังมี dropdown และ conditional format ของคำ "At risk" เมื่อโปรแกรมเขียน "At risk of delay" ลงไป cell นั้นจะไม่มีสีส้มตาม conditional format เดิม และไม่ตรงกับ dropdown เดิม (การเขียนผ่าน COM ไม่ถูก validation บล็อก) ผู้ใช้ต้องแก้ dropdown และ conditional format ในไฟล์เดิมเอง การแก้ไฟล์ Excel เดิมอัตโนมัติอยู่นอกขอบเขต
- การเลือกวิธีสร้างเหตุผลภาษาไทย (ใน `weekly_data` หรือ `weekly_report`) เป็นหน้าที่ architect ตราบที่ AC-9 ผ่านและไม่แก้ `Light.reason`

## การตัดสินใจของผู้ใช้ (2026-11-25)
- เมลต้องสะอาดตา ตัวหนังสือน้อย แต่ข้อมูลครบ
- ครบกำหนดแสดง 1 สัปดาห์ -> R11, AC-14
- ทำข้อเสนอ 1-8 (บรรทัดสรุป, ไม่แสดงซ้ำ, เหตุผล Mobile App, ข้อมูลไม่ครบ, ตัดเรื่องภายใน, แบบตอนมีรายการ, จุดสี, หัวเมลสั้น) -> R5, R8-R13, R3
- "มีความเสี่ยง" เปลี่ยนเป็น "มีความเสี่ยงล่าช้า" ทั้งหน้าจอและเมล -> R1
- ไม่ตัดภาษาอังกฤษ ใส่เฉพาะใต้หัวข้อ -> R2
- Q1 สรุปจาก AI วางบนสุด ใช้ภาษาไทย -> R16, AC-22, AC-23
- Q2 ข้อความเรื่องผู้รับตัดออกจากเมล แจ้ง PMO แยก -> R13, AC-17, AC-21
- Q3 ช้ากว่าเวลาและมี Fail เขียนเหตุผลทั้งสองท่อน -> R8, AC-9
- Q4 "At risk" เปลี่ยนเป็น "At risk of delay" -> R1, AC-1 ถึง AC-3
- Q4 ยืนยันอีกครั้งหลังแจ้งว่ากระทบ Excel: เปลี่ยนทั้งหมดรวม Excel และ template -> R1a, R1b, AC-3a, AC-3b
- ช่วงวันที่ของการเลื่อนวันใช้ "ถึง" คั่นแทน "-" เพราะวันที่มีขีดอยู่แล้ว อ่านยาก: "การเลื่อนวัน {from} ถึง {to}" เช่น "การเลื่อนวัน 18-Nov-2026 ถึง 24-Nov-2026" -> R2, R12, AC-16

## คำถามที่ยังไม่มีคำตอบ
- ไม่มี

---

# 2026-11-25 — ลบโปรเจกต์ (Delete project)

> requirement ผู้ใช้: "ต้องการให้สามารถลบโปรเจคออกได้" ข้อที่ผู้ใช้ตัดสินใจแล้วอยู่ใน "การตัดสินใจของผู้ใช้" ท้าย section ข้อที่ PM ตั้งเป็นสมมติฐานไว้ทำเครื่องหมาย **(สมมติฐาน — รอยืนยัน)**

## เป้าหมาย
ตอนนี้ระบบทำได้แค่ Archive (ซ่อนไว้ ข้อมูลยังอยู่ ย้อนกลับได้) ตามนโยบายเดิม "No delete" (`services/projects.py` docstring, trigger `project_no_delete` ใน `storage/migrations_app/001_app.sql`) โปรเจกต์ที่เลิกใช้หรือสร้างผิด (เช่น โปรเจกต์ทดสอบ) จึงค้างอยู่ในรายการ Archived และในโฟลเดอร์ข้อมูลตลอดไป รอบนี้เพิ่มการลบโปรเจกต์ที่ Archive แล้วออกถาวร: เอาออกจาก `app.db` และลบโฟลเดอร์ข้อมูล `<root>/projects/<id>/` โดยทำ backup zip ก่อนเสมอ และไม่แตะไฟล์ Excel ของผู้ใช้

ผู้ได้ประโยชน์: Admin / PMO ที่ดูแลรายการโปรเจกต์บนเครื่อง ให้รายการสะอาดและไม่มีข้อมูลค้าง

## ขอบเขต (In scope)
- **D1 service** ฟังก์ชันลบโปรเจกต์ใหม่ (เช่น `services/projects.py` `delete_project(base, appdb, person, project_id, *, typed_name)` หรือไฟล์ใหม่ `services/delete_project.py` เลือกโดย architect) ที่ทำตามลำดับ: ตรวจสิทธิ์ -> ตรวจว่า Archive แล้ว -> ตรวจชื่อที่พิมพ์ -> สร้าง backup zip -> ลบ (แบบ fail safely ตาม D4) -> บันทึก event
- **D2 สิทธิ์และเงื่อนไข** เฉพาะ Admin หรือ PMO (กติกาเดียวกับ Archive คือ `require_manager`: Admin/PMO ของโปรเจกต์ใดก็ได้) ลบได้เฉพาะโปรเจกต์ที่ `archived = 1` และต้องพิมพ์ชื่อโปรเจกต์ให้ตรง (ตัดช่องว่างหน้า/หลัง เทียบแบบตรงตัว เหมือน `clear_project`)
- **D3 สิ่งที่ถูกลบ**
  - แถวของโปรเจกต์ในตาราง `project` ของ `app.db`
  - ทั้งโฟลเดอร์ `<root>/projects/<id>/` (tracker.db และไฟล์ -wal/-shm, attachments, snapshots, previews, backups ของ workbook ที่อยู่ในโฟลเดอร์นี้ และไฟล์อื่นทั้งหมดในโฟลเดอร์)
- **D3a สิ่งที่ไม่ถูกลบ / ไม่ถูกแตะ**
  - ไฟล์ Excel ของโปรเจกต์ (`excel_path`) ห้ามเปิด อ่าน เขียน ย้าย หรือลบ
  - แถวเดิมใน `project_event` ของโปรเจกต์นั้น (ประวัติเก็บไว้) **(สมมติฐาน — รอยืนยัน)**
  - ไฟล์ log กลาง `<root>/logs/*.log`, `<root>/clear-project.log`, `<root>/clear-backups/` (backup จาก Clear project data ของโปรเจกต์นั้นที่เคยทำไว้)
  - โฟลเดอร์ workbook backup ที่ตั้งไว้นอกโฟลเดอร์โปรเจกต์ผ่าน `[sync] backup_dir` ใน config (ถ้ามี) **(สมมติฐาน — รอยืนยัน)**
  - โปรเจกต์อื่นทั้งหมด และ shared settings ใน `app.db` (holiday, PMO email, ภาษา, model)
- **D4 backup ก่อนลบ** **(สมมติฐาน — รอยืนยัน)**
  - ไฟล์ `<root>/delete-backups/<id>_<yyyymmdd-HHMMSS>.zip` (นอกโฟลเดอร์ที่ถูกลบ จึงยังอยู่หลังลบ และไม่ถูกลบอัตโนมัติ ตามนโยบาย "backup ไม่ถูกลบ")
  - เนื้อหา: `tracker.db` ที่ copy ด้วย SQLite backup API (ได้ข้อมูลที่ยังอยู่ใน -wal ด้วย เหมือน `clear_project._backup`), ไฟล์อื่นทุกไฟล์ในโฟลเดอร์โปรเจกต์ตามโครงสร้างเดิม (ยกเว้น `tracker.db`, `tracker.db-wal`, `tracker.db-shm`) และ `manifest.json` ที่มีแถว `project` เดิมทั้งแถวจาก `app.db` (id, name, excel_path, created_at, created_by, archived_at, window_start, window_end), เวลาที่ลบ และผู้ลบ
  - ถ้าสร้าง zip ไม่สำเร็จ (ดิสก์เต็ม, อ่านไฟล์ไม่ได้) ต้องหยุด ไม่ลบอะไรเลย และไม่ทิ้งไฟล์ zip ที่ไม่สมบูรณ์ไว้
- **D5 ลบแบบ fail safely** (กรณีไฟล์ถูก lock เช่น โปรเจกต์เปิดอยู่ในแอปอีกหน้าต่าง / scheduled job / Windows Explorer / OneDrive ถือไฟล์) **(สมมติฐาน — รอยืนยัน เรื่องวิธี)**
  1. ก่อนแก้ `app.db` ให้ "กัน" โฟลเดอร์ก่อน โดย rename `projects/<id>/` เป็นชื่อชั่วคราว (เช่น `projects/.deleting-<id>-<stamp>/`) ถ้า rename ไม่ได้ (ไฟล์ถูก lock / PermissionError / OSError) ให้หยุดทันที: แถวใน `app.db` ยังอยู่ครบ โฟลเดอร์ยังอยู่ชื่อเดิม ไม่มี event "deleted" แจ้ง error ที่บอกว่าไฟล์ถูกใช้งานอยู่ให้ปิดโปรแกรม/หน้าต่างอื่นแล้วลองใหม่ และบอก path ของ backup zip ที่สร้างไว้แล้ว (zip นี้เก็บไว้ ไม่ลบ)
  2. ลบแถว `project` และเพิ่ม event "deleted" ใน transaction เดียวของ `app.db` ถ้า transaction ล้มเหลว ให้ rename โฟลเดอร์กลับเป็นชื่อเดิม แล้วแจ้ง error
  3. ลบโฟลเดอร์ชั่วคราวทิ้ง ถ้าลบได้ไม่หมด (ไฟล์บางไฟล์ถูก lock) ถือว่าการลบโปรเจกต์สำเร็จแล้ว (`app.db` ไม่มีโปรเจกต์นี้) แต่ต้องแจ้งผู้ใช้ว่าเหลือโฟลเดอร์ใดค้างอยู่ให้ลบเองภายหลัง และบันทึก warning ลง file log
  4. ถ้าไม่มีโฟลเดอร์ `projects/<id>/` อยู่เลย (เคยถูกลบด้วยมือ) ยังลบแถวใน `app.db` ได้ backup zip มีแค่ `manifest.json`
- **D6 DB migration** เพิ่ม `storage/migrations_app/002_*.sql` ที่ drop trigger `project_no_delete` แล้วสร้าง trigger ใหม่ที่ยอม DELETE เฉพาะแถวที่ `archived = 1` (แถวที่ยังไม่ Archive ยัง abort) เป็นการกันชั้นที่สองนอกเหนือจาก service **(สมมติฐาน — รอยืนยัน)** ห้ามแก้ไฟล์ `001_app.sql` เดิม และแก้คอมเมนต์ "never deleted" ในจุดที่เกี่ยวข้องให้ตรงพฤติกรรมใหม่
- **D7 event** เพิ่มแถว `project_event` action `"deleted"` ที่มี `project_id`, `editor` และ `detail` ประกอบด้วยชื่อโปรเจกต์, `excel_path` และ path ของ backup zip แถว event เก่าของโปรเจกต์นั้นเก็บไว้ทั้งหมด **(สมมติฐาน — รอยืนยัน)**
- **D8 id ไม่ถูกใช้ซ้ำ** `appdb.next_project_id()` ปัจจุบันนับจาก max ของ id ที่อยู่ในตาราง `project` เท่านั้น ถ้าลบโปรเจกต์ที่ id สูงสุด id นั้นจะถูกใช้ซ้ำกับโปรเจกต์ใหม่ ทำให้ event เก่า, ชื่อ backup zip และ clear-backups ปนกับโปรเจกต์ใหม่ ต้องแก้ให้ id ใหม่มากกว่า id ที่เคยใช้ทั้งหมด (เช่น นับจาก `project` + `project_event` ด้วย) **(สมมติฐาน — รอยืนยัน)**
- **D9 ชื่อไฟล์ว่างให้ใช้ใหม่** หลังลบ ชื่อไฟล์ Excel ของโปรเจกต์ที่ลบไม่ถูกนับใน `duplicate_reason` / `import_project` / `relink` อีก (เป็นผลจากการเอาแถวออกจาก `app.db` อยู่แล้ว) จึง import ไฟล์ Excel เดิมเป็นโปรเจกต์ใหม่ได้ ส่วน create ด้วยชื่อเดิมในโฟลเดอร์เดิมยังถูกบล็อกตามกติกาเดิม เพราะไฟล์ Excel ยังอยู่ในโฟลเดอร์นั้น
- **D10 UI (Settings)** ในการ์ดโปรเจกต์ของหน้า Settings (`ui/screens/settings.py` ส่วนรายการ "Archived projects" ข้างปุ่ม "Restore selected") เพิ่มปุ่มสีแดง (ใช้ `DANGER_QSS`) "ลบโปรเจกต์ที่เลือกถาวร" / "Delete selected permanently" แสดงเฉพาะ manager และเฉพาะเมื่อมีรายการ Archived กดแล้วเปิด dialog แบบเดียวกับ `ClearProjectDialog` (`ui/projects.py`):
  - บอกชื่อโปรเจกต์ สิ่งที่จะถูกลบ (D3) สิ่งที่ไม่ถูกแตะ โดยเฉพาะไฟล์ Excel (D3a) และที่เก็บ backup zip (D4)
  - ช่องพิมพ์ชื่อโปรเจกต์ ปุ่มลบสีแดง enable เฉพาะเมื่อชื่อที่พิมพ์ตรง
  - สำเร็จ: แสดง path ของ backup zip (และโฟลเดอร์ที่ค้าง ถ้ามีตาม D5 ข้อ 3) แล้วรายการ Archived ใน Settings ถูก refresh โปรเจกต์ที่ลบหายจากรายการ
  - ไม่สำเร็จ: แสดง error ตาม D5 และรายการไม่เปลี่ยน
  - ไม่มีรายการที่เลือก: ไม่ทำอะไร (เหมือน `_unarchive`)
  - ข้อความทั้งหมดอยู่ใน `ui/i18n.py` ทั้ง EN และ TH
  - ไม่ผูกกับ feature flag `enable_clear_project` (ฟีเจอร์ลบโปรเจกต์ใช้ได้ตลอดสำหรับ manager) **(สมมติฐาน — รอยืนยัน)**
- **D11 CLI** `project delete <project> --confirm-name "<ชื่อโปรเจกต์>"` ใน `cli_project.py` (`<project>` = id หรือชื่อ ตาม `find_project` เดิม) ใช้ service เดียวกับ UI ไม่มี `--confirm-name` หรือชื่อไม่ตรง -> แสดงสิ่งที่จะถูกลบ แล้วจบด้วย "NOT DONE: ..." exit code 2 โดยไม่เปลี่ยนอะไร สำเร็จ -> พิมพ์ path ของ backup zip (และโฟลเดอร์ที่ค้าง ถ้ามี) exit code 0 อัปเดต help/docstring ของ `project` ให้มี `delete` **(สมมติฐาน — รอยืนยัน เรื่องรูปแบบ flag)**
- **D12 เอกสาร** แก้ docstring ของ `services/projects.py` ("Archive only ... No delete."), คอมเมนต์ใน migration ที่เกี่ยวข้อง และ README ส่วนที่อธิบาย Archive / คำสั่ง `project` ให้ตรงกับพฤติกรรมใหม่

## นอกขอบเขต (Out of scope)
- หน้าจอหรือคำสั่ง restore จาก backup zip ของการลบ (ใน zip มี `manifest.json` ไว้กู้ด้วยมือได้ถ้าจำเป็น) และไม่ดัดแปลง `project restore-clear` ให้รองรับ zip นี้
- การลบโปรเจกต์ที่ยังไม่ Archive (ต้อง Archive ก่อนเสมอ) และการลบหลายโปรเจกต์พร้อมกัน
- การลบหรือแก้ไฟล์ Excel ของโปรเจกต์ ไม่ว่ากรณีใด
- การลบ backup zip อัตโนมัติ / การกำหนดอายุ backup
- การลบ user / profile ที่อยู่เฉพาะในโปรเจกต์ที่ถูกลบออกจากที่อื่น (ข้อมูล user อยู่ใน tracker.db ของโปรเจกต์นั้นและหายไปพร้อมโฟลเดอร์)
- การเปลี่ยนกติกาสิทธิ์ของ `require_manager` หรือของ Archive / Unarchive
- การปิดโปรเจกต์ที่เปิดอยู่ในโปรแกรมหรือเครื่องอื่นให้อัตโนมัติ (ถ้าไฟล์ถูก lock ให้ fail ตาม D5)

## User stories
- **US-1** ในฐานะ Admin หรือ PMO ฉันต้องการลบโปรเจกต์ที่ Archive ไว้และไม่ใช้แล้วออกถาวร เพื่อให้รายการโปรเจกต์และโฟลเดอร์ข้อมูลไม่มีของค้าง
- **US-2** ในฐานะ Admin หรือ PMO ฉันต้องการให้ระบบบังคับพิมพ์ชื่อโปรเจกต์และทำ backup ก่อนลบ เพื่อกันการลบผิดและยังกู้ข้อมูลได้ถ้าจำเป็น
- **US-3** ในฐานะเจ้าของไฟล์ Excel ฉันต้องการมั่นใจว่าการลบโปรเจกต์ในแอปไม่แตะไฟล์ Excel ของฉันเลย
- **US-4** ในฐานะผู้ใช้ทั่วไป (ไม่ใช่ Admin/PMO) ฉันต้องไม่สามารถลบโปรเจกต์ได้ เพื่อกันข้อมูลหาย
- **US-5** ในฐานะผู้ดูแลระบบ ฉันต้องการลบโปรเจกต์ผ่าน CLI ได้ด้วยกติกาเดียวกับหน้าจอ

## Acceptance criteria
ทุกข้อทดสอบด้วย pytest โดยใช้ `tmp_path` เป็น root, `AppDb` จริง และโปรเจกต์ที่สร้างด้วย `create_project` / `import_project` หรือ helper ใน `tests/test_projects.py` ไม่ต้องใช้ Excel จริงหรือ COM ส่วน UI ใช้ `QApplication` แบบ offscreen ตามแนวของ `tests/test_ui.py` ส่วนการจำลองไฟล์ถูก lock ใช้ monkeypatch ของฟังก์ชัน rename / rmtree ให้ raise `PermissionError`

### สิทธิ์และเงื่อนไข (D2)
- **AC-1** (สำเร็จ) Given โปรเจกต์ P ที่ Archive แล้ว, ผู้ใช้เป็น Admin หรือ PMO (ทดสอบทั้งสอง role) When เรียก service ลบ P ด้วย `typed_name` = ชื่อ P Then `appdb.get_project(P.id)` เป็น `None`, P ไม่อยู่ใน `appdb.projects(include_archived=True)` และโฟลเดอร์ `<root>/projects/<P.id>/` ไม่มีอยู่
- **AC-2** (ไม่ใช่ manager) Given ผู้ใช้ที่ไม่ได้เป็น Admin/PMO ในโปรเจกต์ใดเลย (เช่น QA) When ลบ P Then raise `ProjectError` (หรือ error class ที่ architect กำหนด) ข้อความบอกว่าไม่มีสิทธิ์ และไม่มีอะไรเปลี่ยน (แถวใน `app.db`, โฟลเดอร์, `project_event`) และไม่มี backup zip ถูกสร้าง
- **AC-3** (ยังไม่ Archive) Given P ที่ `archived = 0` When ลบ P ด้วยชื่อที่ถูก Then raise error ที่บอกว่าต้อง Archive ก่อน ไม่มีอะไรเปลี่ยนและไม่มี backup zip
- **AC-4** (ชื่อไม่ตรง) Given P ที่ Archive แล้ว When `typed_name` ไม่ตรง (เช่น ว่าง, ชื่อโปรเจกต์อื่น, ตัวพิมพ์ต่าง) Then raise error ที่บอกว่าชื่อไม่ตรง ไม่มีอะไรเปลี่ยนและไม่มี backup zip Given `typed_name` = ชื่อ P มีช่องว่างหน้า/หลัง Then ลบได้
- **AC-5** (ไม่มีโปรเจกต์นี้) Given id ที่ไม่มีใน `app.db` When ลบ Then raise error "Unknown project" และไม่มีอะไรเปลี่ยน

### Excel ไม่ถูกแตะ (D3a)
- **AC-6** Given P ที่มีไฟล์ Excel อยู่จริง (บันทึกขนาด, mtime และ hash ไว้ก่อน) และ spy ของ `openpyxl.load_workbook` / opener ของ COM / `copy_shared` When ลบ P สำเร็จ Then ไฟล์ Excel ยังอยู่ที่ path เดิม ขนาด mtime และ hash เท่าเดิม และ spy ไม่ถูกเรียกกับไฟล์นั้น
- **AC-7** Given ไฟล์ Excel ของ P อยู่**ภายใน** `<root>/projects/<P.id>/` (กรณีผิดปกติ) When ลบ P Then ไฟล์ Excel ยังไม่ถูกลบ (ทั้งการ abort พร้อม error ที่ชัดเจนก่อนเปลี่ยนอะไร หรือย้ายไฟล์ไว้ที่อื่นก่อน architect เลือกได้ แต่ห้ามลบไฟล์ Excel)

### Backup (D4)
- **AC-8** (มี zip ก่อนลบ) Given P ที่มี tracker.db ที่มีข้อมูล (เช่น task 3 แถว) และไฟล์แนบ 1 ไฟล์ When ลบ P สำเร็จ Then มีไฟล์ zip ใหม่ 1 ไฟล์ใน `<root>/delete-backups/` ชื่อขึ้นต้นด้วย `<P.id>_`, ผลลัพธ์ของ service คืน path ของ zip นี้, ใน zip มี `tracker.db` ที่เปิดด้วย sqlite ได้และมี task 3 แถว, มีไฟล์แนบตาม path เดิม และมี `manifest.json` ที่มี `id`, `name`, `excel_path` ของ P และชื่อผู้ลบ
- **AC-9** (zip สร้างก่อนลบจริง) Given spy ที่บันทึกลำดับการเรียก When ลบ P Then zip ถูกสร้างเสร็จ (ปิดไฟล์แล้ว) ก่อนที่โฟลเดอร์จะถูก rename/ลบ และก่อนที่แถวใน `app.db` จะถูกลบ
- **AC-10** (backup ล้มเหลว -> ไม่ลบ) Given การสร้าง zip raise `OSError` (monkeypatch) When ลบ P Then error ถูกส่งออก, แถวของ P ยังอยู่ใน `app.db`, โฟลเดอร์ของ P ยังอยู่ครบ, ไม่มี event "deleted" และไม่มีไฟล์ zip ที่ไม่สมบูรณ์เหลือใน `delete-backups/`
- **AC-11** (zip รอดหลังลบ) Given ลบ P สำเร็จ When ตรวจ `<root>/delete-backups/` Then zip ยังอยู่ และ zip ไม่อยู่ภายใต้ `<root>/projects/`

### Fail safely (D5)
- **AC-12** (โฟลเดอร์ถูก lock -> ไม่ลบอะไร) Given การ rename โฟลเดอร์ของ P raise `PermissionError` When ลบ P Then error ที่ส่งออกบอกว่าไฟล์ถูกใช้งานอยู่และมี path ของ backup zip, แถวของ P ยังอยู่ใน `app.db` (ยัง `archived = 1`), โฟลเดอร์ของ P ยังอยู่ชื่อเดิมและไฟล์ครบ, ไม่มี event "deleted", และ backup zip ที่สร้างแล้วยังอยู่
- **AC-13** (app.db ล้มเหลว -> คืนโฟลเดอร์) Given rename โฟลเดอร์สำเร็จแต่การลบแถวใน `app.db` ล้มเหลว (เช่น monkeypatch ให้ transaction raise `sqlite3.OperationalError("database is locked")`) When ลบ P Then error ถูกส่งออก, แถวของ P ยังอยู่, โฟลเดอร์ถูก rename กลับเป็น `<root>/projects/<P.id>/` และไฟล์ครบ, ไม่มี event "deleted"
- **AC-14** (ลบโฟลเดอร์ได้ไม่หมด) Given app.db ถูกอัปเดตสำเร็จแต่การลบโฟลเดอร์ชั่วคราว raise `PermissionError` When ลบ P Then service ไม่ raise, ผลลัพธ์บอก path ของโฟลเดอร์ที่ค้าง, P ไม่อยู่ใน `app.db`, มี event "deleted", มี warning ลง file log (ตรวจด้วย `caplog`) และโฟลเดอร์ที่ค้างไม่ได้ชื่อ `projects/<P.id>` (ไม่ชนกับโปรเจกต์ใหม่)
- **AC-15** (ไม่มีโฟลเดอร์) Given P ที่ Archive แล้วแต่ไม่มีโฟลเดอร์ `projects/<P.id>/` When ลบ P Then ลบสำเร็จ แถวหายจาก `app.db`, มี event "deleted" และ backup zip มี `manifest.json`

### DB migration (D6)
- **AC-16** Given `app.db` เวอร์ชันเดิม (`user_version = 1` ที่มี trigger `project_no_delete`) ที่มีโปรเจกต์อยู่ When เปิดด้วย `AppDb` Then migration ใหม่ทำงาน `user_version` เพิ่มขึ้น และข้อมูลโปรเจกต์เดิมทั้งหมดอยู่ครบ
- **AC-17** Given `app.db` หลัง migration When `DELETE FROM project WHERE id=?` ตรงๆ กับแถวที่ `archived = 0` Then sqlite abort (raise `sqlite3.IntegrityError` หรือ error จาก trigger) และแถวยังอยู่ Given แถวที่ `archived = 1` Then ลบได้

### Event และ id (D7, D8)
- **AC-18** (event) Given P มี event เดิม (created, archived) When ลบ P สำเร็จ Then `appdb.events(P.id)` ยังมี event เดิมทั้งหมด และมี event ล่าสุด action `"deleted"`, editor = ผู้ลบ, detail มีชื่อ P, `excel_path` ของ P และ path ของ backup zip
- **AC-19** (id ไม่ซ้ำ) Given มีโปรเจกต์ p001, p002 แล้วลบ p002 (id สูงสุด) When สร้างโปรเจกต์ใหม่ Then id ใหม่คือ p003 (ไม่ใช่ p002) และ `appdb.events("p003")` ไม่มี event ของโปรเจกต์ที่ถูกลบ

### ชื่อไฟล์ว่าง (D9)
- **AC-20** Given P ที่ import จากไฟล์ `X.xlsx` ถูกลบแล้ว When `import_project` ไฟล์ `X.xlsx` เดิมอีกครั้ง Then import สำเร็จเป็นโปรเจกต์ใหม่ (id ใหม่) และ `duplicate_reason` ไม่อ้างถึงโปรเจกต์ที่ถูกลบ
- **AC-21** Given P ที่สร้างด้วย `create_project` ชื่อ N ในโฟลเดอร์ F ถูกลบแล้ว (ไฟล์ Excel ยังอยู่ใน F) When `create_project` ชื่อ N ในโฟลเดอร์ F Then ยังถูกบล็อกด้วยเหตุผล "already exists in" F (ไม่ใช่เพราะโปรเจกต์ที่ถูกลบ) When `preview_create` ชื่อ N ในโฟลเดอร์อื่น Then ไม่มีปัญหา

### ไม่กระทบโปรเจกต์อื่น
- **AC-22** Given โปรเจกต์ Q (ไม่ Archive) และ R (Archive) อยู่คู่กับ P When ลบ P Then แถว, โฟลเดอร์, tracker.db และไฟล์ Excel ของ Q และ R ไม่เปลี่ยน และ shared settings ใน `app.db` (เช่น `settings.pmo_email`) ไม่เปลี่ยน

### UI (D10)
- **AC-23** (ปุ่มแสดงเมื่อควร) Given manager และมีโปรเจกต์ Archived When เปิดหน้า Settings Then มีปุ่ม Delete (ข้อความจาก `t()`) ใต้รายการ Archived ใช้สไตล์ `DANGER_QSS` Given ผู้ใช้ไม่ใช่ manager หรือไม่มีโปรเจกต์ Archived Then ไม่มีปุ่มนี้
- **AC-24** (dialog ต้องพิมพ์ชื่อ) Given เลือกโปรเจกต์ P ในรายการ Archived แล้วกดปุ่ม Delete When dialog เปิด Then ปุ่มลบใน dialog disabled, ข้อความใน dialog มีชื่อ P และบอกว่าไฟล์ Excel ไม่ถูกแตะ When พิมพ์ชื่อผิด Then ยัง disabled When พิมพ์ชื่อ P ถูก Then enabled
- **AC-25** (สำเร็จ) Given dialog ที่พิมพ์ชื่อถูก When กดลบ Then service ถูกเรียกกับ P, ผู้ใช้เห็นข้อความที่มี path ของ backup zip และหลังปิดข้อความ รายการ Archived ใน Settings ไม่มี P
- **AC-26** (ล้มเหลว) Given service raise error (เช่น ไฟล์ถูก lock) When กดลบ Then แสดง error ผ่าน `error()` และรายการ Archived ยังมี P
- **AC-27** (ไม่ได้เลือก) Given ไม่มีแถวที่เลือก When กดปุ่ม Delete Then ไม่เปิด dialog และไม่เรียก service
- **AC-28** (i18n) Given ทุก key ใหม่ที่ใช้ในฟีเจอร์นี้ When ตรวจ `ui/i18n.py` Then มีทั้งข้อความ EN และ TH ที่ไม่ว่าง

### CLI (D11)
- **AC-29** (สำเร็จ) Given P Archive แล้ว และ profile บนเครื่องเป็น Admin/PMO When รัน `project delete <P.id> --confirm-name "<ชื่อ P>"` (และด้วยชื่อแทน id) Then exit code 0, output มี path ของ backup zip และ P ไม่อยู่ใน `app.db`
- **AC-30** (ไม่ยืนยัน / ยืนยันผิด) Given P Archive แล้ว When รัน `project delete <P.id>` โดยไม่มี `--confirm-name` หรือชื่อไม่ตรง Then exit code 2, output ขึ้นต้นด้วย "NOT DONE:" และไม่มีอะไรเปลี่ยน (รวมถึงไม่มี backup zip)
- **AC-31** (เงื่อนไขอื่น) Given P ยังไม่ Archive / ผู้ใช้ไม่ใช่ manager / ไม่มี profile / project ไม่มีอยู่ When รัน `project delete` Then exit code ไม่ใช่ 0 ตามแบบคำสั่งอื่นใน `cli_project.py` (2 สำหรับ NOT DONE และ Unknown project, 3 สำหรับไม่มี profile) และไม่มีอะไรเปลี่ยน
- **AC-32** `project list` หลังลบ P ไม่แสดง P

### ทั่วไป
- **AC-33** `pytest -q` ผ่านทั้งหมด test เดิมที่ตรวจว่า "ลบโปรเจกต์ไม่ได้" (ถ้ามี เช่น test ที่ยิง `DELETE FROM project` แล้วคาดว่า abort) ต้องปรับให้ตรงกับ AC-17 ห้ามลบ test ทิ้ง และต้องรายงานหัวหน้าว่าแก้ test ใด

## ข้อจำกัด / สมมติฐาน
- Python 3.x, PySide6, SQLite, PEP 8, type hints, pytest
- นโยบายเดิม "No delete" (approved 10-Nov-2026) ถูกเปลี่ยนโดยผู้ใช้ในรอบนี้ เฉพาะโปรเจกต์ที่ Archive แล้ว
- `project_event` ไม่มี foreign key ไปที่ `project` จึงเก็บ event เดิมได้โดยไม่ต้องแก้ schema
- log ไฟล์ของแอปอยู่ที่ `<root>/logs/` (ใช้ร่วมกันทุกโปรเจกต์) ไม่ได้อยู่ใน `projects/<id>/` "logs" ของโปรเจกต์ในที่นี้คือตาราง log ใน tracker.db (sync_event, email_log, job_run ฯลฯ) ซึ่งหายไปพร้อม tracker.db และอยู่ใน backup zip
- โปรเจกต์ที่ Archive แล้วไม่ถูกเปิดเป็นโปรเจกต์ปัจจุบันในแอป (`open_best` เลือกจากโปรเจกต์ที่ไม่ Archive) แต่ทุกจุดที่เปิด tracker.db ของโปรเจกต์ชั่วคราว (`_with_repo` เช่น `role_in`, `can_manage`, `find_person`) ต้องปิด connection ก่อนลบ ถ้า service เองเปิด tracker.db (เช่น ตรวจสิทธิ์ หรือทำ backup) ต้องปิดก่อนขั้น rename ตาม D5
- หลังลบ ถ้าโปรเจกต์ที่ลบเป็นที่เดียวที่ผู้ใช้เป็น Admin/PMO ผู้ใช้จะเสียสิทธิ์ manager (ตาม `can_manage` เดิม) และถ้าลบจนไม่เหลือโปรเจกต์เลย ระบบกลับไปเป็นสถานะ first use (ทุกคนจัดการได้) ซึ่งเป็นพฤติกรรมของกติกาเดิม ไม่ต้องแก้
- `last_project` ใน profile ที่ชี้ไปยังโปรเจกต์ที่ลบไม่ต้องล้าง เพราะ `open_best` / `cli.py` ข้ามไปเองเมื่อหาไม่เจอ (ต้องไม่ error)
- ถ้าทุกโปรเจกต์ถูก Archive จนเข้าหน้า Settings ไม่ได้ ให้ใช้ CLI `project delete` แทน (ไม่ต้องเพิ่มทางเข้า UI ใหม่)
- ขนาด backup zip อาจใหญ่ถ้าโฟลเดอร์ `backups` มี workbook backup จำนวนมาก ยอมรับได้
- รายละเอียดทางเทคนิค (ชื่อฟังก์ชัน, ที่อยู่ของ service, error class, รูปแบบผลลัพธ์) เป็นหน้าที่ architect ตราบที่ผ่าน AC

### สรุปสมมติฐานที่รอผู้ใช้ยืนยัน
- **A1** backup zip เก็บที่ `<root>/delete-backups/<id>_<stamp>.zip` อยู่นอกโฟลเดอร์ที่ถูกลบ เก็บถาวรไม่ลบอัตโนมัติ (D4)
- **A2** เก็บ event เดิมของโปรเจกต์ใน `project_event` และเพิ่ม event "deleted" (D7)
- **A3** migration ใหม่ `002` แทน trigger เดิมด้วย trigger ที่ยอมลบเฉพาะแถว `archived = 1` (D6)
- **A4** วิธี fail safely: สร้าง zip -> rename โฟลเดอร์ -> ลบแถว app.db -> ลบโฟลเดอร์ ถ้า rename ไม่ได้ไม่ลบอะไร ถ้าลบโฟลเดอร์ได้ไม่หมดถือว่าลบสำเร็จแต่แจ้งโฟลเดอร์ที่ค้าง (D5)
- **A5** id โปรเจกต์ไม่ถูกใช้ซ้ำหลังลบ (D8)
- **A6** ฟีเจอร์ลบไม่ผูกกับ feature flag `enable_clear_project` (D10)
- **A7** CLI ใช้ `--confirm-name "<ชื่อ>"` แทนการพิมพ์ชื่อแบบ interactive (D11)
- **A8** workbook backup ที่ตั้งไว้นอกโฟลเดอร์โปรเจกต์ผ่าน `[sync] backup_dir` ไม่ถูกลบและไม่อยู่ใน zip (D3a)

## การตัดสินใจของผู้ใช้ (2026-11-25)
- ลบอะไร: เอาโปรเจกต์ออกจาก `app.db` และลบโฟลเดอร์ `projects/<id>/` (tracker.db, attachments, logs) ห้ามแตะไฟล์ Excel ทำ backup zip ก่อนลบ -> D3, D3a, D4, AC-1, AC-6 ถึง AC-11
- ใครและเมื่อไร: เฉพาะ Admin หรือ PMO (กติกาเดียวกับ Archive) เฉพาะโปรเจกต์ที่ Archive แล้ว ต้องพิมพ์ชื่อโปรเจกต์ยืนยัน (แบบ Clear project data) -> D2, AC-1 ถึง AC-4
- ที่ไหน: หน้า Settings (ข้างรายการ Archived projects) และ CLI `project delete` -> D10, D11, AC-23 ถึง AC-32

## คำถามที่ยังไม่มีคำตอบ
- ขอให้ผู้ใช้ยืนยันสมมติฐาน A1 ถึง A8 ด้านบน (ถ้าไม่ตอบ ใช้ตามที่เขียนไว้)

---

# 2026-11-25 — เลือกโฟลเดอร์เก็บไฟล์ผลทดสอบแยกตามโปรเจกต์ (Per-project attachment folder)

> requirement ผู้ใช้: ต้องเลือกได้ว่าไฟล์ผลทดสอบที่อัปโหลดจะเก็บที่ไหน ผู้ใช้เลือกแบบ **ตั้งค่าแยกต่อโปรเจกต์** แต่ละโปรเจกต์มีโฟลเดอร์ปลายทางของตัวเองได้ (เช่น โฟลเดอร์ OneDrive / SharePoint ที่ sync ลงเครื่อง) ไฟล์ทุกไฟล์ที่อัปโหลดในโปรเจกต์นั้นไปที่โฟลเดอร์นั้น ถ้าโปรเจกต์ไม่ได้ตั้งไว้ ให้ทำงานเหมือนเดิม (`<data_dir>/attachments` = `%LOCALAPPDATA%\PlanningProgress\projects\<id>\attachments`)
> ข้อที่ PM ตั้งเป็นสมมติฐานไว้ทำเครื่องหมาย **(สมมติฐาน — รอยืนยัน)** และสรุปไว้ท้าย section (B1 ถึง B9)

## เป้าหมาย
ตอนนี้ `services/attachments.py` `upload()` copy ไฟล์ไปที่ `cfg.attachment_dir` (`config.py`: `data_dir / "attachments"`) เสมอ ซึ่งอยู่ใน `%LOCALAPPDATA%` ของเครื่องที่อัปโหลด คนอื่นในทีมจึงเปิดไฟล์ผลทดสอบไม่ได้ และไฟล์ไม่ได้ถูก backup ไปกับ OneDrive / SharePoint ของทีม รอบนี้ให้แต่ละโปรเจกต์ตั้งโฟลเดอร์ปลายทางของไฟล์แนบเองได้ ถ้าไม่ตั้ง ใช้ที่เดิม

ผู้ได้ประโยชน์: Admin / PMO ที่ต้องการให้ไฟล์ผลทดสอบอยู่ในที่เก็บกลางของทีม และ QA / User BU/PM ที่อัปโหลดไฟล์ (ไม่ต้องทำอะไรเพิ่ม ไฟล์ไปที่ที่ถูกต้องเอง)

## ขอบเขต (In scope)
- **F1 การเก็บค่า** เพิ่มค่า "โฟลเดอร์ไฟล์แนบ" ต่อโปรเจกต์ เก็บเป็น column ใหม่ที่ว่างได้ (เช่น `attachment_dir TEXT`, NULL = ค่าเริ่มต้น) ในตาราง `project` ของ `app.db` ผ่าน migration ใหม่ `storage/migrations_app/003_*.sql` (ห้ามแก้ `001_app.sql`, `002_project_delete.sql`) และเพิ่มใน `PROJECT_COLS` ของ `storage/app_db.py` ให้ `update_project` ใช้ได้ **(สมมติฐาน — รอยืนยัน: เก็บใน `app.db` แบบเดียวกับ `window_start` / `window_end` ไม่ใช่ใน settings ของ tracker.db)**
- **F2 การใช้ค่า** `projects.project_config()` ส่งค่านี้เข้า `Config` ทำให้ `cfg.attachment_dir` คืน
  - path ที่ตั้งไว้ ถ้าโปรเจกต์มีค่า
  - `data_dir / "attachments"` ถ้าค่าเป็น NULL หรือว่าง (พฤติกรรมเดิมทุกประการ)
  - โหมด single-file (ไม่มี project, `cfg.project_id is None`) ใช้ `data_dir / "attachments"` เหมือนเดิม
  ทุกทางที่อัปโหลด (หน้า Test Progress และ CLI `cli_alerts.py` ที่เรียก `upload()`) ได้ผลเดียวกันเพราะใช้ `cfg` ของโปรเจกต์
- **F3 ผู้ที่ตั้ง / เปลี่ยนได้** เฉพาะ Admin หรือ PMO **ของโปรเจกต์นั้น** (ตรวจด้วย `role_in(base, project, email)` อยู่ใน `MANAGER_ROLES`) role อื่น (QA, User BU/PM) และผู้ที่ไม่ได้เป็นสมาชิกโปรเจกต์นั้น ทำไม่ได้ **(สมมติฐาน — รอยืนยัน: ต่างจาก Relink / Archive ที่ใช้ `require_manager` คือ Admin/PMO ของโปรเจกต์ใดก็ได้)**
- **F4 service** ฟังก์ชันใหม่ใน `services/projects.py` (เช่น `set_attachment_dir(base, appdb, person, project_id, folder: Path | None)`, ชื่อจริง architect เลือก) ที่ทำตามลำดับ: ตรวจว่ามีโปรเจกต์ -> ตรวจสิทธิ์ (F3) -> ตรวจโฟลเดอร์ (F5) -> บันทึกค่าและ event (F8) ใน transaction เดียวของ `app.db` `folder=None` (หรือ string ว่าง) = กลับไปใช้ค่าเริ่มต้น
- **F5 การตรวจตอนตั้งค่า** (ไม่ผ่านข้อใด -> raise `ProjectError` ที่บอกเหตุผล และไม่บันทึกอะไร ไม่มี event)
  1. ตัดช่องว่างหน้า/หลังและเครื่องหมาย `"` ที่ครอบ path ออกก่อนตรวจ
  2. ต้องเป็น absolute path (รวม UNC เช่น `\\server\share\...`) path แบบ relative ไม่รับ
  3. ต้องมีอยู่จริงและเป็นโฟลเดอร์ (`is_dir()`) ระบบไม่สร้างโฟลเดอร์ปลายทางให้เอง
  4. ต้องเขียนได้: ทดลองสร้างไฟล์ชั่วคราวในโฟลเดอร์นั้นแล้วลบทิ้ง (สร้างหรือลบไม่ได้ = ไม่ผ่าน) หลังตรวจต้องไม่มีไฟล์ทดสอบเหลือในโฟลเดอร์
  5. ห้ามอยู่ภายใต้ `<root>/projects/` (โฟลเดอร์ข้อมูลของโปรเจกต์ใดก็ได้ในเครื่อง) เพราะ Delete project / Clear project data ลบโฟลเดอร์ในนั้นได้ ถ้าต้องการใช้ที่เดิมให้กด "ใช้ค่าเริ่มต้น" แทน **(สมมติฐาน — รอยืนยัน)**
  6. ตั้งค่าเป็นค่าเดิม (path เดียวกันหลัง normalize) -> ไม่บันทึกและไม่มี event ใหม่ (ไม่ใช่ error)
  การกลับไปใช้ค่าเริ่มต้นไม่ต้องตรวจข้อ 2 ถึง 5
- **F6 การตรวจตอนอัปโหลด** (โฟลเดอร์ที่ตั้งไว้อาจหายหรือใช้ไม่ได้ภายหลัง เช่น OneDrive ยังไม่ sync / ยังไม่ login, ไดรฟ์เครือข่ายหลุด, โฟลเดอร์ถูกลบหรือเปลี่ยนชื่อ, สิทธิ์ถูกถอน)
  1. ก่อน copy ถ้าโฟลเดอร์ที่ตั้งไว้ไม่มีอยู่หรือไม่ใช่โฟลเดอร์ -> raise error เฉพาะ (เช่น subclass ของ `UploadError` ชื่อ `AttachmentFolderUnavailable` ที่มี path อยู่ในตัว, architect เลือกชื่อ) ระบบ**ไม่สร้าง**โฟลเดอร์ที่ตั้งไว้ใหม่เอง (กันการสร้างโฟลเดอร์ปลอมนอก OneDrive ตอนไดรฟ์ยังไม่ mount) แต่ยังสร้างโฟลเดอร์ย่อยของ item (F7) ภายในโฟลเดอร์ที่มีอยู่ได้เหมือนเดิม
  2. ถ้าสร้างโฟลเดอร์ย่อย / copy / อ่านไฟล์ปลายทางเพื่อตรวจ hash แล้วเกิด `OSError` (เช่น `PermissionError`, ดิสก์เต็ม, ไดรฟ์หลุดกลางทาง) -> raise error เดียวกับข้อ 1 และลบไฟล์ปลายทางที่ copy ไปได้ไม่ครบ (ถ้าลบได้)
  3. ทั้งข้อ 1 และ 2: ไม่มีแถว attachment ใหม่, ไม่มี audit, TC/Result Complete ไม่เปลี่ยน, ไม่มี pending write-back (เหมือนกรณี error อื่นของ `upload()` ในปัจจุบัน)
  4. **ไม่ fallback** ไปเก็บที่ `<data_dir>/attachments` แบบเงียบ ๆ เพราะผู้ใช้ตั้งใจให้ไฟล์อยู่ในที่เก็บกลาง **(สมมติฐาน — รอยืนยัน)**
  5. บันทึก warning ลง file log (`<root>/logs/`) ที่มี project id, item id และ path ของโฟลเดอร์
  6. หน้า Test Progress แสดงข้อความแปลแล้ว (F10 key `tp.folder_unavailable`) ที่มี path ของโฟลเดอร์ และแนะนำให้ตรวจว่า OneDrive / ไดรฟ์ sync อยู่ แล้วลองใหม่ หรือให้ Admin/PMO เปลี่ยนโฟลเดอร์
  7. ไม่ตรวจสถานะการ sync ของ OneDrive โดยตรง (ไม่มี API ที่เชื่อถือได้) ถ้าโฟลเดอร์มีอยู่และเขียนได้ ถือว่าใช้ได้
- **F7 โครงสร้างโฟลเดอร์ย่อย** ภายในโฟลเดอร์ที่เลือกใช้โครงสร้างเดียวกับปัจจุบัน: `<โฟลเดอร์ที่เลือก>/<item_id ที่ sanitize แล้ว>/<stamp>_<ชื่อไฟล์เดิม>` (กติกา sanitize, stamp และการเติม `_2`, `_3` เมื่อชื่อซ้ำ คงเดิม ไฟล์เดิมไม่ถูกเขียนทับไม่ว่ากรณีใด) **(สมมติฐาน — รอยืนยัน: ไม่เพิ่มชั้นชื่อโปรเจกต์ ดูคำถาม Q2)**
- **F8 audit / event**
  - ทุกครั้งที่ค่าเปลี่ยน (ตั้งใหม่, เปลี่ยน, กลับเป็นค่าเริ่มต้น) เพิ่มแถว `project_event` ใน `app.db`: `action = "attachment_dir"`, `project_id`, `editor = person.editor`, `detail = "<ค่าเดิม> -> <ค่าใหม่>"` โดยค่าว่างแสดงเป็น `(default)` (แบบเดียวกับ action `"window"` ของ `set_window`)
  - การอัปโหลดยังบันทึก audit ใน tracker.db เหมือนเดิม (field `attachment` และ `tc_result_complete`) และแถว attachment มี `stored_path` เป็น path เต็มของไฟล์ในโฟลเดอร์ที่ใช้จริง ณ ตอนอัปโหลด
- **F9 ไฟล์เดิมไม่ย้าย** เมื่อเปลี่ยนโฟลเดอร์หรือกลับเป็นค่าเริ่มต้น ไฟล์ที่อัปโหลดไปแล้วอยู่ที่เดิม ไม่ถูก copy / move / ลบ และ `stored_path` ของแถว attachment เดิมไม่ถูกแก้ เฉพาะการอัปโหลดหลังเปลี่ยนเท่านั้นที่ไปโฟลเดอร์ใหม่ สถานะ "อัปโหลดแล้ว" ของ item (`views.py`, `daily_job.py` ที่ดูจาก `repo.attachments()`) ไม่เปลี่ยน
- **F10 UI (หน้า Settings, การ์ดโปรเจกต์ `_project_card` ใน `ui/screens/settings.py`)** เพิ่มแถว "โฟลเดอร์เก็บไฟล์ผลทดสอบ" ใต้แถว Excel file
  - แสดงค่าปัจจุบัน: path ที่ตั้งไว้ หรือข้อความ "ค่าเริ่มต้น (เครื่องนี้)" พร้อม path `<data_dir>/attachments`
  - ปุ่ม "เลือกโฟลเดอร์…" (เปิด `QFileDialog.getExistingDirectory`) และปุ่ม "ใช้ค่าเริ่มต้น" (แสดง/enable เฉพาะเมื่อมีค่าที่ตั้งไว้) แสดงเฉพาะผู้ที่มีสิทธิ์ตาม F3 ผู้อื่นเห็นค่าปัจจุบันแบบอ่านอย่างเดียว
  - หมายเหตุสีจาง: ไฟล์ที่อัปโหลดแล้วไม่ย้ายตาม (F9) และค่านี้ตั้งบนเครื่องนี้เท่านั้น
  - เลือกโฟลเดอร์ / กดใช้ค่าเริ่มต้นแล้วบันทึกทันที (แบบ Relink ไม่ผูกกับปุ่ม Save ของหน้า) โดยถามยืนยันก่อน สำเร็จ -> แสดงข้อความสำเร็จ, reload `cfg` ของโปรเจกต์ปัจจุบันให้การอัปโหลดถัดไปใช้ค่าใหม่ทันทีโดยไม่ต้อง restart และ refresh การ์ด ไม่สำเร็จ -> แสดง error ผ่าน `error()` และค่าไม่เปลี่ยน
  - ยกเลิก dialog เลือกโฟลเดอร์ หรือตอบ "ไม่" ตอนยืนยัน -> ไม่เรียก service
- **F11 i18n** ข้อความใหม่ทุกข้อความอยู่ใน `ui/i18n.py` มีทั้ง EN และ TH (ข้อความจริงปรับได้โดย architect/developer ตราบที่ความหมายตรง) ข้อเสนอ:

  | key | EN | TH |
  |---|---|---|
  | `proj.att_dir` | Result files folder | โฟลเดอร์เก็บไฟล์ผลทดสอบ |
  | `proj.att_default` | Default (this PC): {p} | ค่าเริ่มต้น (เครื่องนี้): {p} |
  | `proj.att_choose` | Choose folder… | เลือกโฟลเดอร์… |
  | `proj.att_reset` | Use default | ใช้ค่าเริ่มต้น |
  | `proj.att_note` | New uploads of this project go here. Files uploaded earlier stay where they are. This setting is for this PC only. | ไฟล์ที่อัปโหลดใหม่ของโปรเจกต์นี้จะเก็บที่นี่ ไฟล์ที่อัปโหลดไปแล้วอยู่ที่เดิม ค่านี้ใช้กับเครื่องนี้เท่านั้น |
  | `proj.att_q` | Save uploaded result files of '{n}' in:\n{p}? | เก็บไฟล์ผลทดสอบที่อัปโหลดของ '{n}' ที่:\n{p} ใช่หรือไม่ |
  | `proj.att_reset_q` | Use the default folder on this PC for '{n}'? | ให้ '{n}' กลับไปใช้โฟลเดอร์ค่าเริ่มต้นบนเครื่องนี้ใช่หรือไม่ |
  | `proj.att_saved` | Result files folder saved. | บันทึกโฟลเดอร์เก็บไฟล์ผลทดสอบแล้ว |
  | `proj.att_not_found` | Folder not found: {p} | ไม่พบโฟลเดอร์: {p} |
  | `proj.att_not_writable` | The app can't write to this folder: {p} | แอปเขียนไฟล์ลงโฟลเดอร์นี้ไม่ได้: {p} |
  | `proj.att_not_absolute` | Choose a full folder path | กรุณาเลือกโฟลเดอร์แบบ path เต็ม |
  | `proj.att_inside_data` | This folder is inside the app's project data and may be deleted with a project. Choose another folder. | โฟลเดอร์นี้อยู่ในข้อมูลโปรเจกต์ของแอปและอาจถูกลบไปพร้อมโปรเจกต์ กรุณาเลือกโฟลเดอร์อื่น |
  | `tp.folder_unavailable` | Can't save to the result files folder:\n{p}\nCheck that OneDrive / the network drive is connected and synced, then try again, or ask an Admin/PMO to change the folder. Nothing was saved. | บันทึกลงโฟลเดอร์เก็บไฟล์ผลทดสอบไม่ได้:\n{p}\nกรุณาตรวจว่า OneDrive / ไดรฟ์เครือข่ายเชื่อมต่อและ sync อยู่ แล้วลองใหม่ หรือแจ้ง Admin/PMO ให้เปลี่ยนโฟลเดอร์ ยังไม่มีการบันทึกใด ๆ |

  error จาก service (`ProjectError`) ปัจจุบันเป็นภาษาอังกฤษ ให้ UI แสดงข้อความแปลจาก key ข้างบนสำหรับกรณี F5 ข้อ 2 ถึง 5 (architect เลือกวิธี map เช่น error code ใน exception)
- **F12 ความเข้ากันกับ Clear project data และ Delete project** ไฟล์ในโฟลเดอร์ที่ผู้ใช้เลือก (นอก `<root>/projects/<id>/`) **ไม่ถูกลบ ไม่ถูก zip และไม่ถูกเขียนทับ** โดย Clear project data, การ restore ของ Clear project data และ Delete project ทั้งสามฟีเจอร์ทำงานกับ `<data_dir>/attachments` ของโปรเจกต์เหมือนเดิมเท่านั้น (ปัจจุบัน `clear_project._backup` และ `restore_backup` อ้าง `cfg.attachment_dir` ต้องเปลี่ยนให้อ้างโฟลเดอร์ local `<data_dir>/attachments` ไม่ใช่โฟลเดอร์ที่ตั้งไว้) เพราะโฟลเดอร์ที่เลือกอาจเป็นที่เก็บกลางที่ใช้ร่วมกับโปรเจกต์อื่นหรือคนอื่น **(สมมติฐาน — รอยืนยัน)**
- **F13 เอกสาร** แก้ docstring บรรทัดแรกของ `services/attachments.py` ("copied under <data_dir>/attachments/...") ให้ตรงกับพฤติกรรมใหม่ และ README ส่วนไฟล์แนบ (ถ้ามี)

## นอกขอบเขต (Out of scope)
- การย้าย / copy ไฟล์ที่อัปโหลดไปแล้วไปโฟลเดอร์ใหม่ และการแก้ `stored_path` เดิม (F9)
- การตั้งค่าโฟลเดอร์รวมสำหรับทุกโปรเจกต์ (global) หรือแยกตาม user
- การ sync ค่านี้ข้ามเครื่อง (`app.db` อยู่ในเครื่อง แต่ละเครื่องต้องตั้งเอง)
- การตรวจสถานะ sync ของ OneDrive / SharePoint, การเรียก SharePoint / Graph API โดยตรง
- การ fallback ไปเก็บ local แล้ว sync ขึ้นภายหลัง (คิวอัปโหลดค้าง)
- การสร้างโฟลเดอร์ปลายทางให้อัตโนมัติ
- หน้าจอเปิด / ดาวน์โหลด / ลบไฟล์แนบ
- คำสั่ง CLI สำหรับตั้งโฟลเดอร์ (ดูคำถาม Q5)
- การเปลี่ยนชนิดไฟล์ที่รับ, ขนาดสูงสุด, หรือกติกาการตรวจ SHA-256
- การเปลี่ยนสิทธิ์การอัปโหลด (`auth.UPLOAD`)

## User stories
- **US-1** ในฐานะ Admin หรือ PMO ของโปรเจกต์ ฉันต้องการตั้งโฟลเดอร์เก็บไฟล์ผลทดสอบของโปรเจกต์เป็นโฟลเดอร์ OneDrive / SharePoint ของทีม เพื่อให้ทีมเปิดไฟล์ได้และไฟล์ถูก backup
- **US-2** ในฐานะ Admin หรือ PMO ฉันต้องการเปลี่ยนโฟลเดอร์หรือกลับไปใช้ค่าเริ่มต้นได้ โดยไฟล์เดิมไม่หายหรือย้าย เพื่อไม่ให้ลิงก์ไฟล์เดิมเสีย
- **US-3** ในฐานะ QA / User BU/PM ที่อัปโหลดไฟล์ ฉันต้องการให้ไฟล์ไปอยู่ในโฟลเดอร์ของโปรเจกต์เองโดยไม่ต้องเลือก และถ้าโฟลเดอร์ใช้ไม่ได้ ฉันต้องรู้ทันทีพร้อมวิธีแก้ เพื่อไม่ให้คิดว่าอัปโหลดสำเร็จทั้งที่ไม่สำเร็จ
- **US-4** ในฐานะ PMO ฉันต้องการเห็นประวัติว่าใครเปลี่ยนโฟลเดอร์เมื่อไร จากไหนไปไหน เพื่อตรวจสอบย้อนหลัง
- **US-5** ในฐานะผู้ใช้โปรเจกต์ที่ไม่ได้ตั้งโฟลเดอร์ ฉันต้องการให้ทุกอย่างทำงานเหมือนเดิม

## Acceptance criteria
ทุกข้อทดสอบด้วย pytest โดยใช้ `tmp_path` เป็น root และเป็นโฟลเดอร์ปลายทาง (จำลองโฟลเดอร์ OneDrive ด้วยโฟลเดอร์ธรรมดา), `AppDb` จริง และโปรเจกต์ที่สร้างด้วย helper ใน `tests/test_projects.py` / `tests/test_attachments*.py` ไม่ต้องใช้ Excel จริงหรือ COM ส่วน UI ใช้ `QApplication` แบบ offscreen ตามแนว `tests/test_ui.py` โดย monkeypatch `QFileDialog.getExistingDirectory`, `confirm`, `info`, `error` ส่วนกรณีโฟลเดอร์เขียนไม่ได้ / ไดรฟ์หลุดจำลองด้วย monkeypatch ให้ raise `PermissionError` / `OSError` (ไม่พึ่ง ACL ของ Windows)

### การเก็บค่าและ config (F1, F2)
- **AC-1** (migration) Given `app.db` เวอร์ชันก่อนหน้า (มีโปรเจกต์อยู่ 2 โปรเจกต์) When เปิดด้วย `AppDb` Then `user_version` เพิ่มขึ้น, ตาราง `project` มี column ใหม่ ค่าเป็น NULL ทุกแถว และข้อมูลเดิมทุก column อยู่ครบ
- **AC-2** (ค่าเริ่มต้น) Given โปรเจกต์ที่ไม่ได้ตั้งโฟลเดอร์ (NULL และกรณีเป็น string ว่าง) When `project_config(base, project)` Then `cfg.attachment_dir == <root>/projects/<id>/attachments`
- **AC-3** (ตั้งไว้) Given โปรเจกต์ที่ตั้งโฟลเดอร์ `X` When `project_config(base, project)` Then `cfg.attachment_dir == Path(X)` และ `cfg.data_dir` ยังเป็น `<root>/projects/<id>` (ค่าอื่นไม่เปลี่ยน)
- **AC-4** (single-file) Given `Config` ที่ `project_id is None` Then `cfg.attachment_dir == data_dir / "attachments"`
- **AC-5** (แยกต่อโปรเจกต์) Given โปรเจกต์ A ตั้งโฟลเดอร์ `X` และโปรเจกต์ B ไม่ได้ตั้ง When อัปโหลดไฟล์ในแต่ละโปรเจกต์ Then ไฟล์ของ A อยู่ใต้ `X` ไฟล์ของ B อยู่ใต้ `<root>/projects/<B.id>/attachments`

### การตั้งค่าและสิทธิ์ (F3, F4, F5)
- **AC-6** (สำเร็จ) Given ผู้ใช้เป็น Admin ของโปรเจกต์ P (ทดสอบซ้ำด้วย PMO) และโฟลเดอร์ `X` ที่มีอยู่และเขียนได้ When เรียก service ตั้งโฟลเดอร์ของ P เป็น `X` Then `appdb.get_project(P.id)` มีค่า `X` และ `project_config` ของ P ให้ `attachment_dir == X`
- **AC-7** (ไม่มีสิทธิ์) Given ผู้ใช้เป็น QA หรือ User BU/PM ของ P หรือไม่ได้เป็นสมาชิก P (แม้จะเป็น Admin ของโปรเจกต์อื่น Q) When ตั้งโฟลเดอร์ของ P Then raise `ProjectError` ที่บอกว่าไม่มีสิทธิ์, ค่าไม่เปลี่ยน และไม่มี event ใหม่
- **AC-8** (ไม่มีโฟลเดอร์ / ไม่ใช่โฟลเดอร์) Given path ที่ไม่มีอยู่ หรือเป็น path ของไฟล์ When ตั้งค่า Then raise `ProjectError`, ค่าไม่เปลี่ยน, ไม่มี event และ path นั้นยังไม่ถูกสร้าง
- **AC-9** (เขียนไม่ได้) Given โฟลเดอร์ที่มีอยู่แต่การสร้างไฟล์ทดสอบ raise `PermissionError` (monkeypatch) When ตั้งค่า Then raise `ProjectError` ที่บอกว่าเขียนไม่ได้, ค่าไม่เปลี่ยน, ไม่มี event
- **AC-10** (ไม่ทิ้งไฟล์ทดสอบ) Given ตั้งค่าสำเร็จ (AC-6) Then ในโฟลเดอร์ `X` ไม่มีไฟล์ใหม่เหลืออยู่ (รายชื่อไฟล์ก่อน/หลังเท่ากัน)
- **AC-11** (relative path) Given path แบบ relative (เช่น `"attachments"`) When ตั้งค่า Then raise `ProjectError` และค่าไม่เปลี่ยน
- **AC-12** (อยู่ในข้อมูลโปรเจกต์) Given path ที่อยู่ใต้ `<root>/projects/` (ทั้งของ P เอง และของโปรเจกต์อื่น) When ตั้งค่า Then raise `ProjectError` และค่าไม่เปลี่ยน
- **AC-13** (ตัดช่องว่าง / quote) Given path `'  "X"  '` (มีช่องว่างและ `"` ครอบ) When ตั้งค่า Then ค่าที่บันทึกคือ `X`
- **AC-14** (กลับค่าเริ่มต้น) Given P ตั้งโฟลเดอร์ `X` ไว้ When เรียก service ด้วย `None` (และด้วย `""`) Then ค่าใน `app.db` เป็น NULL/ว่าง และ `project_config` ของ P ให้ `<root>/projects/<P.id>/attachments` แม้ `X` จะถูกลบไปแล้วก็กลับค่าเริ่มต้นได้
- **AC-15** (ค่าเดิม) Given P ตั้ง `X` ไว้แล้ว When ตั้ง `X` อีกครั้ง Then ไม่ error และจำนวน event `"attachment_dir"` ของ P ไม่เพิ่ม
- **AC-16** (ไม่มีโปรเจกต์) Given id ที่ไม่มีใน `app.db` When ตั้งค่า Then raise `ProjectError` "Unknown project"

### Event (F8)
- **AC-17** Given P ไม่ได้ตั้งค่า When ตั้งเป็น `X` แล้วเปลี่ยนเป็น `Y` แล้วกลับค่าเริ่มต้น Then `appdb.events(P.id)` มี event action `"attachment_dir"` 3 แถว, `editor` = ผู้ตั้ง, detail ตามลำดับคือ `"(default) -> X"`, `"X -> Y"`, `"Y -> (default)"`

### อัปโหลดไปโฟลเดอร์ที่ตั้ง (F2, F7)
- **AC-18** (ตำแหน่งไฟล์) Given P ตั้งโฟลเดอร์ `X` และ item `TP-1` ที่ active When `upload()` ไฟล์ `result.pdf` Then ไฟล์อยู่ที่ `X/<sanitized TP-1>/<stamp>_result.pdf`, `UploadResult.stored_path` และ `stored_path` ในแถว attachment เป็น path นั้น, SHA-256 ตรงกับต้นฉบับ, TC/Result Complete = Complete, มี audit `attachment` และมี pending write-back (เหมือนพฤติกรรมเดิม)
- **AC-19** (ไม่มีไฟล์ใน local) Given AC-18 Then ไม่มีไฟล์ใหม่ใต้ `<root>/projects/<P.id>/attachments`
- **AC-20** (ชื่อซ้ำ) Given อัปโหลดไฟล์ชื่อเดียวกันใน item เดียวกันด้วย `now` เดียวกัน 2 ครั้ง ไปที่ `X` Then ได้ 2 ไฟล์ (ไฟล์ที่สองมี `_2_`) และไฟล์แรกไม่ถูกเขียนทับ
- **AC-21** (ไม่ตั้งค่า = เดิม) Given P ไม่ได้ตั้งโฟลเดอร์ When `upload()` Then ไฟล์อยู่ที่ `<root>/projects/<P.id>/attachments/<item>/<stamp>_<name>` เหมือนก่อนฟีเจอร์นี้ (test เดิมของ attachments ผ่านโดยไม่ต้องแก้)

### โฟลเดอร์ใช้ไม่ได้ตอนอัปโหลด (F6)
- **AC-22** (โฟลเดอร์หาย) Given P ตั้งโฟลเดอร์ `X` แล้ว `X` ถูกลบ (จำลอง OneDrive ไม่ได้ mount) When `upload()` Then raise error ชนิดเฉพาะ (subclass ของ `UploadError`) ที่มี path `X`, `X` ยังไม่มีอยู่ (ระบบไม่สร้าง), ไม่มีแถว attachment ใหม่, ไม่มี audit ใหม่, TC/Result Complete ไม่เปลี่ยน, ไม่มี pending write-back และไม่มีไฟล์ใหม่ใต้ `<root>/projects/<P.id>/attachments` (ไม่ fallback)
- **AC-23** (copy ล้มเหลว) Given `X` มีอยู่แต่ `shutil.copy2` raise `OSError` (monkeypatch) When `upload()` Then raise error ชนิดเดียวกับ AC-22, ไม่มีแถว attachment / audit ใหม่ และไม่มีไฟล์ที่ copy ไม่ครบเหลืออยู่ใต้ `X`
- **AC-24** (สร้างโฟลเดอร์ย่อยไม่ได้) Given `X` มีอยู่แต่ `mkdir` ของโฟลเดอร์ item raise `PermissionError` When `upload()` Then raise error ชนิดเดียวกับ AC-22 และไม่มีอะไรถูกบันทึก
- **AC-25** (log) Given AC-22 Then มี log ระดับ WARNING ที่มี project id, item id และ path `X` (ตรวจด้วย `caplog`)
- **AC-26** (กลับมาใช้ได้) Given AC-22 แล้วสร้าง `X` กลับมา When `upload()` อีกครั้ง Then สำเร็จตาม AC-18

### ไฟล์เดิมไม่ย้าย (F9)
- **AC-27** Given P ไม่ได้ตั้งค่า อัปโหลดไฟล์ A แล้ว (อยู่ใน local) จากนั้นตั้งโฟลเดอร์ `X` และอัปโหลดไฟล์ B จากนั้นเปลี่ยนเป็น `Y` และอัปโหลดไฟล์ C Then ไฟล์ A อยู่ใน local, B อยู่ใต้ `X`, C อยู่ใต้ `Y`, ทั้งสามไฟล์ยังอยู่และ hash ตรงเดิม และ `stored_path` ของแถว A และ B ไม่เปลี่ยนจากตอนอัปโหลด
- **AC-28** Given AC-27 When กลับไปใช้ค่าเริ่มต้น Then ไฟล์ใน `X` และ `Y` ไม่ถูกย้ายหรือลบ และ `repo.attachments(TEST_PROGRESS)` ยังคืนครบ 3 แถว (item ยังนับว่าอัปโหลดแล้ว)

### Clear project data / Delete project (F12)
- **AC-29** (Clear ไม่แตะโฟลเดอร์ที่เลือก) Given P (เปิด `enable_clear_project`) ตั้งโฟลเดอร์ `X` ที่มีไฟล์ที่อัปโหลดแล้ว และมีไฟล์อื่นที่ไม่เกี่ยวกับแอปใน `X` When `clear_project` Then ไฟล์ทุกไฟล์ใน `X` ยังอยู่ครบ และ backup zip ไม่มีไฟล์จาก `X` (มีเฉพาะไฟล์จาก `<data_dir>/attachments`)
- **AC-30** (Restore ไม่เขียนลงโฟลเดอร์ที่เลือก) Given backup zip ของ Clear project data ที่มี entry `attachments/...` และ P ตั้งโฟลเดอร์ `X` When `restore_backup` Then ไฟล์ถูกเขียนกลับใต้ `<data_dir>/attachments` และไม่มีไฟล์ใหม่ใน `X`
- **AC-31** (Delete ไม่แตะโฟลเดอร์ที่เลือก) Given P (Archive แล้ว) ตั้งโฟลเดอร์ `X` ที่มีไฟล์ When ลบ P ด้วย Delete project Then ไฟล์ใน `X` ยังอยู่ครบ

### UI (F10)
- **AC-32** (ผู้มีสิทธิ์เห็นปุ่ม) Given ผู้ใช้เป็น Admin/PMO ของโปรเจกต์ปัจจุบัน When เปิดหน้า Settings Then การ์ดโปรเจกต์มีป้าย `t("proj.att_dir")`, ค่าปัจจุบัน และปุ่ม `t("proj.att_choose")` Given โปรเจกต์ไม่ได้ตั้งค่า Then แสดงข้อความ `t("proj.att_default", p=...)` ที่มี path `<data_dir>/attachments` และไม่มีปุ่ม (หรือปุ่ม disabled) `t("proj.att_reset")`
- **AC-33** (ผู้ไม่มีสิทธิ์) Given ผู้ใช้เป็น QA ของโปรเจกต์ปัจจุบัน When เปิดหน้า Settings Then เห็นค่าปัจจุบัน แต่ไม่มีปุ่มเลือกโฟลเดอร์ / ใช้ค่าเริ่มต้น (หรือ disabled)
- **AC-34** (เลือกสำเร็จ) Given Admin, `getExistingDirectory` คืน `X` ที่ถูกต้อง และ `confirm` ตอบ yes When กดปุ่มเลือกโฟลเดอร์ Then service ถูกเรียกด้วย `X`, แสดง `t("proj.att_saved")`, การ์ดแสดง `X` และ `self.ctx.cfg.attachment_dir == X` (การอัปโหลดถัดไปจากหน้า Test Progress ไปที่ `X` โดยไม่ต้องเปิดโปรเจกต์ใหม่)
- **AC-35** (ยกเลิก) Given `getExistingDirectory` คืน `""` หรือ `confirm` ตอบ no When กดปุ่ม Then service ไม่ถูกเรียกและค่าไม่เปลี่ยน
- **AC-36** (ไม่ผ่านการตรวจ) Given service raise `ProjectError` (เช่น โฟลเดอร์เขียนไม่ได้) When กดปุ่ม Then `error()` ถูกเรียกด้วยข้อความ (จาก `t()` ตามภาษาปัจจุบัน) และค่าบนการ์ดไม่เปลี่ยน
- **AC-37** (ใช้ค่าเริ่มต้น) Given โปรเจกต์ตั้งโฟลเดอร์ไว้ และ `confirm` ตอบ yes When กด `t("proj.att_reset")` Then service ถูกเรียกด้วย `None` และการ์ดแสดงข้อความค่าเริ่มต้น
- **AC-38** (อัปโหลดล้มเพราะโฟลเดอร์) Given หน้า Test Progress และ `upload()` raise error ชนิดของ AC-22 When กดอัปโหลด Then `error()` ถูกเรียกด้วยข้อความจาก `t("tp.folder_unavailable", p=X)` (มี path `X`) และไม่เรียก `info(t("tp.uploaded"))`

### i18n (F11)
- **AC-39** Given ทุก key ใหม่ในตาราง F11 (และ key อื่นที่เพิ่มในฟีเจอร์นี้) When ตรวจ `ui/i18n.py` Then มีทั้ง EN และ TH ที่ไม่ว่าง และ placeholder (`{p}`, `{n}`) ตรงกันทั้งสองภาษา
- **AC-40** Given ภาษา TH When render ข้อความ AC-32, AC-36, AC-38 Then เป็นข้อความ TH (ไม่ใช่ EN)

### ทั่วไป
- **AC-41** `pytest -q` ผ่านทั้งหมด test เดิมของ attachments, projects, clear_project และ delete_project ผ่านโดยไม่ต้องแก้ ยกเว้นจุดที่ต้องปรับเพราะ migration ใหม่ (เช่น ค่าที่คาดของ `user_version`) ต้องรายงานหัวหน้าว่าแก้ test ใด ห้ามลบ test ทิ้ง

## ข้อจำกัด / สมมติฐาน
- Python 3.x, PySide6, SQLite, PEP 8, type hints, pytest
- `app.db` อยู่บนเครื่องแต่ละเครื่อง (`%LOCALAPPDATA%\PlanningProgress\app.db`) ค่าโฟลเดอร์จึงเป็นของเครื่องนั้น ถ้าทีมใช้แอปหลายเครื่อง แต่ละเครื่องต้องตั้งเอง และ path ของ OneDrive อาจต่างกันในแต่ละเครื่อง (มีชื่อ user อยู่ใน path) ดูคำถาม Q1
- `stored_path` ที่บันทึกเป็น path เต็มของเครื่องที่อัปโหลด เครื่องอื่นอาจเปิด path นั้นไม่ได้ถ้า OneDrive ของเขาอยู่คนละ path (ไม่แก้ในรอบนี้)
- OneDrive Files On-Demand: ไฟล์ที่เพิ่ง copy อาจยังไม่ขึ้น cloud ตอนที่แอปแจ้งสำเร็จ แอปรับประกันแค่ว่าไฟล์ถูกเขียนลงโฟลเดอร์ local ที่ sync และ hash ตรง
- ถ้าหลายโปรเจกต์ตั้งโฟลเดอร์เดียวกัน ไฟล์จะอยู่ในโครงสร้างเดียวกัน (`<item_id>/...`) ชื่อไฟล์ไม่ทับกันเพราะมี stamp และเลขต่อท้าย แต่ item id ของต่างโปรเจกต์อาจเป็นโฟลเดอร์เดียวกัน (ดู Q2)
- การตรวจ "เขียนได้" ทำตอนตั้งค่าเท่านั้น ตอนอัปโหลดตรวจแค่ว่ามีโฟลเดอร์ แล้วอาศัย error จริงจากการ copy (F6)
- รายละเอียดทางเทคนิค (ชื่อ column, ชื่อฟังก์ชัน, ชื่อ exception, วิธีให้ `Config` รู้ค่าใหม่ เช่น field ใหม่ใน `Config` ที่ `attachment_dir` property ใช้, วิธี map error เป็นข้อความแปล) เป็นหน้าที่ architect ตราบที่ผ่าน AC

### สรุปสมมติฐานที่รอผู้ใช้ยืนยัน
- **B1** เก็บค่าเป็น column ใหม่ในตาราง `project` ของ `app.db` (F1)
- **B2** ผู้ตั้งได้คือ Admin/PMO **ของโปรเจกต์นั้น** เท่านั้น (F3)
- **B3** ไม่ยอมรับโฟลเดอร์ใต้ `<root>/projects/` (F5 ข้อ 5)
- **B4** ระบบไม่สร้างโฟลเดอร์ปลายทางให้เอง ทั้งตอนตั้งและตอนอัปโหลด (F5 ข้อ 3, F6 ข้อ 1)
- **B5** โฟลเดอร์ใช้ไม่ได้ตอนอัปโหลด -> แจ้ง error และไม่บันทึกอะไร ไม่ fallback ไป local (F6 ข้อ 4)
- **B6** โครงสร้างโฟลเดอร์ย่อยเหมือนเดิม `<folder>/<item_id>/<stamp>_<name>` ไม่มีชั้นชื่อโปรเจกต์ (F7)
- **B7** ไฟล์เดิมไม่ย้าย `stored_path` เดิมไม่แก้ (F9, ตามที่หัวหน้าแนะนำ)
- **B8** Clear project data / restore / Delete project ไม่แตะโฟลเดอร์ที่ผู้ใช้เลือก (F12)
- **B9** บันทึกค่าทันทีเมื่อเลือกโฟลเดอร์ (ไม่ผูกกับปุ่ม Save) และ audit เป็น `project_event` ใน `app.db` เท่านั้น (F8, F10)

## คำถามที่ยังไม่มีคำตอบ
- **Q1** แอปใช้หลายเครื่องกับโปรเจกต์เดียวกันหรือไม่ ถ้าใช่ ยอมรับได้ไหมที่แต่ละเครื่องต้องตั้งโฟลเดอร์เอง (ค่าอยู่ใน `app.db` ของเครื่อง) หรือต้องการให้รองรับตัวแปรอย่าง `%OneDrive%` / `%OneDriveCommercial%` ใน path เพื่อให้ใช้ค่าเดียวกันได้ทุกเครื่อง (ถ้าต้องการ = เพิ่ม scope)
- **Q2** ต้องการให้มีโฟลเดอร์ย่อยชื่อโปรเจกต์ (เช่น `<folder>/<ชื่อโปรเจกต์>/<item_id>/...`) หรือไม่ กรณีหลายโปรเจกต์ใช้โฟลเดอร์ทีมเดียวกัน และต้องการให้โฟลเดอร์ย่อยระดับ item ใช้ชื่อแอป (Application name) แทน item id เพื่อให้คนในทีมหาไฟล์ง่ายขึ้นหรือไม่
- **Q3** สิทธิ์: ให้เป็น Admin/PMO ของโปรเจกต์นั้นเท่านั้น (ข้อเสนอ) หรือใช้กติกาเดียวกับ Relink / Archive (Admin/PMO ของโปรเจกต์ใดก็ได้) หรือรวม User BU/PM ด้วยแบบ `set_window`
- **Q4** เมื่อโฟลเดอร์ใช้ไม่ได้ตอนอัปโหลด ยืนยันว่าให้ block (ข้อเสนอ) ไม่ใช่ fallback เก็บ local ชั่วคราว
- **Q5** ต้องการคำสั่ง CLI (เช่น `project attachment-dir <project> <folder>`) ด้วยหรือไม่ (ตอนนี้ไม่อยู่ใน scope)
- **Q6** ต้องการให้ข้อความสำเร็จหลังอัปโหลด (`tp.uploaded`) บอกตำแหน่งที่เก็บไฟล์ด้วยหรือไม่ และต้องการให้ audit ของการอัปโหลดใน tracker.db ระบุโฟลเดอร์ปลายทางเพิ่มหรือไม่ (ตอนนี้มีใน `stored_path` ของแถว attachment อยู่แล้ว)
- **Q7** การเปลี่ยนโฟลเดอร์ควรบันทึก audit ใน tracker.db ของโปรเจกต์ด้วยหรือไม่ (ข้อเสนอ: `project_event` ใน `app.db` อย่างเดียว เหมือน Relink / Date window)

## การเปลี่ยนแปลงตามคำขอผู้ใช้ (2026-11-25, ภายหลัง) — แทน F7, AC-18, AC-20, AC-21, AC-38 ในส่วนที่ขัดกัน
- **โฟลเดอร์ย่อยใช้ชื่อแอป** โครงสร้างใหม่ `<โฟลเดอร์ไฟล์แนบ>/<App name>/<ชื่อไฟล์เดิม>` ทั้งโฟลเดอร์ที่เลือกและ `<data_dir>/attachments` (App name = บรรทัดแรกของคอลัมน์ App ใน Test Progress) ทำให้ปลอดภัยบน Windows: แทน `<>:"/\|?*` และ control character ด้วย `_`, ตัดช่องว่างนำหน้าและจุด/ช่องว่างท้าย, ยาวไม่เกิน 100 ตัวอักษร, ชื่อสงวน (CON, PRN, AUX, NUL, COM1-9, LPT1-9) เติม `_` ท้าย, คงภาษาไทยและช่องว่าง ถ้า App name ว่างใช้ item id แทน
- **ชื่อไฟล์เดิม** ไม่มี stamp นำหน้า ถ้าชื่อซ้ำในโฟลเดอร์เดียวกันบันทึกเป็น `name (2).ext`, `name (3).ext` ... ไม่เขียนทับไฟล์เดิมไม่ว่ากรณีใด (จองชื่อด้วยการสร้างไฟล์แบบ exclusive)
- **อัปโหลดหลายไฟล์พร้อมกัน** ปุ่มด้านบนและปุ่ม Upload ของแถวใช้ `QFileDialog.getOpenFileNames` แต่ละไฟล์ตรวจ (ชนิด, ขนาด, ไฟล์ว่าง, hash) และบันทึกแยกกันผ่าน `upload_many()` แต่ละไฟล์มีแถว attachment และ audit ของตัวเอง TC/Result Complete ถูกตั้งครั้งเดียวโดยไฟล์แรกที่สำเร็จ ไฟล์ที่ล้มไม่หยุดไฟล์อื่น จบแล้วแสดงข้อความสรุปเดียว: ไฟล์ที่สำเร็จ (และชื่อที่บันทึกจริงถ้าถูกเปลี่ยน), ไฟล์ที่ไม่สำเร็จพร้อมเหตุผล และโฟลเดอร์ปลายทาง (ถ้าโฟลเดอร์ที่เลือกใช้ไม่ได้ แสดงข้อความ `tp.folder_unavailable` แทน) สำเร็จอย่างน้อย 1 ไฟล์ = `info()`, ไม่สำเร็จทั้งหมด = `error()` i18n ใหม่: `tp.up_ok`, `tp.up_failed`, `tp.up_saved_as`, `tp.up_folder_short`, `tp.up_dest`, `tp.up_complete`
- **วางลิงก์ SharePoint (ผู้ใช้อนุมัติ 2026-11-25)** แถว "โฟลเดอร์เก็บไฟล์ผลทดสอบ" มีปุ่ม "วางลิงก์ SharePoint…" เพิ่ม (สิทธิ์เดียวกับ เลือกโฟลเดอร์…) แอปแปลงลิงก์เป็นโฟลเดอร์ในเครื่องที่ OneDrive client sync อยู่ (อ่าน registry `HKCU\Software\SyncEngines\Providers\OneDrive\<id>`: `UrlNamespace`, `MountPoint`; library ที่ `UrlNamespace` เป็น prefix ยาวที่สุดชนะ) แล้ว copy ไฟล์ไปที่นั่นเหมือนเดิม OneDrive อัปโหลดขึ้น SharePoint เอง ไม่ใช้ Graph / MSAL / login ไม่มี dependency ใหม่ รับลิงก์แบบ path ตรง, แบบ address bar (`?id=` หรือ `RootFolder=`) และ `/:f:/r/<path>` ลิงก์แชร์แบบ opaque (`/:f:/s/...`) ไม่รับ แจ้งให้ copy URL จาก address bar แทน ถ้าไม่มี library ที่ sync ตรงกับลิงก์ แจ้งวิธีแก้ (กด Sync หรือ "Add shortcut to My files" ใน SharePoint รอ sync เสร็จ แล้วลองใหม่ กรณี shortcut ให้ใช้ เลือกโฟลเดอร์…) โฟลเดอร์ที่ได้ต้องผ่านการตรวจ F5 เดิม เก็บลิงก์ใน column ใหม่ `project.attachment_url` (migration `004_attachment_url.sql`) คู่กับ `attachment_dir` (path ในเครื่อง ที่การอัปโหลดใช้) "เลือกโฟลเดอร์…" และ "ใช้ค่าเริ่มต้น" ล้างลิงก์ event `attachment_dir` มีลิงก์ใน detail การ์ดแสดงลิงก์และ path ในเครื่อง ข้อความสรุปการอัปโหลดบอกลิงก์ SharePoint ด้วย (`tp.up_dest_sp`) i18n ใหม่: `proj.att_link*`

---

# 2026-11-25 — ดาวน์โหลดฟอร์ม Excel Progress เปล่า (Download blank form)

> requirement ผู้ใช้: "อยากให้สามารถ ดาวน์โหลด form excel progress" ผู้ใช้ตอบคำถามเพิ่มแล้วว่า
> (1) ต้องการ **ฟอร์มเปล่า (template)** มีแค่ header + formatting ไม่มีข้อมูล เพื่อส่งให้คนอื่นกรอก
> (2) ครบ **3 sheet** (UAT Plan, Test Progress, Defect List) ในไฟล์ .xlsx เดียว
> (3) **ทุก role** ดาวน์โหลดได้
> ข้อที่ PM เสนอค่าเริ่มต้นไว้ทำเครื่องหมาย **(ข้อเสนอ — รอยืนยัน)** และสรุปไว้ท้าย section (C1 ถึง C5)

## เป้าหมาย
ตอนนี้ฟอร์มเปล่าถูกสร้างได้เฉพาะตอนสร้างโปรเจกต์ใหม่ (`services/projects.py` เรียก `services/project_template.py` `create_workbook(path)`) ผู้ใช้ไม่มีทางเอาฟอร์มเปล่าไปให้คนอื่นกรอก นอกจาก copy ไฟล์ของโปรเจกต์แล้วลบข้อมูลเอง (เสี่ยงข้อมูลหลุดและ format เพี้ยน) รอบนี้เพิ่มปุ่มดาวน์โหลดฟอร์มเปล่าที่หน้าตาเหมือนไฟล์โปรเจกต์ใหม่ทุกประการ

ผู้ได้ประโยชน์: ผู้ใช้ทุก role (Admin, PMO, QA, User BU/PM) ที่ต้องส่งฟอร์มให้ทีม / vendor / BU กรอกข้อมูล

## ขอบเขต (In scope)
- **D1 เนื้อหาไฟล์** ไฟล์ที่ได้มีเนื้อหาเหมือนไฟล์ที่ `create_workbook()` สร้างทุกประการ (ใช้ `resources/project_template.json` ชุดเดียวกัน ห้ามสร้าง template แยก): 3 sheet ตามลำดับ `SHEET_ORDER`, header row, style, column width / hidden column (รวม App ID ที่ซ่อนอยู่), freeze header, drop-down (data validation), conditional formatting, legend ของ Test Progress, คอลัมน์ Report By Email / CC Email **ไม่มีข้อมูลจากโปรเจกต์ใด ๆ** (ไม่มีแถวข้อมูล ไม่มีชื่อโปรเจกต์ ไม่มี email ผู้ใช้)
- **D2 service** ฟังก์ชันใหม่ (เช่น `export_blank_form(dest: Path, protected: Iterable[Path]) -> Path` ชื่อและที่อยู่จริง architect เลือก) ที่ reuse ตรรกะการสร้าง workbook ของ `create_workbook()` (เช่นแยกส่วนสร้าง `Workbook` ออกเป็นฟังก์ชันภายใน) โดย**พฤติกรรมของ `create_workbook()` เดิมต้องไม่เปลี่ยน** (ยัง raise `FileExistsError` ถ้ามีไฟล์อยู่แล้ว)
- **D3 เขียนทับไฟล์ที่ผู้ใช้เลือก** ถ้าไฟล์ปลายทางมีอยู่แล้ว (ผู้ใช้ยืนยันการเขียนทับใน save dialog แล้ว) ให้เขียนทับได้ ไม่ raise `FileExistsError` วิธีเขียน: save ลงไฟล์ชั่วคราวในโฟลเดอร์เดียวกัน แล้ว `replace` ไปที่ปลายทาง ถ้าล้มเหลว ณ จุดใด ไฟล์ปลายทางเดิม (ถ้ามี) ต้องไม่เสีย และไม่มีไฟล์ชั่วคราวเหลือ
- **D4 ห้ามแตะไฟล์ Excel ของโปรเจกต์** ถ้า path ที่เลือก (หลัง normalize: `resolve()` + เทียบแบบไม่สนตัวพิมพ์บน Windows) ตรงกับ `cfg.excel_path` ของโปรเจกต์ปัจจุบัน -> ปฏิเสธ (raise error เฉพาะ) ไม่เขียนอะไรเลย ไฟล์เดิมไม่เปลี่ยนแม้แต่ byte เดียว **(ข้อเสนอ — รอยืนยัน: ป้องกันเฉพาะไฟล์ของโปรเจกต์ปัจจุบัน ดู Q2)**
- **D5 ไม่แตะฐานข้อมูล** การดาวน์โหลดไม่เขียน tracker.db / app.db, ไม่มี audit, ไม่มี pending write-back, ไม่เปิด Excel ผ่าน COM
- **D6 สิทธิ์** ทุก role ใช้ได้ (ไม่ตรวจ `auth`) และใช้ได้แม้ผู้ใช้ไม่มี role ในโปรเจกต์ (ฟอร์มไม่มีข้อมูล)
- **D7 UI** หน้า Test Progress (`ui/screens/test_progress.py`) เพิ่มปุ่ม "ดาวน์โหลดฟอร์มเปล่า / Download blank form" ใน top bar (ก่อนปุ่ม Upload) แสดงทุก role **(ข้อเสนอ — รอยืนยัน: ตำแหน่งปุ่ม)**
  - กดแล้วเปิด `QFileDialog.getSaveFileName` ชื่อเริ่มต้น `Planning Progress Form.xlsx` filter `Excel (*.xlsx)` (แบบเดียวกับ Export ของหน้า History) **(ข้อเสนอ — รอยืนยัน: ชื่อไฟล์)**
  - ยกเลิก dialog (path ว่าง) -> ไม่ทำอะไร
  - ถ้าผู้ใช้พิมพ์ชื่อไม่มีนามสกุล `.xlsx` ให้เติม `.xlsx` ให้
  - สำเร็จ -> `info()` ข้อความแปลที่มี path เต็มของไฟล์ที่บันทึก
  - ล้มเหลว -> `error()` ข้อความแปล: กรณีเป็นไฟล์ Excel ของโปรเจกต์ (D4) และกรณีบันทึกไม่ได้ (`OSError` เช่น `PermissionError` เพราะไฟล์เปิดอยู่ใน Excel / ไม่มีสิทธิ์เขียนโฟลเดอร์) ข้อความกรณีหลังแนะนำให้ปิดไฟล์ใน Excel หรือเลือกที่อื่น แล้วลองใหม่
  - งานสร้างไฟล์ใช้เวลาไม่นาน (< 1–2 วินาที) ไม่ต้องใช้ worker thread **(สมมติฐาน)**
- **D8 i18n** ข้อความใหม่อยู่ใน `ui/i18n.py` เป็น tuple (EN, TH) ข้อเสนอ (ปรับถ้อยคำได้ตราบที่ความหมายตรง):

  | key | EN | TH |
  |---|---|---|
  | `tp.form` | Download blank form | ดาวน์โหลดฟอร์มเปล่า |
  | `tp.form_saved` | Blank form saved:\n{p} | บันทึกฟอร์มเปล่าแล้ว:\n{p} |
  | `tp.form_is_project` | This is the project's Excel file. Choose another file name. | ไฟล์นี้คือไฟล์ Excel ของโปรเจกต์ กรุณาเลือกชื่อไฟล์อื่น |
  | `tp.form_failed` | Can't save the form:\n{p}\nClose the file if it is open in Excel, or choose another folder, then try again. | บันทึกฟอร์มไม่ได้:\n{p}\nกรุณาปิดไฟล์ถ้าเปิดอยู่ใน Excel หรือเลือกโฟลเดอร์อื่น แล้วลองใหม่ |

## นอกขอบเขต (Out of scope)
- การ export ข้อมูลจริงของโปรเจกต์ (ทั้งหมดหรือบางส่วน)
- การเลือกเฉพาะบาง sheet หรือบางคอลัมน์
- ปุ่มดาวน์โหลดในหน้าอื่น (UAT Plan, Defect List, Settings ฯลฯ) และคำสั่ง CLI
- การแก้ template (`project_template.json`, `tools/build_template.py`) หรือหน้าตาฟอร์ม
- การนำเข้า (import) ฟอร์มที่คนอื่นกรอกกลับเข้าระบบ
- การใส่ชื่อโปรเจกต์ / วันที่ / drop-down รายชื่อแอปของโปรเจกต์ลงในฟอร์ม

## User stories
- **US-1** ในฐานะผู้ใช้ทุก role ฉันต้องการดาวน์โหลดฟอร์ม Excel Progress เปล่าที่มีครบ 3 sheet เพื่อส่งให้คนอื่นกรอกข้อมูลในรูปแบบเดียวกับไฟล์ของระบบ
- **US-2** ในฐานะผู้ใช้ ฉันต้องการเลือกที่เก็บและชื่อไฟล์เอง และรู้ว่าไฟล์ถูกบันทึกที่ไหน เพื่อหาไฟล์เจอ
- **US-3** ในฐานะ PMO ฉันต้องการมั่นใจว่าการดาวน์โหลดฟอร์มจะไม่เขียนทับไฟล์ Excel ของโปรเจกต์โดยไม่ตั้งใจ เพื่อไม่ให้ข้อมูลหาย

## Acceptance criteria
ทุกข้อทดสอบด้วย pytest ใช้ `tmp_path`, ตรวจไฟล์ด้วย `openpyxl.load_workbook` ไม่ใช้ Excel / COM ส่วน UI ใช้ `QApplication` แบบ offscreen ตามแนว `tests/test_ui.py` โดย monkeypatch `QFileDialog.getSaveFileName`, `info`, `error` กรณีไฟล์เปิดอยู่ / ไม่มีสิทธิ์จำลองด้วย monkeypatch ให้ raise `PermissionError` (ไม่พึ่ง file lock ของ Windows)

### เนื้อหาไฟล์ (D1, D2)
- **AC-1** (ตรงกับ template) Given `dest` ที่ยังไม่มีไฟล์ When เรียก service Then ไฟล์ถูกสร้าง, `wb.sheetnames == list(SHEET_ORDER)`, และ header row ของแต่ละ sheet เท่ากับ `project_template.headers(sheet)`
- **AC-2** (เหมือน create_workbook) Given ไฟล์ A จาก service และไฟล์ B จาก `create_workbook()` When เปรียบเทียบทุก sheet Then ค่าในทุก cell, column width, column ที่ซ่อน (รวม App ID), `freeze_panes`, จำนวน/range ของ data validation และ conditional formatting, และ legend ของ Test Progress เหมือนกัน
- **AC-3** (ไม่มีข้อมูล) Given ไฟล์จาก AC-1 Then ทุก sheet ไม่มีค่าในแถวที่ 2 ขึ้นไปของคอลัมน์ข้อมูล (นับเฉพาะคอลัมน์ในตาราง ไม่รวม legend)
- **AC-4** (create_workbook เดิมไม่เปลี่ยน) Given ไฟล์ที่มีอยู่แล้ว When เรียก `create_workbook(path)` Then ยัง raise `FileExistsError` และไฟล์เดิมไม่เปลี่ยน (test เดิมของ project_template / projects ผ่านโดยไม่ต้องแก้)

### เขียนทับและความปลอดภัยของไฟล์ (D3, D4, D5)
- **AC-5** (เขียนทับ) Given `dest` มีไฟล์อื่นอยู่แล้ว (เช่น ไฟล์ข้อความ) When เรียก service Then ไม่ error, `dest` เป็นฟอร์มเปล่าที่ผ่าน AC-1 และไม่มีไฟล์ชั่วคราวเหลือในโฟลเดอร์
- **AC-6** (บันทึกล้มเหลว) Given `dest` มีไฟล์เดิมอยู่ และการ `replace` (หรือ save) raise `PermissionError` (monkeypatch) When เรียก service Then error ถูกส่งออก (เป็น `OSError` หรือ error เฉพาะที่ architect กำหนด), ไฟล์เดิมที่ `dest` มีเนื้อหาเหมือนก่อนเรียก และไม่มีไฟล์ชั่วคราวเหลือ
- **AC-7** (ไฟล์ของโปรเจกต์) Given `dest` เท่ากับ `cfg.excel_path` ของโปรเจกต์ (ทดสอบทั้ง path เดียวกันตรง ๆ, ต่างตัวพิมพ์ใหญ่เล็ก และแบบ relative / มี `..` ที่ชี้ไฟล์เดียวกัน) When เรียก service Then raise error เฉพาะ, bytes และ mtime ของไฟล์โปรเจกต์ไม่เปลี่ยน และไม่มีไฟล์ชั่วคราวเกิดขึ้น
- **AC-8** (ไม่แตะ DB) Given โปรเจกต์ที่มี tracker.db When ดาวน์โหลดฟอร์ม Then จำนวนแถว audit และ pending write-back ไม่เปลี่ยน

### UI (D6, D7)
- **AC-9** (ปุ่มแสดงทุก role) Given ผู้ใช้ role Admin, PMO, QA, User BU/PM (parametrize) When เปิดหน้า Test Progress Then ปุ่มข้อความ `t("tp.form")` มองเห็นได้และ enabled
- **AC-10** (dialog) Given กดปุ่ม When `getSaveFileName` ถูกเรียก Then ชื่อเริ่มต้นคือ `Planning Progress Form.xlsx` และ filter มี `*.xlsx`
- **AC-11** (สำเร็จ) Given `getSaveFileName` คืน `<tmp>/form.xlsx` When กดปุ่ม Then ไฟล์ถูกสร้างและผ่าน AC-1 และ `info()` ถูกเรียกด้วยข้อความ `t("tp.form_saved", p=...)` ที่มี path เต็ม
- **AC-12** (เติมนามสกุล) Given `getSaveFileName` คืน `<tmp>/form` (ไม่มีนามสกุล) When กดปุ่ม Then ไฟล์ที่ได้คือ `<tmp>/form.xlsx`
- **AC-13** (ยกเลิก) Given `getSaveFileName` คืน `""` When กดปุ่ม Then service ไม่ถูกเรียก และ `info()` / `error()` ไม่ถูกเรียก
- **AC-14** (ไฟล์ของโปรเจกต์) Given `getSaveFileName` คืน `cfg.excel_path` When กดปุ่ม Then `error()` ถูกเรียกด้วย `t("tp.form_is_project")`, `info()` ไม่ถูกเรียก และไฟล์โปรเจกต์ไม่เปลี่ยน
- **AC-15** (บันทึกไม่ได้) Given service raise `PermissionError` (monkeypatch) When กดปุ่ม Then `error()` ถูกเรียกด้วย `t("tp.form_failed", p=...)` ที่มี path และแอปไม่ crash

### i18n (D8)
- **AC-16** Given ทุก key ใหม่ในตาราง D8 When ตรวจ `ui/i18n.py` Then มีทั้ง EN และ TH ที่ไม่ว่าง และ placeholder (`{p}`) ตรงกันทั้งสองภาษา และเมื่อภาษาเป็น TH ข้อความปุ่ม / ข้อความใน AC-11, AC-14 เป็นภาษาไทย

### ทั่วไป
- **AC-17** `pytest -q` ผ่านทั้งหมด ห้ามลบหรือแก้ test เดิม (ถ้าจำเป็นต้องแก้ต้องรายงานหัวหน้า)

## ข้อจำกัด / สมมติฐาน
- Python 3.x, PySide6, openpyxl, PEP 8, type hints, pytest ไม่มี dependency ใหม่
- ฟอร์มสร้างจาก template ในโค้ด (`project_template.json`) ไม่ใช่จากไฟล์ Excel ของโปรเจกต์ ถ้าไฟล์ของโปรเจกต์เคยถูกปรับหน้าตาเอง ฟอร์มจะไม่ตามการปรับนั้น
- ฟอร์มเป็นแบบเดียวกันทุกโปรเจกต์ (ไม่มีข้อมูลเฉพาะโปรเจกต์)
- การยืนยันเขียนทับเป็นหน้าที่ของ `QFileDialog` (ค่าเริ่มต้นของ Qt ถามยืนยันอยู่แล้ว) service ไม่ถามซ้ำ
- ชื่อฟังก์ชัน, ที่อยู่ของฟังก์ชัน, ชื่อ exception, ชื่อไฟล์ชั่วคราว และวิธี normalize path เป็นหน้าที่ architect ตราบที่ผ่าน AC

### สรุปข้อเสนอที่รอผู้ใช้ยืนยัน
- **C1** ปุ่มอยู่ที่ top bar หน้า Test Progress ข้อความ "ดาวน์โหลดฟอร์มเปล่า / Download blank form" (D7)
- **C2** ชื่อไฟล์เริ่มต้น `Planning Progress Form.xlsx` (D7)
- **C3** เขียนทับไฟล์ที่ผู้ใช้เลือกได้ (หลังยืนยันใน dialog) แบบปลอดภัยผ่านไฟล์ชั่วคราว (D3)
- **C4** ปฏิเสธถ้าเลือก path เดียวกับไฟล์ Excel ของโปรเจกต์ปัจจุบัน (D4)
- **C5** หลังบันทึกแสดงข้อความพร้อม path (ไม่เปิดไฟล์หรือเปิดโฟลเดอร์ให้อัตโนมัติ) (D7)

## คำถามที่ยังไม่มีคำตอบ
- **Q1** ตำแหน่งปุ่ม: top bar หน้า Test Progress (ข้อเสนอ) พอไหม หรือต้องการที่อื่น (เช่น หน้า Settings หรือเมนูหลัก)
- **Q2** ต้องป้องกันการเขียนทับไฟล์ Excel ของ**ทุกโปรเจกต์**ในเครื่อง (รวมโปรเจกต์ที่ไม่ได้เปิดอยู่ / Archive) ด้วยหรือไม่ หรือเฉพาะโปรเจกต์ปัจจุบัน (ข้อเสนอ)
- **Q3** ชื่อไฟล์เริ่มต้นใช้ `Planning Progress Form.xlsx` (ข้อเสนอ) หรือต้องการใส่ชื่อโปรเจกต์ / วันที่ในชื่อ
- **Q4** หลังบันทึกสำเร็จต้องการให้เปิดไฟล์หรือเปิดโฟลเดอร์ให้อัตโนมัติหรือไม่ (ข้อเสนอ: ไม่ แสดงแค่ข้อความพร้อม path)
- **Q5** ในฟอร์มเปล่า คอลัมน์ App ID ที่ซ่อนอยู่ และคอลัมน์ Report By Email / CC Email ควรคงไว้เหมือนไฟล์โปรเจกต์ (ข้อเสนอ ตามที่ `create_workbook()` สร้าง) หรือควรตัดออกเพราะคนกรอกภายนอกไม่ต้องใช้

## การตัดสินใจของผู้ใช้ (25-Nov-2026)
ผู้ใช้อนุมัติ spec และเลือกข้อเสนอทั้งหมด: Q1 ปุ่มที่ top bar หน้า Test Progress / Q2 กันเฉพาะไฟล์ Excel ของโปรเจกต์ปัจจุบัน / Q3 ชื่อเริ่มต้น `Planning Progress Form.xlsx` / Q4 ไม่เปิดไฟล์หรือโฟลเดอร์ แสดงแค่ข้อความพร้อม path / Q5 คง App ID (ซ่อน) และ Report By Email / CC Email ไว้
รอบนี้ไม่เขียน test ใหม่ (ผู้ใช้สั่ง "ไม่ต้องเทส")

---

# 2026-11-26 — นำเข้า / เปลี่ยนไฟล์ Excel ของโปรเจกต์จากลิงก์ SharePoint (Import Excel from SharePoint)

> requirement ผู้ใช้: "ต้องให้สามารถ Import Excel จาก share point ได้"
> การตัดสินใจที่ผู้ใช้ให้มาแล้ว:
> (1) **วิธีเข้าถึง: ผ่าน OneDrive sync** แบบเดียวกับฟีเจอร์ "วางลิงก์ SharePoint" ของโฟลเดอร์ไฟล์แนบที่อนุมัติแล้ว (`services/sharepoint_link.py`): แปลงลิงก์เป็น path ในเครื่องที่ OneDrive client sync อยู่ (อ่าน registry `HKCU\Software\SyncEngines\Providers\OneDrive`) ไม่ใช้ Graph / MSAL / login จากนั้นใช้ไฟล์ `.xlsx` ในเครื่องนั้นเหมือนไฟล์ local ทุกประการ (การอ่าน import และการเขียนกลับ Excel แบบเดิมยังทำงาน)
> (2) **ขอบเขต: ทั้งสองทาง** (a) นำเข้าโปรเจกต์ใหม่จากลิงก์ไฟล์ SharePoint ที่วาง (คู่กับปุ่ม "นำเข้าไฟล์ Excel" เดิม) และ (b) เปลี่ยนไฟล์ Excel ของโปรเจกต์ที่มีอยู่เป็นไฟล์จากลิงก์ SharePoint (คู่กับ Relink เดิม)
> ข้อที่ PM เสนอค่าเริ่มต้นไว้ทำเครื่องหมาย **(ข้อเสนอ — รอยืนยัน)** และสรุปไว้ท้าย section (E1 ถึง E6)

## เป้าหมาย
ตอนนี้การนำเข้าโปรเจกต์ (`ui/projects.py` `import_project_flow` -> `services/projects.py` `import_project`) และการเปลี่ยนไฟล์ (`relink_flow` -> `relink`) รับได้แค่การเลือกไฟล์ผ่าน `QFileDialog` ผู้ใช้ที่มีแค่ลิงก์ของไฟล์ Excel บน SharePoint (ได้มาจากเมล / Teams) ต้องไปหาเองว่าไฟล์นั้นอยู่ที่ไหนในโฟลเดอร์ OneDrive ของเครื่อง ซึ่งยากและเลือกผิดไฟล์ได้ง่าย รอบนี้ให้วางลิงก์ไฟล์ SharePoint ได้โดยตรง แอปหาไฟล์ที่ OneDrive sync ไว้ในเครื่องให้ แล้วใช้ขั้นตอนนำเข้า / Relink เดิมทั้งหมด

ผู้ได้ประโยชน์: Admin / PMO ที่สร้าง (นำเข้า) โปรเจกต์ หรือย้ายโปรเจกต์ไปใช้ไฟล์ Excel บน SharePoint ของทีม

## ขอบเขต (In scope)
- **E-1 การแปลงลิงก์ไฟล์เป็น path ในเครื่อง** ฟังก์ชันใหม่ (เช่น `resolve_file_link(link) -> Path` ใน `services/sharepoint_link.py` หรือ `services/projects.py` ชื่อจริง architect เลือก) ที่ใช้กลไกเดิมของ `sharepoint_link.py` (`clean_link`, `link_location`, การจับคู่ library ด้วย `UrlNamespace` ที่เป็น prefix ยาวที่สุด, `MountPoint` + ส่วนที่เหลือของ path, ห้าม `.` / `..`) และรับรายการ library แบบ inject ได้ (เหมือน `resolve_link(link, libraries)`) เพื่อให้ test ไม่ต้องอ่าน registry จริง
- **E-2 รูปแบบลิงก์ไฟล์ที่รับ** (query อื่นที่ไม่ใช่ `id` / `RootFolder` เช่น `?web=1`, `?d=...&csf=1&web=1&e=...` ไม่มีผล)
  1. **path ตรง** `https://<t>.sharepoint.com/sites/<Team>/Shared%20Documents/UAT/File.xlsx` (รวมแบบที่มี `?web=1` จาก Excel desktop: File > Info > Copy path และแบบที่มีช่องว่างไม่ encode)
  2. **address bar** ที่มี `?id=` (หรือ `RootFolder=`) ชี้ไปที่ไฟล์ เช่น `.../Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents%2FUAT%2FFile.xlsx&parent=...` และ `.../_layouts/15/onedrive.aspx?id=%2Fpersonal%2F...%2FFile.xlsx`
  3. **ลิงก์แชร์แบบ path** `/:x:/r/<path>` (ลิงก์ "People with existing access") -> ใช้ `<path>`
- **E-3 ลิงก์ที่ไม่รับ (opaque) -> error ที่บอกวิธีได้ลิงก์ที่ใช้ได้** ลิงก์แชร์แบบ opaque (`/:x:/s/...`, `/:x:/g/...` และ `/:<kind>:/<k>/` อื่นที่ไม่ใช่ `r`) และลิงก์เปิดไฟล์ใน Office Online ที่ไม่มี path (`_layouts/15/Doc.aspx?sourcedoc=...`, `_layouts/15/WopiFrame.aspx?sourcedoc=...`) -> raise error code `link_sharing` (หรือ code ใหม่ที่ architect เลือก) ข้อความแนะนำ: ใน Excel desktop ใช้ File > Info > Copy path หรือใน SharePoint ใช้ Copy link แบบ "People with existing access" แล้ววางใหม่ **ไม่**เดาไฟล์จากพารามิเตอร์ `file=` ของ `Doc.aspx` (ไม่มี path ของโฟลเดอร์)
- **E-4 ลิงก์ต้องชี้ไฟล์ `.xlsx`** หลังแปลงลิงก์ (E-2) segment สุดท้ายต้องลงท้ายด้วย `.xlsx` (ไม่สนตัวพิมพ์) ไม่ใช่ -> error code ใหม่ (เช่น `link_not_xlsx`) **ก่อน**ตรวจ library ที่ sync (ข้อความบอกว่าลิงก์ต้องเป็นไฟล์ .xlsx ไม่ใช่โฟลเดอร์ / ไฟล์ชนิดอื่น เช่น `.xls`, `.xlsm`, `.csv`) ลิงก์โฟลเดอร์ (`/:f:/r/...`, หน้า library ที่ไม่มี `?id=`) เข้ากรณีนี้
- **E-5 ลิงก์ที่ไม่ใช่ URL http(s)** -> error code `link_invalid` (ข้อความสำหรับไฟล์ ไม่ใช่ข้อความ "โฟลเดอร์" เดิม)
- **E-6 library ไม่ได้ sync บนเครื่องนี้** (ไม่มี library ใดที่ host ตรงและ `UrlNamespace` เป็น prefix ของ path) -> error code `link_not_synced` ข้อความบอกวิธีแก้: เปิด library ใน SharePoint กด Sync (หรือ "Add shortcut to My files") รอ OneDrive sync เสร็จแล้วลองใหม่ กรณี shortcut ให้ใช้ปุ่มเลือกไฟล์เดิม ("นำเข้าไฟล์ Excel" / "เลือกไฟล์ใหม่ (Relink)…") แล้วเลือกไฟล์ใน OneDrive ของตน
- **E-7 ไฟล์ยังไม่มีในเครื่อง / หาย** path ในเครื่องที่ได้ไม่มีอยู่ (OneDrive ยัง sync ไม่เสร็จ, โฟลเดอร์ถูกตัดออกใน "Choose folders", ไฟล์ถูกลบ / ย้าย / ลิงก์พิมพ์ผิด) -> error code ใหม่ (เช่น `link_file_not_found`) ที่มี path ในเครื่อง ข้อความแนะนำให้รอ OneDrive sync / ตรวจว่าโฟลเดอร์ไม่ถูกตัดออกจากการ sync / ตรวจลิงก์ แล้วลองใหม่ path ที่มีอยู่แต่เป็นโฟลเดอร์ -> error เดียวกับ E-4
- **E-8 อ่านไฟล์ไม่ได้** ไฟล์มีอยู่ (เช่น placeholder ของ Files On-Demand) แต่การ copy / อ่านเพื่อตรวจ (`validate_workbook` -> `copy_shared`) raise `OSError` (OneDrive ดาวน์โหลดไม่ได้, offline, ไม่มีสิทธิ์) -> error ที่แปลได้ (เช่น code `link_unreadable`) ที่มี path ในเครื่อง ข้อความแนะนำให้ตรวจว่า OneDrive ทำงานและออนไลน์ แล้วลองใหม่ ไม่มี exception หลุดออกไปเป็น crash และไม่มีการบันทึกใด ๆ
- **E-9 (a) นำเข้าโปรเจกต์ใหม่จากลิงก์** service ใหม่ (เช่น `import_project_link(base, appdb, person, *, link) -> dict`) ลำดับ: ตรวจสิทธิ์ (`require_manager` เหมือน `import_project`) -> แปลงลิงก์ (E-1 ถึง E-8) -> ใช้การตรวจเดียวกับ `import_project` ทุกข้อ (`validate_workbook` กับ `header_overrides`, ชื่อไฟล์ซ้ำกับโปรเจกต์ใดก็ได้รวม Archive, ชื่อโปรเจกต์จากชื่อไฟล์ `name_from_file`) -> สร้างโปรเจกต์ด้วย `excel_path` = path ในเครื่องที่แปลงได้ (ไม่เปลี่ยนชื่อไฟล์ ไม่ copy ไฟล์) ผู้สร้างเป็น Admin ของโปรเจกต์ใหม่เหมือนเดิม ถ้าไม่ผ่านข้อใด raise `ProjectError` (หรือ subclass ที่มี code) และไม่มีโปรเจกต์ / event / user ใหม่
- **E-10 (b) เปลี่ยนไฟล์ Excel ของโปรเจกต์จากลิงก์** service ใหม่ (เช่น `relink_link(base, appdb, person, project_id, link) -> str`) ลำดับ: ตรวจสิทธิ์ (`require_manager` เหมือน `relink`) -> มีโปรเจกต์ -> แปลงลิงก์ (E-1 ถึง E-8) -> ใช้การตรวจเดียวกับ `relink` ทุกข้อ (`validate_workbook`, ชื่อไฟล์ซ้ำกับโปรเจกต์อื่น, App ID ตรงอย่างน้อย `RELINK_MIN_MATCH` = 80%) -> บันทึก `excel_path` ใหม่ คืนข้อความสรุปเดียวกับ `relink` ถ้าไม่ผ่านข้อใด `excel_path` ไม่เปลี่ยนและไม่มี event
- **E-11 audit / event** (`project_event` ใน `app.db` แบบเดียวกับเดิม)
  - (a) action `"imported"` detail = `"existing file <path ในเครื่อง> (SharePoint: <ลิงก์ที่ clean แล้ว>)"`
  - (b) action `"relinked"` detail = `"<path เดิม> -> <path ใหม่> (<summary>) (SharePoint: <ลิงก์ที่ clean แล้ว>)"`
  - audit ใน tracker.db ของผู้สร้างเป็น Admin (`profile_created`) คงเดิม
  - การนำเข้า / Relink แบบเลือกไฟล์เดิม detail ไม่เปลี่ยน (ไม่มีข้อความ SharePoint)
- **E-12 หลังสำเร็จใช้เหมือนไฟล์ local** โปรเจกต์ที่ได้มี `excel_path` เป็น path ในเครื่องธรรมดา การอ่านครั้งแรก (`run_import` ใน worker เหมือน `import_project_flow`), sync / watcher, การเขียนกลับ Excel ผ่าน COM, backup และ banner "ไม่พบไฟล์ Excel" ทำงานแบบเดิมโดยไม่ต้องแก้
- **E-13 UI (a)** ปุ่มใหม่ "นำเข้าจากลิงก์ SharePoint…" ในที่เดียวกับปุ่ม "นำเข้าไฟล์ Excel" (`t("proj.import")`) ทั้งสองที่: sidebar ของ `MainWindow` (ใต้หรือข้าง `imp_proj_btn`, แสดงเฉพาะเมื่อ `imp_proj_btn` แสดง คือผู้ที่ `can_manage`) และ `NoProjectWindow` (เฉพาะ manager) **(ข้อเสนอ — รอยืนยัน: ตำแหน่ง/รูปแบบ ดู E1)** กดแล้ว:
  1. `QInputDialog.getText` ขอลิงก์ (ยกเลิก หรือข้อความว่าง -> ไม่ทำอะไร)
  2. แปลงลิงก์ + `validate_workbook` (ไม่ผ่าน -> `error()` ข้อความแปลตาม code หรือรายการปัญหาของ workbook)
  3. `confirm()` ที่แสดงลิงก์, path ในเครื่อง, ชื่อไฟล์ และชื่อโปรเจกต์ที่จะได้ (ตอบ "ไม่" -> ไม่สร้างโปรเจกต์)
  4. เรียก service (E-9) แล้วอ่านไฟล์ครั้งแรกใน worker และเปิดโปรเจกต์ใหม่ แบบเดียวกับ `import_project_flow` (`on_done` -> `open_project`)
- **E-14 UI (b)** หน้า Settings การ์ดโปรเจกต์ แถว "ไฟล์ Excel" เพิ่มปุ่ม "เลือกไฟล์ใหม่จากลิงก์ SharePoint…" ข้างปุ่ม Relink เดิม แสดงเงื่อนไขเดียวกับปุ่ม Relink (`self.manager`) กดแล้ว: `getText` ขอลิงก์ (ยกเลิก/ว่าง -> ไม่ทำอะไร) -> `confirm()` ที่แสดงลิงก์และ path ในเครื่อง -> เรียก service (E-10) ใน worker (เหมือน `relink_flow`, ข้อความ `work.relinking`) -> สำเร็จ: `info(t("proj.relinked", s=summary))` แล้ว `open_project(pid)` เหมือนเดิม, ไม่สำเร็จ: `error()` ข้อความแปลตาม code (หรือข้อความ `ProjectError` เดิมสำหรับกรณีที่ไม่มี code เช่น App ID ไม่ถึง 80%)
- **E-15 i18n** ข้อความใหม่ทุกข้อความอยู่ใน `ui/i18n.py` มีทั้ง EN และ TH (ถ้อยคำจริงปรับได้ตราบที่ความหมายตรง) ข้อความ error ของไฟล์**แยก key** จาก `proj.att_link_*` ของโฟลเดอร์ไฟล์แนบ (ข้อความเดิมพูดถึง "โฟลเดอร์" ห้ามแก้ความหมายของ key เดิม) ข้อเสนอ:

  | key | EN | TH |
  |---|---|---|
  | `proj.import_link` | Import from SharePoint link… | นำเข้าจากลิงก์ SharePoint… |
  | `proj.relink_link` | Relink from SharePoint link… | เลือกไฟล์ใหม่จากลิงก์ SharePoint… |
  | `proj.xl_link_title` | SharePoint Excel file link | ลิงก์ไฟล์ Excel บน SharePoint |
  | `proj.xl_link_prompt` | Paste the link of the .xlsx file on SharePoint (in Excel: File > Info > Copy path): | วางลิงก์ของไฟล์ .xlsx บน SharePoint (ใน Excel: File > Info > Copy path): |
  | `proj.import_link_q` | Import the SharePoint file:\n{u}\n\nOneDrive file on this PC:\n{p}\n\nas project '{n}'? The file is not renamed and its other sheets are not touched. | นำเข้าไฟล์บน SharePoint:\n{u}\n\nไฟล์ที่ OneDrive sync บนเครื่องนี้:\n{p}\n\nเป็นโปรเจกต์ '{n}' ใช่ไหม? ไฟล์จะไม่ถูกเปลี่ยนชื่อ และชีตอื่นจะไม่ถูกแตะต้อง |
  | `proj.relink_link_q` | Use the SharePoint file:\n{u}\n\nOneDrive file on this PC:\n{p}\n\nas the Excel file of '{n}'? | ใช้ไฟล์บน SharePoint:\n{u}\n\nไฟล์ที่ OneDrive sync บนเครื่องนี้:\n{p}\n\nเป็นไฟล์ Excel ของ '{n}' ใช่ไหม? |
  | `proj.xl_link_invalid` | This is not a SharePoint file link: {p}\nIn Excel use File > Info > Copy path, then paste it here. | ลิงก์นี้ไม่ใช่ลิงก์ไฟล์บน SharePoint: {p}\nใน Excel ให้ใช้ File > Info > Copy path แล้ววางที่นี่ |
  | `proj.xl_link_sharing` | This kind of link can't be used: {p}\nIn Excel use File > Info > Copy path, or in SharePoint use Copy link with "People with existing access", then paste it here. | ใช้ลิงก์แบบนี้ไม่ได้: {p}\nใน Excel ให้ใช้ File > Info > Copy path หรือใน SharePoint ใช้ Copy link แบบ "People with existing access" แล้ววางที่นี่ |
  | `proj.xl_link_not_xlsx` | The link must point to an Excel .xlsx file (not a folder or another file type): {p} | ลิงก์ต้องชี้ไปที่ไฟล์ Excel .xlsx (ไม่ใช่โฟลเดอร์หรือไฟล์ชนิดอื่น): {p} |
  | `proj.xl_link_not_synced` | OneDrive on this PC does not sync the library of this link:\n{p}\n\nTo fix: open the SharePoint library in the browser, click Sync or "Add shortcut to My files", wait until OneDrive has finished, then try again. With a shortcut, use "Import Excel" / "Relink file…" and pick the file inside your OneDrive. | OneDrive บนเครื่องนี้ยังไม่ได้ sync library ของลิงก์นี้:\n{p}\n\nวิธีแก้: เปิด library ใน SharePoint ด้วยเบราว์เซอร์ กด Sync หรือ "Add shortcut to My files" รอ OneDrive sync เสร็จ แล้วลองใหม่ ถ้าใช้ shortcut ให้กด "นำเข้าไฟล์ Excel" / "เลือกไฟล์ใหม่ (Relink)…" แล้วเลือกไฟล์ใน OneDrive ของคุณ |
  | `proj.xl_link_not_found` | The link points to this file, which is not on this PC yet:\n{p}\nWait until OneDrive has finished syncing, check that its folder is not excluded in OneDrive "Choose folders", or check the link, then try again. | ลิงก์นี้ชี้ไปที่ไฟล์ที่ยังไม่มีบนเครื่องนี้:\n{p}\nกรุณารอ OneDrive sync ให้เสร็จ ตรวจว่าโฟลเดอร์ของไฟล์ไม่ได้ถูกตัดออกใน "Choose folders" ของ OneDrive หรือตรวจลิงก์อีกครั้ง แล้วลองใหม่ |
  | `proj.xl_link_unreadable` | The file can't be read from OneDrive on this PC:\n{p}\nCheck that OneDrive is running and online, then try again. | อ่านไฟล์จาก OneDrive บนเครื่องนี้ไม่ได้:\n{p}\nกรุณาตรวจว่า OneDrive ทำงานและเชื่อมต่ออินเทอร์เน็ตอยู่ แล้วลองใหม่ |

## นอกขอบเขต (Out of scope)
- การเข้าถึง SharePoint โดยตรง (Graph / MSAL / REST / login), การดาวน์โหลดไฟล์มาเก็บเอง, dependency ใหม่
- การสั่งให้ OneDrive sync library / ดาวน์โหลดไฟล์ (เช่น "Always keep on this device") หรือการตรวจสถานะ sync ของ OneDrive
- การ resolve ลิงก์แชร์แบบ opaque (`/:x:/s/`, `/:x:/g/`, `Doc.aspx?sourcedoc=`) ด้วยการเรียก server
- การหาไฟล์ผ่าน shortcut "Add shortcut to My files" (ไม่มี registry ที่จับคู่ URL ได้แบบเดียวกัน ใช้การเลือกไฟล์เดิมแทน)
- การเก็บลิงก์ SharePoint ของไฟล์ Excel ใน column ใหม่ของ `project` / การแสดงลิงก์บนการ์ดโปรเจกต์ (ลิงก์อยู่ใน event detail เท่านั้น ดู Q2)
- การเปลี่ยนการทำงานของปุ่ม "นำเข้าไฟล์ Excel" / Relink แบบเลือกไฟล์เดิม และของฟีเจอร์ลิงก์ SharePoint ของโฟลเดอร์ไฟล์แนบ
- ปุ่มลิงก์ SharePoint ใน banner "ไม่พบไฟล์ Excel" ของ `MainWindow` (ยังใช้ Relink แบบเลือกไฟล์ ดู Q3)
- คำสั่ง CLI (`cli_project.py` import / relink) แบบรับลิงก์ (ดู Q4)
- โหมด single-file (ช่อง Excel path ใน Settings) และการสร้างโปรเจกต์ใหม่ (`create_project`) ลงโฟลเดอร์ SharePoint
- ไฟล์ชนิดอื่นนอกจาก `.xlsx`

## User stories
- **US-1** ในฐานะ Admin หรือ PMO ฉันต้องการวางลิงก์ไฟล์ Excel บน SharePoint เพื่อนำเข้าเป็นโปรเจกต์ใหม่ โดยไม่ต้องไปหาไฟล์เองในโฟลเดอร์ OneDrive ของเครื่อง
- **US-2** ในฐานะ Admin หรือ PMO ฉันต้องการเปลี่ยนไฟล์ Excel ของโปรเจกต์ที่มีอยู่เป็นไฟล์บน SharePoint ด้วยการวางลิงก์ เพื่อย้ายโปรเจกต์ไปใช้ไฟล์กลางของทีม
- **US-3** ในฐานะผู้ใช้ที่วางลิงก์ที่ใช้ไม่ได้ (ลิงก์แชร์ opaque, ไม่ใช่ไฟล์ .xlsx, library ยังไม่ sync, ไฟล์ยังไม่มาถึงเครื่อง) ฉันต้องการเห็นข้อความที่บอกสาเหตุและวิธีแก้ เพื่อแก้เองได้โดยไม่ต้องถามผู้ดูแล
- **US-4** ในฐานะ PMO ฉันต้องการเห็นในประวัติโปรเจกต์ว่าโปรเจกต์ถูกนำเข้า / เปลี่ยนไฟล์จากลิงก์ SharePoint ใด เพื่อตรวจสอบย้อนหลัง

## Acceptance criteria
ทุกข้อทดสอบด้วย pytest โดยใช้ `tmp_path` จำลอง `MountPoint` ของ OneDrive (โฟลเดอร์ธรรมดา), รายการ `SyncedLibrary` ปลอมที่ inject เข้า resolver (หรือ monkeypatch `synced_libraries`) ไม่อ่าน registry จริง, ไฟล์ workbook ที่สร้างด้วย `create_workbook()` / helper ใน `tests/test_projects.py`, `AppDb` จริง ไม่ใช้ Excel / COM ส่วน UI ใช้ `QApplication` แบบ offscreen ตามแนว `tests/test_ui.py` โดย monkeypatch `QInputDialog.getText`, `confirm`, `info`, `error` และ `run_in_worker` (ให้รันทันที) กรณีอ่านไฟล์ไม่ได้จำลองด้วย monkeypatch ให้ `copy_shared` raise `OSError`

ในตัวอย่างข้างล่าง: library ปลอม `L` = `SyncedLibrary("https://t.sharepoint.com/sites/Team/Shared Documents", "<tmp>/OneDrive - Org/Team - Documents")` และไฟล์จริง `F` = `<MountPoint>/UAT/Planning Progress Demo.xlsx`

### การแปลงลิงก์ไฟล์ (E-1 ถึง E-7)
- **AC-1** (path ตรง) Given `L` และลิงก์ `https://t.sharepoint.com/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx` (ทดสอบซ้ำแบบมี `?web=1` และแบบช่องว่างไม่ encode และแบบมี `"` / ช่องว่างครอบ) When แปลงลิงก์ Then ได้ path `F`
- **AC-2** (address bar) Given `L` และลิงก์ `https://t.sharepoint.com/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents%2FUAT%2FPlanning%20Progress%20Demo.xlsx&parent=%2Fsites%2FTeam%2FShared%20Documents%2FUAT` When แปลงลิงก์ Then ได้ path `F`
- **AC-3** (ลิงก์แชร์แบบ path) Given `L` และลิงก์ `https://t.sharepoint.com/:x:/r/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx?d=w123&csf=1&web=1&e=abc` When แปลงลิงก์ Then ได้ path `F`
- **AC-4** (library ที่ prefix ยาวที่สุดชนะ) Given library 2 ตัวที่ `UrlNamespace` เป็น `.../sites/Team/Shared Documents` และ `.../sites/Team/Shared Documents/UAT` (MountPoint ต่างกัน) When แปลงลิงก์ของ AC-1 Then ได้ path ใต้ MountPoint ของตัวที่สอง
- **AC-5** (opaque -> error) Given ลิงก์ `https://t.sharepoint.com/:x:/s/Team/EaBcDeF123`, `https://t.sharepoint.com/:x:/g/personal/u/EaBc`, และ `https://t.sharepoint.com/sites/Team/_layouts/15/Doc.aspx?sourcedoc=%7BGUID%7D&file=Planning%20Progress%20Demo.xlsx&action=default` (parametrize) When แปลงลิงก์ Then raise error ที่มี code ของ E-3 และ `path` เป็นลิงก์ที่ clean แล้ว
- **AC-6** (ไม่ใช่ .xlsx) Given ลิงก์ที่ชี้ `.../UAT/Report.xls`, `.../UAT/Report.xlsm`, `.../UAT/Data.csv`, โฟลเดอร์ `.../UAT` (path ตรง), และ `/:f:/r/sites/Team/Shared%20Documents/UAT` (parametrize) When แปลงลิงก์ Then raise error code ของ E-4 แม้จะไม่มี library ที่ sync เลย (ตรวจ .xlsx ก่อน sync)
- **AC-7** (นามสกุลตัวพิมพ์ใหญ่) Given ลิงก์ที่ลงท้าย `.XLSX` และไฟล์นั้นมีอยู่ When แปลงลิงก์ Then ไม่ error
- **AC-8** (ไม่ใช่ URL) Given `"abc"`, `"C:\\x.xlsx"`, `"ftp://t.sharepoint.com/a.xlsx"` (parametrize) When แปลงลิงก์ Then raise error code `link_invalid`
- **AC-9** (ไม่ sync) Given ลิงก์ของ AC-1 แต่รายการ library ว่าง หรือมีแค่ library ของ host / site อื่น When แปลงลิงก์ Then raise error code `link_not_synced` ที่มีลิงก์
- **AC-10** (ไฟล์ยังไม่มีในเครื่อง) Given `L` แต่ไม่มีไฟล์ `F` (ทดสอบทั้งกรณีโฟลเดอร์ `UAT` มีอยู่และไม่มี) When แปลงลิงก์ (ผ่านเส้นทางของ service E-9 / E-10) Then raise error code ของ E-7 ที่มี path `F` และระบบไม่สร้างไฟล์หรือโฟลเดอร์ใด ๆ
- **AC-11** (ปลอดภัย) Given ลิงก์ที่มี segment `..` (เช่น `.../Shared%20Documents/UAT/../../x.xlsx`) When แปลงลิงก์ Then raise `link_invalid` และไม่คืน path นอก MountPoint
- **AC-12** (ไม่ถดถอย) Given test / พฤติกรรมเดิมของลิงก์โฟลเดอร์ไฟล์แนบ (`resolve_link`, `resolve_attachment_link`, `set_attachment_link`) When รัน Then ผลเหมือนเดิมทุกกรณี (รวม `/:f:/r/` ยังรับเป็นโฟลเดอร์ได้)

### (a) นำเข้าโปรเจกต์ใหม่จากลิงก์ (E-8, E-9, E-11, E-12)
- **AC-13** (สำเร็จ) Given ผู้ใช้เป็น Admin/PMO ของโปรเจกต์ใดก็ได้ (หรือยังไม่มีโปรเจกต์เลย), `L` และไฟล์ `F` ที่ผ่าน `validate_workbook` When เรียก service (a) ด้วยลิงก์ของ AC-1 Then มีโปรเจกต์ใหม่ชื่อ `Demo` (จาก `name_from_file`), `excel_path == str(F)`, ผู้เรียกเป็น Admin ของโปรเจกต์ใหม่ และไฟล์ `F` ไม่ถูกเปลี่ยนชื่อ / แก้ไข (bytes เท่าเดิม)
- **AC-14** (event) Given AC-13 Then `appdb.events(<id ใหม่>)` มี action `"imported"` 1 แถว ที่ `editor` = ผู้เรียก และ detail มีทั้ง path `F` และข้อความ `SharePoint: <ลิงก์ที่ clean แล้ว>`
- **AC-15** (ไม่มีสิทธิ์) Given ผู้ใช้ที่เป็นแค่ QA / User BU/PM ในทุกโปรเจกต์ When เรียก service (a) Then raise `ProjectError` ไม่มีสิทธิ์ และไม่มีโปรเจกต์ / event ใหม่ (ตรวจสิทธิ์ก่อนแปลงลิงก์ ไม่ต้องมี library)
- **AC-16** (ชื่อไฟล์ซ้ำ) Given มีโปรเจกต์อื่น (ทดสอบทั้งแบบปกติและแบบ Archive) ใช้ไฟล์ชื่อ `Planning Progress Demo.xlsx` อยู่แล้ว (คนละโฟลเดอร์) When นำเข้าจากลิงก์ของ AC-1 Then raise `ProjectError` แบบเดียวกับ `import_project` และไม่มีโปรเจกต์ใหม่
- **AC-17** (workbook ไม่ถูกต้อง) Given `F` เป็น .xlsx ที่ขาดชีต / header ที่ต้องใช้ When นำเข้าจากลิงก์ Then raise `ProjectError` ที่มีรายการปัญหาเดียวกับ `validate_workbook` และไม่มีโปรเจกต์ใหม่
- **AC-18** (อ่านไฟล์ไม่ได้) Given `F` มีอยู่แต่ `copy_shared` raise `OSError` When นำเข้าจากลิงก์ Then raise error code ของ E-8 ที่มี path `F` (ไม่ใช่ `OSError` ดิบ) และไม่มีโปรเจกต์ / event ใหม่
- **AC-19** (error ของลิงก์ไม่สร้างอะไร) Given แต่ละกรณีของ AC-5, AC-6, AC-8, AC-9, AC-10 When เรียก service (a) Then raise error ตาม code ของกรณีนั้น และจำนวนโปรเจกต์ / event ใน `app.db` ไม่เปลี่ยน
- **AC-20** (ใช้เหมือนไฟล์ local) Given โปรเจกต์จาก AC-13 When `project_config` + `run_import` (เหมือน job ใน `import_project_flow`) Then อ่านข้อมูลจาก `F` ได้เท่ากับการนำเข้าไฟล์เดียวกันด้วย `import_project(path=F)` และ `cfg.excel_path == F`

### (b) เปลี่ยนไฟล์ของโปรเจกต์จากลิงก์ (E-8, E-10, E-11)
- **AC-21** (สำเร็จ) Given โปรเจกต์ P ที่มีรายการ (App ID) อยู่แล้ว ผู้ใช้เป็น Admin/PMO และ `F` มี App ID ของ P ครบ When เรียก service (b) ด้วยลิงก์ของ AC-1 Then `appdb.get_project(P.id)["excel_path"] == str(F)` และค่าที่คืนเท่ากับ summary ที่ `relink(..., F)` จะคืน (`"<n> of <m> App IDs found"`)
- **AC-22** (event) Given AC-21 Then มี event action `"relinked"` ใหม่ 1 แถว detail มี path เดิม, path `F`, summary และ `SharePoint: <ลิงก์ที่ clean แล้ว>`
- **AC-23** (App ID ไม่ถึง 80%) Given `F` มี App ID ของ P น้อยกว่า 80% When เรียก service (b) Then raise `ProjectError` แบบเดียวกับ `relink`, `excel_path` ไม่เปลี่ยน และไม่มี event ใหม่
- **AC-24** (ชื่อไฟล์ซ้ำกับโปรเจกต์อื่น) Given โปรเจกต์อื่น Q ใช้ชื่อไฟล์ `Planning Progress Demo.xlsx` When เปลี่ยนไฟล์ของ P จากลิงก์ของ AC-1 Then raise `ProjectError` แบบเดียวกับ `relink` และ `excel_path` ของ P ไม่เปลี่ยน
- **AC-25** (ไม่มีสิทธิ์ / ไม่มีโปรเจกต์) Given ผู้ใช้ที่ไม่ใช่ Admin/PMO ของโปรเจกต์ใดเลย หรือ `project_id` ที่ไม่มีอยู่ When เรียก service (b) Then raise `ProjectError` (ไม่มีสิทธิ์ / `Unknown project`) และไม่มีอะไรเปลี่ยน
- **AC-26** (error ของลิงก์ไม่เปลี่ยนอะไร) Given แต่ละกรณีของ AC-5, AC-6, AC-8, AC-9, AC-10, AC-18 When เรียก service (b) Then raise error ตาม code ของกรณีนั้น, `excel_path` ของ P ไม่เปลี่ยน และไม่มี event ใหม่
- **AC-27** (ไม่ถดถอย) Given `import_project(path=...)` และ `relink(..., path)` แบบเดิม When เรียก Then detail ของ event เหมือนเดิม (ไม่มีข้อความ `SharePoint`) และ test เดิมของ projects ผ่านโดยไม่ต้องแก้

### UI (E-13, E-14)
- **AC-28** (ปุ่ม (a) แสดงเฉพาะ manager) Given ผู้ใช้ที่ `can_manage` เป็น True / False When เปิด `MainWindow` และ `NoProjectWindow` Then ปุ่มข้อความ `t("proj.import_link")` มองเห็นได้เมื่อ True และไม่แสดงเมื่อ False (ตามเงื่อนไขเดียวกับปุ่ม `t("proj.import")`)
- **AC-29** (a สำเร็จ) Given `getText` คืนลิงก์ของ AC-1 + ok, `confirm` ตอบ yes When กดปุ่ม (a) Then `confirm` ถูกเรียกด้วยข้อความ `t("proj.import_link_q", ...)` ที่มีลิงก์, path `F` และชื่อโปรเจกต์ `Demo`, มีโปรเจกต์ใหม่ใน `app.db`, `info()` ถูกเรียกด้วย `t("proj.imported", ...)` และ `open_project` ถูกเรียกด้วย id ใหม่
- **AC-30** (a ยกเลิก) Given `getText` คืน ok = False หรือข้อความว่าง หรือ `confirm` ตอบ no When กดปุ่ม (a) Then ไม่มีโปรเจกต์ใหม่ และ `error()` / `info()` ไม่ถูกเรียก
- **AC-31** (a error แปลแล้ว) Given ลิงก์ของแต่ละกรณี AC-5, AC-6, AC-9, AC-10 When กดปุ่ม (a) Then `error()` ถูกเรียกด้วยข้อความจาก key `proj.xl_link_*` ที่ตรงกับ code (มี path / ลิงก์อยู่ในข้อความ), `confirm` ไม่ถูกเรียก และแอปไม่ crash
- **AC-32** (ปุ่ม (b)) Given หน้า Settings ของโปรเจกต์ When ผู้ใช้เป็น manager / ไม่ใช่ Then ปุ่ม `t("proj.relink_link")` แสดง / ไม่แสดง ตามเงื่อนไขเดียวกับปุ่ม `t("proj.relink")`
- **AC-33** (b สำเร็จ) Given `getText` คืนลิงก์ของ AC-1, `confirm` ตอบ yes When กดปุ่ม (b) Then `confirm` ถูกเรียกด้วย `t("proj.relink_link_q", ...)` ที่มีลิงก์และ path `F`, `excel_path` ของโปรเจกต์เป็น `F`, `info()` ถูกเรียกด้วย `t("proj.relinked", s=...)` และ `open_project(pid)` ถูกเรียก
- **AC-34** (b ยกเลิก / error) Given `getText` ยกเลิก หรือ `confirm` ตอบ no Then service ไม่ถูกเรียก Given service raise error ที่มี code (เช่น `link_not_synced`) Then `error()` ถูกเรียกด้วยข้อความ `proj.xl_link_*` ที่ตรง code Given service raise `ProjectError` ที่ไม่มี code (เช่น App ID ไม่ถึง 80%) Then `error()` แสดงข้อความของ exception และในทุกกรณี `excel_path` ไม่เปลี่ยนและ `open_project` ไม่ถูกเรียก

### i18n (E-15)
- **AC-35** Given ทุก key ใหม่ในตาราง E-15 (และ key อื่นที่เพิ่มในฟีเจอร์นี้) When ตรวจ `ui/i18n.py` Then มีทั้ง EN และ TH ที่ไม่ว่าง และ placeholder (`{p}`, `{u}`, `{n}`) ตรงกันทั้งสองภาษา และทุก error code ที่ resolver / service ของฟีเจอร์นี้ raise ได้มี key ข้อความที่ตรงกัน
- **AC-36** Given ภาษา TH When render ข้อความปุ่ม (AC-28, AC-32) และข้อความ error ของ AC-31 Then เป็นข้อความภาษาไทย
- **AC-37** Given key เดิม `proj.att_link_*` When ตรวจ Then ข้อความเดิมไม่ถูกแก้

### ทั่วไป
- **AC-38** `pytest -q` ผ่านทั้งหมด ห้ามลบหรือแก้ test เดิม (ถ้าจำเป็นต้องแก้ต้องรายงานหัวหน้าว่าแก้ test ใดเพราะอะไร)

## ข้อจำกัด / สมมติฐาน
- Python 3.x, PySide6, openpyxl, SQLite, PEP 8, type hints, pytest ไม่มี dependency ใหม่ (`winreg` เป็น stdlib)
- ใช้ได้เฉพาะ Windows ที่ติดตั้ง OneDrive client และ sync library ของไฟล์นั้นอยู่ (registry `HKCU\Software\SyncEngines\Providers\OneDrive\<id>`: `UrlNamespace`, `MountPoint`) บนเครื่องอื่น / ไม่มี OneDrive `synced_libraries()` คืน `[]` -> error `link_not_synced`
- `excel_path` ที่บันทึกเป็น path ในเครื่องที่แปลงได้ ณ ตอนนำเข้า (มีชื่อ user / ชื่อองค์กรอยู่ใน path) `app.db` อยู่ในเครื่อง แต่ละเครื่องต้องนำเข้า / Relink เอง เหมือนการเลือกไฟล์เดิม
- ถ้าภายหลัง OneDrive เปลี่ยนตำแหน่ง sync หรือเลิก sync library -> ไฟล์หาย -> banner "ไม่พบไฟล์ Excel" + Relink แบบเดิม (E-12)
- OneDrive Files On-Demand: ไฟล์ที่เป็น placeholder (cloud-only) ถือว่ามีอยู่ การอ่านครั้งแรกทำให้ OneDrive ดาวน์โหลด ถ้าดาวน์โหลดไม่ได้ -> E-8 แอปไม่บังคับ "Always keep on this device"
- การเขียนกลับ Excel ผ่าน COM ไปที่ไฟล์ใน OneDrive ใช้กลไก lock / retry / backup เดิม การชนกันกับคนอื่นที่เปิดไฟล์เดียวกันบน SharePoint / Excel Online (co-authoring) อยู่ในพฤติกรรมเดิมของ OneDrive / AutoSave ซึ่งนอกขอบเขต (ตามที่ section B1–B5 กำหนดไว้)
- ลิงก์แชร์ `/:x:/r/` ใช้ได้เพราะมี path ในตัว ลิงก์แชร์แบบอื่น (`/:x:/s/`, `/:x:/g/`) และลิงก์ Office Online (`Doc.aspx?sourcedoc=`) ไม่มี path ต้องถาม server จึงไม่รับ
- ชื่อฟังก์ชัน, ที่อยู่ของฟังก์ชัน, ชื่อ exception / error code (ยกเว้นที่ใช้ร่วมกับของเดิม `link_invalid`, `link_sharing`, `link_not_synced`), วิธี map code เป็นข้อความ และวิธีแยกข้อความไฟล์ออกจากข้อความโฟลเดอร์ เป็นหน้าที่ architect ตราบที่ผ่าน AC (เช่น `AttachmentDirError` เดิมมีชื่อผูกกับไฟล์แนบ architect อาจเพิ่ม exception ใหม่หรือ reuse ก็ได้)

### สรุปข้อเสนอที่รอผู้ใช้ยืนยัน
- **E1** ปุ่ม (a) "นำเข้าจากลิงก์ SharePoint…" เป็นปุ่มแยก วางที่ sidebar ใต้ปุ่ม "สร้างโปรเจกต์ / นำเข้าไฟล์ Excel" และใน `NoProjectWindow` (E-13)
- **E2** ปุ่ม (b) "เลือกไฟล์ใหม่จากลิงก์ SharePoint…" อยู่ข้างปุ่ม Relink ในการ์ดโปรเจกต์หน้า Settings (E-14)
- **E3** สิทธิ์ใช้กติกาเดิมของ import / relink คือ Admin/PMO ของโปรเจกต์ใดก็ได้ (`require_manager`) (E-9, E-10)
- **E4** มีหน้าต่างยืนยันที่แสดงลิงก์ + path ในเครื่องก่อนนำเข้า / เปลี่ยนไฟล์ทุกครั้ง (Relink แบบเลือกไฟล์เดิมไม่มีหน้าต่างยืนยัน) (E-13, E-14)
- **E5** ลิงก์ SharePoint เก็บเฉพาะใน event detail ไม่เก็บเป็น column ของโปรเจกต์ และไม่แสดงบนการ์ด (E-11)
- **E6** ลิงก์ opaque (`/:x:/s/`, `/:x:/g/`, `Doc.aspx?sourcedoc=`) ไม่รับ แจ้งให้ใช้ Copy path ใน Excel หรือ Copy link แบบ "People with existing access" (E-3)

## คำถามที่ยังไม่มีคำตอบ
- **Q1** ตำแหน่ง / รูปแบบปุ่ม (a): ปุ่มแยกที่ sidebar และหน้าจอ "ยังไม่มีโปรเจกต์" (ข้อเสนอ E1) หรือให้ปุ่ม "นำเข้าไฟล์ Excel" เดิมถามก่อนว่าจะเลือกไฟล์หรือวางลิงก์ (sidebar มีที่จำกัด)
- **Q2** ต้องการเก็บลิงก์ SharePoint ของไฟล์ Excel ไว้กับโปรเจกต์ (column ใหม่ + migration) และแสดงบนการ์ดโปรเจกต์ใน Settings แบบเดียวกับลิงก์โฟลเดอร์ไฟล์แนบหรือไม่ (ข้อเสนอ E5: ไม่ เก็บแค่ใน event detail) ถ้าเก็บ การ Relink แบบเลือกไฟล์ควรล้างลิงก์หรือไม่
- **Q3** banner "ไม่พบไฟล์ Excel" ใน `MainWindow` ควรมีตัวเลือก "เลือกไฟล์ใหม่จากลิงก์ SharePoint…" ด้วยหรือไม่ (ตอนนี้นอกขอบเขต ยังเป็น Relink แบบเลือกไฟล์)
- **Q4** ต้องการให้คำสั่ง CLI `cli_project.py` (import / relink) รับลิงก์ SharePoint ด้วยหรือไม่ (ตอนนี้นอกขอบเขต)
- **Q5** สิทธิ์ (b): ใช้กติกาเดิมของ Relink (Admin/PMO ของโปรเจกต์ใดก็ได้, ข้อเสนอ E3) หรือจำกัดเฉพาะ Admin/PMO ของโปรเจกต์นั้น แบบโฟลเดอร์ไฟล์แนบ
- **Q6** ต้องการหน้าต่างยืนยันก่อนเปลี่ยนไฟล์จากลิงก์ (ข้อเสนอ E4) หรือทำเหมือน Relink เดิมที่ไม่ถาม

## การตัดสินใจของผู้ใช้ (26-Nov-2026) — อนุมัติ spec
- Q1 **ไม่ใช้ปุ่มแยก (แทน E1):** ปุ่ม "นำเข้าไฟล์ Excel" เดิม (`proj.import`) ใน sidebar และ `NoProjectWindow` ถามก่อนว่าจะ "เลือกไฟล์…" หรือ "วางลิงก์ SharePoint…" แล้วไปต่อตาม flow เดิม / flow ลิงก์ (AC ที่อ้างตำแหน่งปุ่มแยกใน E-13 ให้ตีความตามข้อนี้)
- Q2 ไม่เก็บลิงก์เป็น column ของโปรเจกต์ เก็บแค่ใน event detail (ตาม E5)
- Q3 banner "ไม่พบไฟล์ Excel" -> นอกขอบเขตรอบนี้
- Q4 CLI รับลิงก์ -> นอกขอบเขตรอบนี้
- Q5 สิทธิ์ (b) ใช้กติกาเดิมของ Relink: Admin/PMO ของโปรเจกต์ใดก็ได้ (ตาม E3)
- Q6 มีหน้าต่างยืนยันแสดงลิงก์ + path ในเครื่องก่อนนำเข้า / เปลี่ยนไฟล์ (ตาม E4)
- E2 (ปุ่ม (b) ข้าง Relink ในหน้า Settings) และ E6 (ไม่รับลิงก์ opaque) -> อนุมัติตามข้อเสนอ
- **(26-Nov-2026 ภายหลัง) ลิงก์ใน event detail ไม่มี query string:** ข้อความ ` (SharePoint: <url>)` ของ AC-14 / AC-22 ใช้ URL ตรงของไฟล์ในรูป `https://<host>/<path segments ที่ encode แล้ว>` (เช่น `https://t.sharepoint.com/sites/Team/Shared%20Documents/UAT/Demo.xlsx`) สร้างจาก host + path ที่ `link_location` แยกได้ ไม่มี `?web=1`, `?d=...&e=...` หรือ `?id=` (ลิงก์ address bar จะได้ path ของไฟล์จาก `?id=` แทน) ส่วน `exc.path` ของ error ใน AC-5 ยังเป็นลิงก์ที่ผู้ใช้วาง (clean แล้ว) เหมือนเดิม
