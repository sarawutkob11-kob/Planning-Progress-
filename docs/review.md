# Code Review — 2026-11-25

ขอบเขต: T14-T19 "หน้าตาใหม่ของเมล Weekly Report + มีความเสี่ยงล่าช้า" ตาม `docs/spec.md` / `docs/design.md` section 2026-11-25 ดูการเปลี่ยนแปลงที่ยังไม่ commit ทั้งหมด (รวม T13 ที่ปนอยู่: `alerts/templates.py`, `cli_alerts.py`, `ui/i18n.py`) และ test ใหม่ 4 ไฟล์
รัน `pytest -q` ซ้ำด้วย venv ของโปรเจกต์: **646 passed**

## สรุป
**ผ่าน (มีข้อควรแก้)** — 🔴 0, 🟡 3, 🟢 9

ตรงตาม spec:
- **R1/AC-1..3b**: label ของ Amber แก้ที่จุดเดียว (`traffic_light.py` `LABEL_EN/LABEL_TH`) i18n, `COLOR_TH`, `daily_job._queue_traffic_lights` (เทียบกับ `label()` แล้วเขียนทับคำเก่า) และ template JSON เปลี่ยนครบ ค่า internal Red/Amber/Green ใน DB และ style key ไม่เปลี่ยน
- **`missing_fields`**: `compute` ให้ผลและ `reason` เหมือนเดิม (มี test ยืนยัน `test_t16_compute_reason_unchanged`)
- **weekly_data**: `days_ahead=7`, `due_next_7_days`, `days_late` เฉพาะ `overdue`, `missing_data` เรียงตามชื่อ Complete ไม่ถูกนับ `weekly_ai` ไม่อ้างชื่อ key เดิม (ตรวจแล้วไม่มี `red_amber` / `due_next_2_weeks` เหลือใน `src/`)
- **HTML escape**: ข้อความที่มาจากผู้ใช้หรือ AI (ชื่อโปรเจกต์, ชื่อแอป/งาน, System/Detail/Status ของ Defect, เหตุผลการเลื่อนวัน, ข้อความ AI) ผ่าน `_e` ทุกจุด ค่าที่ไม่ escape เป็นตัวเลข int จาก `_n` / ค่าสีจากค่าคงที่ / `_width` ที่ parse เป็น float แล้วเท่านั้น banner DRY RUN ยังเข้าหลัง `<body>` (render มี `\n` หลัง body ตามที่ `with_banner` ต้องใช้)
- **Outlook (ระดับโครงสร้าง)**: ไม่มี flex / grid / `<style>` / border-radius ใช้ table ซ้อน + inline style กว้าง 640 มี `width` attribute คู่กับ style
- **เรื่องภายในของ PMO**: render ไม่รับ notes แล้ว ข้อความ AI error, ไม่มีผู้รับ, อีเมลไม่ถูกต้อง, คิวตรวจ ไม่อยู่ใน HTML; `run_weekly` แจ้ง PMO 3 เรื่องตาม AC-21; preview แสดงอีเมลไม่ถูกต้องผ่าน `preview_notes`
- **โค้ดตาย**: `_TH_COLOR`, `_KIND`, `_STYLE_CELL`, `_h`, `_table`, `_ai_html`, import `footer` ถูกลบหมด ไม่มีผู้เรียก `render` แบบ signature เดิมเหลือ
- **Test**: ครอบคลุม AC-4 ถึง AC-23 ครบ assert เนื้อหาจริง (regex ตำแหน่ง/สี/ลำดับ, escape, dry-run banner, จำกัด 10 รายการ)

## 🔴 Critical (ต้องแก้ก่อน merge)
- ไม่มี

## 🟡 Warning (ควรแก้)
- [src/uat_tracker/services/weekly_report.py:167, 100-107, 132, 294, 200-201, 219-220] ใช้ `padding` / `margin` / `border` บน `<div>` และ `<span>` (กล่องบรรทัดสรุป `padding:10px 14px`, เส้นคั่นหัวข้อ `padding-top:8px`, ป้ายสถานะ/ป้ายหน้า `border + padding:0 6px`) และแถบสีใช้ `&nbsp;` + `font-size:0;line-height:0` — Outlook desktop (render ด้วย Word) ไม่รองรับ padding บน div/span และไม่เคารพ `line-height:0` ทำให้ตัวหนังสือชิดขอบกล่องเทา ป้ายสถานะไม่มีระยะ และแถบ 6px/12px สูงขึ้นเป็นความสูงบรรทัด (~15-18px) หน้าตาจะต่างจากแบบที่อนุมัติ (test AC-18 ตรวจแค่ไม่มี flex/grid/style จึงจับไม่ได้) — ย้าย padding/พื้นหลัง/เส้นไปไว้ที่ `<td>` ของตารางซ้อน, เพิ่ม `mso-line-height-rule:exactly;line-height:6px` ใน cell ของแถบ และส่ง dry-run เข้า Outlook จริงหนึ่งครั้งก่อนเปิดใช้ (design.md ระบุไว้แล้วแต่ยังไม่มีหลักฐานว่าทำ)
- [src/uat_tracker/services/weekly_data.py:145] `review_queue` ยังอยู่ใน `data` ที่ส่งให้ AI ทั้งก้อน — เมื่อ AI สำเร็จ AI อาจเขียนถึง "คิวตรวจของ PMO มี n รายการ" ใน `summary_th` / `why_th` แล้วข้อความนั้นถูกวางในเมลผู้บริหารตรง ๆ ขัดกับ R13/US-4 (AC-17 ทดสอบเฉพาะกรณีไม่มี AI) — ทางแก้ต้องให้ผู้ใช้ตัดสินใจเพราะ spec ระบุ "เนื้อหาที่ส่งให้ AI" เป็นนอกขอบเขต: เอา `review_queue` ออกจาก payload ที่ส่ง AI (ยังเก็บใน `data_json` ได้) หรือยอมรับความเสี่ยงไว้
- [README.md:94] แถว "Color history" ยังเขียนว่า "The Excel Traffic Light column is not changed (write-back is Part 3)" ขณะที่แถว Traffic Light ด้านบน (บรรทัด 93) ที่เพิ่งแก้บอกว่าคอลัมน์ Excel แสดงคำใหม่และ daily job เขียนทับ — เอกสารขัดกันเองในตารางเดียว (R1b) ผู้อ่านไม่รู้ว่าโปรแกรมแตะคอลัมน์นี้หรือไม่ — แก้ประโยคให้ตรงกับ `_queue_traffic_lights`

## 🟢 Suggestion (ทางเลือก)
- [src/uat_tracker/services/weekly_report.py:510-512] `_invalid_detail` ใช้ `b['app']` ตรง ๆ ถ้าแถวไม่มีชื่อแอปจะได้ "None (bad-address)" ในเมลแจ้ง PMO ขณะที่ `preview_notes` (บรรทัด 349+) ใช้ `b.get('app') or b.get('id')` — ใช้ fallback แบบเดียวกัน และพิจารณารวมเป็น helper เดียวกัน
- [src/uat_tracker/ui/weekly.py:103] ค่าอีเมลที่ไม่ถูกต้องมาจาก Excel ตรง ๆ แล้วใส่ใน `QLabel` ที่เป็น `Qt.AutoText` ถ้าค่ามีรูปแบบ `<...>` Qt อาจตีความเป็น rich text ทำให้ค่าหายจากบรรทัด preview — ตั้ง `setTextFormat(Qt.PlainText)` ให้ label นี้
- [src/uat_tracker/services/weekly_report.py:314] การเลื่อนวันที่ `where == "app"` แต่ `reason` ว่าง จะแสดง "แก้ใน Excel ไม่ได้ระบุเหตุผล" ซึ่งไม่จริง (ปัจจุบันแทบไม่เกิดเพราะแอปบังคับเหตุผลและ restore มี reason) — แยกข้อความกรณีแก้ในแอปแต่ไม่มีเหตุผล
- [src/uat_tracker/services/weekly_report.py:179] `summary_th` ที่มีขึ้นบรรทัดใหม่จะถูกรวมเป็นย่อหน้าเดียว และถ้า AI ใช้ — / – ในข้อความ HTML จะมีเครื่องหมายที่ AC-18 ห้าม — แปลง `\n` เป็น `<br>` หลัง escape และพิจารณาแทน dash ในข้อความ AI
- [src/uat_tracker/services/weekly_report.py:100-107] `_heading` สร้าง `en_html` แต่ branch `inline` ไม่ได้ใช้ (สร้าง span ซ้ำเอง) และถ้าเรียก `inline=True` กับ `en=""` จะได้ "· " ลอย ๆ — ใช้ `en_html` ตัวเดียวทั้งสอง branch
- [src/uat_tracker/services/weekly_report.py:219] ความกว้างช่องในแถบสถานะมาจาก `pct()` ที่ปัดทศนิยม ผลรวมอาจเป็น 100.1% (เช่น 14.3+28.6+42.9+14.3) ไม่มีผลในเบราว์เซอร์ แต่ถ้าอยากให้แน่นอนให้ช่องสุดท้ายไม่ใส่ width
- [src/uat_tracker/services/weekly_report.py:33-34] import `.traffic_light` อยู่ก่อน `.importer` ไม่เรียงตามตัวอักษรเหมือนบรรทัดอื่น; `cell(..., value, ...)` (บรรทัด 203) ไม่มี type hint; หลายบรรทัดในไฟล์ใหม่ยาวเกิน 120 ตัวอักษร (เช่น 251, 297, 321, 325)
- [tests/test_t16_t18_weekly_email.py:372-386] ข้อมูลทดสอบใช้ชื่อ System "DMS-RAMC" / "Collins" และ Defect #21/#24 ที่ดูเหมือนมาจากไฟล์จริง — ตามกฎไม่ให้ข้อมูลบริษัทเข้า git ควรเปลี่ยนเป็นชื่อสมมติ (เช่น "SYS-A")
- [Claude outputs/, config.demo.toml, run-demo.cmd] เป็นไฟล์ untracked ที่ไม่อยู่ใน `.gitignore` (mockup HTML 3 ไฟล์) ระวังอย่าให้ติดไปตอน commit รอบนี้ ถ้าไม่ได้ตั้งใจ

---

# Code Review — 2026-11-24

ขอบเขต: รอบแก้บั๊ก T1–T6 ตาม `docs/tasks.md` — การเปลี่ยนแปลงที่ยังไม่ commit ใน 12 ไฟล์ใต้ `src/uat_tracker/` และ `tests/test_bugfixes.py` (untracked)
หมายเหตุ: ไม่มี `docs/spec.md` / `docs/design.md` ในรอบนี้ จึงใช้ "ต้องการ" ใน `docs/tasks.md` เป็นเกณฑ์ ไม่ได้รัน pytest ซ้ำ (ผู้สั่งงานแจ้งว่าผ่าน 307 passed)

## สรุป
**ผ่าน (มีข้อควรแก้)** — 🔴 0, 🟡 2, 🟢 7

ทุก task ทำครบตาม "ต้องการ":
- **T1**: ตรวจทุกชีตก่อนเขียน (`_check_rows_unmoved`) ถ้าแถวขยับจะ raise `RowsMoved` โดยไม่แตะ workbook เลย และ sync จะ retry หลังอ่านไฟล์ใหม่ เติม App ID ได้เฉพาะ cell ว่าง, ID ที่ไม่รู้จัก หรือ ID ซ้ำจากแถวที่ copy มา ตรวจ logic ของ `_moved_rows` แล้ว ทั้งกรณี insert, delete, sort, swap และ copy ถูกต้องตรงกับการ bind ใน `binding.bind` (pass 0: ID ซ้ำ แถวแรกได้ ID ไป)
- **T2**: A) reset 'writing' แล้วตั้ง `next_attempt_at=now` และคง `first_failed_at` ไว้ B) ข้าม field เฉพาะตอน ROW ยังรอเขียนอยู่ ถ้า ROW เป็น failed ก็ให้ field fail พร้อมเหตุผล
- **T3**: ใช้ชุด `skip` ขนาดจำกัด จึงลองแต่ละโปรเจกต์ครั้งเดียว ถ้าไม่เหลือโปรเจกต์ที่เปิดได้ก็ไปหน้า no-project ไม่มี recursion ไม่รู้จบ
- **T4**: resolve path และตรวจ `is_relative_to(root)` ก่อนแตะ DB หรือไฟล์ใด ๆ รองรับกรณี `../`, absolute path, drive letter, backslash และ `attachments//x` แล้ว
- **T5**: `apply_overrides` ถูกเรียกทุกครั้งที่ส่ง alert และไม่ได้แก้ `self.cfg`
- **T6**: ลบ import ที่ไม่ได้ใช้ครบทั้ง 7 ชื่อ

## 🔴 Critical (ต้องแก้ก่อน merge)
- ไม่มี

## 🟡 Warning (ควรแก้)

- **[src/uat_tracker/excel/writer.py:100 กับ src/uat_tracker/services/importer.py:128 / src/uat_tracker/excel/reader.py:103] กฎ "แถวนี้เป็นรายการไหม" ของ importer กับ writer ยังไม่ตรงกัน 100%: importer นับคอลัมน์ `app_id` แต่ writer ไม่นับ**
  - ปัญหา: `read_sheets` ใส่ `app_id` ไว้ใน `RawRow.values` และ normalizer ส่ง `values` ทั้งก้อนเข้า `is_defect_item` / `is_item_row` ส่วน `_holds_item` ตัด `APP_ID_KEY` ออกก่อนเรียก `is_item_row` ผลคือแถวที่มีแค่ App ID (UAT Plan ใต้กลุ่ม) หรือมีแค่ No. + App ID (Defect List) importer ถือเป็นรายการ แต่ writer ถือว่า "ไม่ใช่รายการ"
  - สถานการณ์จริง: หลัง clear project หรือย้าย workbook ไฟล์ Excel ยังมี App ID เก่าค้างอยู่ใน hidden column ซึ่งตอนนี้กลายเป็น ID ที่ระบบไม่รู้จัก ถ้ามีแถวที่ผู้ใช้ลบเนื้อหาด้วยการเลือกเฉพาะคอลัมน์ที่มองเห็นแล้วกด Delete (App ID ที่ซ่อนอยู่จะยังค้าง) importer จะสร้างรายการใหม่ N ที่แถวนั้น ใน `_moved_rows` จะได้ `cur` = ID ที่ไม่รู้จัก ซึ่งไม่อยู่ใน `expected` แล้ว `_holds_item` คืน False จึงขึ้น "row R no longer holds an item" ไฟล์บนดิสก์ไม่ได้เปลี่ยน การอ่านซ้ำจึงไม่ช่วย ทุก flush จะเป็น `RowsMoved` วนไป retry ครบ 6 รอบแล้วขึ้น error + อีเมล PMO และเนื่องจาก `_check_rows_unmoved` ตรวจทุกชีต **การเขียนกลับ Excel ของทุกชีตจะหยุดถาวร** จนกว่าจะมีคนไปแก้แถวนั้นเอง
  - แนวทางแก้: ให้ทั้งสองฝั่งใช้ชุดคอลัมน์เดียวกัน โดยให้ `is_item_row` / `is_defect_item` ไม่นับ `APP_ID_KEY` เสมอ (แก้ใน `sheets.py` จุดเดียว) จะได้ไม่ import แถวที่มีแค่ ID ด้วย หรือให้ normalizer ตัด `app_id` ออกก่อนเรียกกฎ แล้วเพิ่ม test: แถว UAT Plan ที่มีแค่ App ID ที่ไม่รู้จัก และแถว Defect ที่มี No. + App ID ที่ไม่รู้จัก ต้องไม่ทำให้เกิด `RowsMoved` ค้างถาวร

- **[src/uat_tracker/excel/writer.py:111-114 (และ :434-444 `ComWorkbook.last_row/get`)] สแกนคอลัมน์ App ID ทีละ cell ผ่าน COM จนถึง `UsedRange` ทุกชีตในทุกการเขียน**
  - ปัญหา: `last = max(ws.last_row(), ...)` ใช้ `UsedRange` ซึ่งใน Excel จริงมักยาวเกินข้อมูลมาก เช่น จัด format ทั้งคอลัมน์หรือเคยพิมพ์ไว้ไกลแล้วลบ บางทีไปถึงแถว 1,048,576 แต่ละ `ws.get` คือ COM round-trip 1 ครั้งที่ห่อด้วย `_busy_retry`
  - สถานการณ์จริง: ถ้า UsedRange ยาวถึงแถวล้าน การเขียนแต่ละครั้งจะเรียก COM เป็นล้านครั้งต่อชีต ใช้เวลาหลายนาทีถึงหลายสิบนาที ถ้าเกิดตอน `flush` บน UI thread ตอนสลับโปรเจกต์ (`ui/app.py:144`) แอปจะค้าง ส่วน sync thread ก็ถือ lock ไว้นาน โค้ดเดิมก่อนแก้ไม่มีการสแกนทั้งคอลัมน์แบบนี้
  - แนวทางแก้: อ่านคอลัมน์ ID ทีละก้อนด้วย `Range(Cells(2,c), Cells(last,c)).Value` ครั้งเดียว (เพิ่มเมธอด `column_values` ให้ ComWorkbook/OpenpyxlWorkbook) และ/หรือหาแถวสุดท้ายจริงด้วย `Cells(Rows.Count, c).End(xlUp)` แทน UsedRange

## 🟢 Suggestion (ทางเลือก)

- **[src/uat_tracker/excel/writer.py:139-140, 162-177] การเขียนครั้งแรกที่ยังไม่มีคอลัมน์ App ID ยังไม่ได้รับการป้องกันแบบ T1** — ชีตที่ไม่มีคอลัมน์ ID จะข้ามการตรวจ แล้วเติม ID ตามเลขแถวจากการอ่านครั้งล่าสุด ทำให้การตรวจ `_id_text(...) != w.item_id` ที่บรรทัด 183 ผ่านเสมอ ซึ่งเป็นรูปแบบเดียวกับบั๊กเดิม แต่โอกาสเกิดแคบ เพราะ `pull()` อ่านไฟล์ก่อนเขียนทุกครั้ง จะเกิดได้ก็ต่อเมื่อ Excel ที่เปิดอยู่มีข้อมูลต่างจากไฟล์บนดิสก์ ถ้าต้องการป้องกันด้วย อาจเทียบ fingerprint/alt_key ของแถว (เช่น Application/No. หรือ Group/Task No) ก่อนเติม ID ครั้งแรก หรือบังคับให้ `init_ids` รันก่อน flush แรก
- **[src/uat_tracker/services/sync.py:207-212] ยังไม่มี test สำหรับ ROW ที่ค้างสถานะ 'writing' และ attempts > 0** — หลัง T2A รายการแบบนี้จะถึงคิวทันที ถ้าโปรเซสตายหลังบันทึกแถวใหม่ลง Excel แต่ก่อน `pull(force=True)` รอบถัดไปจะพึ่ง `pull()` ที่บรรทัด 222 ให้ล้าง ROW pending (ผ่าน `_pending_matches`) ก่อน ถ้าล้างไม่ได้จะเกิดแถวซ้ำ ควรเพิ่ม test: ROW 'writing' attempts=2 ที่เขียนลงไฟล์ไปแล้ว ต้องไม่ถูก insert ซ้ำ
- **[tests/test_bugfixes.py:265-268, 346-349] test ใช้ `try: ... except RowsMoved: pass`** — test จะผ่านแม้ไม่ได้ raise แต่มี test เวอร์ชันเข้มกว่าครอบไว้แล้ว (`test_t1_deleted_last_task_last_group_header_below_raises` และ parametrize ใน `test_t1_tp_deleted_last_row_above_grand_total_raises`) จะใช้ `pytest.raises` ให้ตรงกันก็ได้
- **[tests/test_bugfixes.py:940] assertion เช็ก `Path(tmp_path.anchor).glob("evil*")`** (root ของไดรฟ์) แทบไม่มีความหมาย ชื่อไฟล์ใน `_evil_names` ชี้ไปที่ `tmp_path` และ `attachment_dir/..` อยู่แล้ว และที่บรรทัด 937-939 ก็ครอบไว้แล้ว ลบบรรทัดนี้ได้
- **[tests/test_bugfixes.py:398-400, 87] PEP 8**: ก่อน `def _edit_three_sheets` มีบรรทัดว่างแค่ 1 บรรทัด (E302 ต้อง 2) และบรรทัด 87 ยาว 121 ตัวอักษร (โปรเจกต์ใช้ 120)
- **[src/uat_tracker/excel/writer.py:92, 96, 103, 131] type hints ไม่ครบ**: พารามิเตอร์ `ws` / `book` ไม่มีชนิด อาจใช้ `Protocol` (เช่น `SheetLike` ที่มี `get`, `last_row`, `header_values`) หรืออย่างน้อย `Any` ให้ชัด ส่วน `sheets.py:89-105` ใช้ `values: dict` ควรเป็น `Mapping[str, object]`
- **[src/uat_tracker/excel/sheets.py:82-115, src/uat_tracker/services/normalize.py] ขอบเขต**: การย้ายกฎ "เป็นรายการไหม" จาก normalize ไปไว้ที่ `sheets.py` เกินกว่าที่ tasks.md ระบุ แต่จำเป็นเพื่อให้ writer ใช้กฎเดียวกับ importer และมี test ยืนยันว่าพฤติกรรมเหมือนเดิม (`test_t1_rules_*`, `test_t1_normalize_keeps_same_rows_as_old_rules`) ยอมรับได้ ควรแจ้งผู้ใช้ไว้ในสรุปงาน (ดูข้อ 🟡 แรกเรื่อง `app_id` ประกอบ)

## ความครอบคลุมของ test ต่อ acceptance
| Task | Acceptance ตาม tasks.md | Test |
|---|---|---|
| T1 | แทรกแถวหลังอ่าน → ไม่ลงผิดแถว, ID ไม่เปลี่ยน, test เดิมผ่าน | ครบ: `test_t1_moved_rows_raise_and_nothing_changes` (10 แบบ), `test_t1_insert_top_reported_case_...`, `test_t1_sync_rows_moved_is_retry_then_written_to_right_row`, กรณี copy/blank/unknown ID, กรณีลบแถวสุดท้าย |
| T2A | 'writing' + attempts > 0 ถูกลองใหม่, คง first_failed_at | ครบ: `test_t2_interrupted_writing_*` (3 test) |
| T2B | ROW failed → field fail พร้อมเหตุผล; ROW ที่ยังรอเขียน → ข้าม | ครบ: `test_t2_failed_new_row_fails_its_field_edits`, `test_t2_queued_new_row_still_carries_field_edits` |
| T3 | ไม่ recursion, ไม่เหลือโปรเจกต์ก็ไปหน้า no-project | ครบ: `test_t3_*` (2 test) |
| T4 | entry นอก attachment_dir → ClearError | ครบ: 6 แบบ + ไฟล์ปกติที่อยู่ในโฟลเดอร์ย่อย |
| T5 | alert ใช้ PMO email ล่าสุด | ครบ: ทั้ง alert แบบ blocked และ error |
| T6 | ลบ import | ครบ + ตรวจด้วย AST ว่าไม่มี import ที่ไม่ใช้เหลืออยู่ |

## ความปลอดภัย
- T4: ป้องกัน path traversal ได้ถูกต้อง ตรวจทุก entry ก่อนแก้ DB (ไม่เกิดการ restore แค่บางส่วน) ส่วน symlink ที่อยู่ใน attachment_dir ถูก `resolve()` แล้วเทียบกับ root จึงปลอดภัย
- ไม่พบ secret, อีเมล หรือชื่อจริงของบริษัทในโค้ดหรือ test ที่เปลี่ยน (test ใช้ `@example.test`)

## ข้อสังเกตเพิ่มจาก qa-tester ระหว่างทดสอบ T7 (24-Nov-2026) — ยังไม่ได้ทำ รอผู้ใช้ตัดสิน
- 🟢 [src/uat_tracker/excel/writer.py:120] ถ้าผู้ใช้ลบเนื้อหาแถวใน Excel ที่เปิดอยู่ *หลัง* การอ่าน แต่ App ID ยังอยู่ (`cur == item_id` -> `continue` ข้าม `_holds_item`) writer จะยังเขียนค่า (เช่น Remark) ลงแถวว่างนั้น ไม่ค้างถาวร: import รอบถัดไปตั้งรายการเป็น inactive ตามปกติ แต่ค่าที่เขียนจะค้างอยู่ในแถวที่ไม่ใช่รายการแล้ว

---

# รอบ T7–T8 (24-Nov-2026)

ขอบเขต: T7 (`src/uat_tracker/excel/sheets.py` `_content` / `is_item_row` / `is_defect_item`, `src/uat_tracker/services/normalize.py` `normalize_uat_plan`) และ T8 (`src/uat_tracker/excel/writer.py` `_moved_rows`, `ComSheet.last_data_row` / `column_values`, `OpenpyxlSheet.last_data_row` / `column_values`, `XL_UP`) รวม test `test_t7_*` และ `test_t8_*` ใน `tests/test_bugfixes.py` ไม่ได้รัน pytest ซ้ำ (ผู้สั่งงานแจ้งว่าผ่าน 388 passed)

## สรุป
**ผ่าน (มีข้อควรแก้ 1 ข้อ)** — 🔴 0, 🟡 1, 🟢 5

สถานะของ 🟡 สองข้อจากรอบก่อน:
- **🟡 ข้อ 1 (กฎ "เป็นรายการไหม" ไม่ตรงกันเรื่อง `app_id`) ปิดแล้ว** `sheets.py:90-121` ตัด `APP_ID_KEY` ออกใน `is_item_row` และ `is_defect_item` ทำให้ importer (`normalize.py:81`, `:180`) กับ writer (`writer.py:101`) ใช้ชุดคอลัมน์เดียวกัน ไล่ดูทุกกรณีแล้ว: แถว UAT ที่มีแค่ App ID ที่ไม่รู้จัก และแถว Defect ที่มีแค่ No. + App ID จะไม่ถูก import จึงไม่อยู่ใน `expected` ของ writer และไม่มีทางเกิด `RowsMoved` ค้างถาวรแบบเดิม ส่วน Test Progress ตรงกันอยู่แล้วตั้งแต่ก่อนแก้ (`is_test_progress_item` ดูแค่ Application/No.) มี test ยืนยันทั้งระดับกฎ, normalize, import และ sync (`test_t7_rules_ignore_app_id`, `test_t7_sync_writes_with_only_unknown_app_id_rows`, `test_t7_qa_*`)
- **🟡 ข้อ 2 (สแกน App ID ทีละ cell ไปจนสุด UsedRange) ปิดแล้วในกรณีหลัก** `_moved_rows` (`writer.py:112-116`) อ่านคอลัมน์ ID ด้วย `Range(...).Value` ครั้งเดียว และหาแถวสุดท้ายด้วย `Cells(Rows.Count, c).End(xlUp)` (`writer.py:440-455`) จัดการกรณี cell สุดท้ายของชีต (แถว 1,048,576) มีค่า และกรณีอ่านแค่ 1 cell แล้ว COM คืนค่าเป็น scalar ถูกต้อง `last` ยังเอา `max(expected)` มาเทียบด้วย จึงอ่านครอบทุกแถวที่คาดไว้เสมอ ผล RowsMoved ใน openpyxl เหมือนเดิมทุกกรณี (`last_data_row` = `max_row` = `last_row` เดิม) ส่วน COM ให้ผลเหมือนเดิมยกเว้นกรณีแถวถูกซ่อนโดยไม่ใช้ filter (ดู 🟡 ด้านล่าง)

T7 ไม่ทำให้แถว UAT ที่ควรเป็นรายการหายไป: แถวที่ reader ส่งมา (`reader.py:103` ตัดแถวที่ทุกคอลัมน์ว่างตาม `is_blank` ซึ่งใช้ `cell_text` ตัวเดียวกัน) แล้วถูก `is_item_row` ตัดเพิ่ม มีแค่แถวที่**ทุกคอลัมน์ยกเว้น App ID ว่าง** เท่านั้น ส่วนแถวที่ app สร้างเองจะไม่ตกอยู่ในกรณีนี้ เพราะ `items.py:81-145` บังคับให้กรอก Task name/Responsible, Application หรือ System/Issue detail ก่อนสร้างรายการ

## 🔴 Critical (ต้องแก้ก่อน merge)
- ไม่มี

## 🟡 Warning (ควรแก้)

- **[src/uat_tracker/excel/writer.py:440-455 (`ComSheet.last_data_row`) ร่วมกับ :112] `End(xlUp)` ข้ามแถวที่ซ่อนอยู่ทุกแบบ ไม่ได้ข้ามเฉพาะแถวที่ AutoFilter ซ่อน แต่ fallback ไป UsedRange จะทำเฉพาะตอน `FilterMode`**
  - ปัญหา: ใน Excel การกด Ctrl+↑ หรือเรียก `Range.End` จะข้าม cell ในแถวที่ซ่อนอยู่ ไม่ว่าจะซ่อนเอง (Hide) หรือยุบ Group/Outline (ปุ่ม −) หรือซ่อนด้วย filter ของ Table (ListObject) ซึ่ง `Worksheet.FilterMode` อาจไม่ได้สะท้อนสถานะนี้ แต่โค้ดถือว่ามีแค่ AutoFilter ของชีตที่ทำให้ End หยุดก่อนถึงแถวสุดท้ายจริง ส่วน test QA จำลอง End ให้ข้ามแถวซ่อนเฉพาะตอน `filter_on` (`tests/test_bugfixes.py:1515-1517`) จึงไม่ครอบกรณีนี้
  - สถานการณ์จริง: ชีต UAT Plan มักยุบ group ไว้ ถ้ากลุ่มสุดท้ายถูกยุบ (แถวท้ายซ่อนอยู่) แล้วผู้ใช้ sort หรือ insert แถวจนรายการ X ไปอยู่ในแถวที่ซ่อนอยู่และต่ำกว่า `max(expected)` ส่วนแถวที่ X เคยอยู่กลายเป็นแถวใหม่ที่มีเนื้อหาแต่ยังไม่มี ID (`_holds_item` = True) ตอนนี้ `last_data_row` จะคืนแถวที่มองเห็นแถวสุดท้าย ทำให้ `found` ไม่มี X และไม่มี problem จากนั้น `apply_writes` (`writer.py:175-178`) จะเขียน ID ของ X ลงแถวผิด แล้วเขียนค่าตามมา ซึ่งเป็นบั๊กเดียวกับ T1 แต่เกิดได้เฉพาะกรณีแคบนี้ (โค้ดก่อน T8 ใช้ UsedRange จึงไม่เจอปัญหานี้)
  - แนวทางแก้ (เลือกอย่างใดอย่างหนึ่ง): (ก) หาแถวสุดท้ายด้วย `ws.Columns(col).Find("*", LookIn=xlFormulas(-4123), SearchOrder=xlByRows(1), SearchDirection=xlPrevious(2))` วิธีนี้ (LookIn=xlFormulas) เจอ cell ในแถวที่ซ่อนด้วย และไม่ต้องแยกกรณี FilterMode (ข) หรือคง End(xlUp) ไว้ แต่ fallback ไป UsedRange เมื่อชีตมีแถวซ่อนอยู่ระหว่าง End กับ UsedRange (เช่น ตรวจ `Range(Rows(e+1), Rows(ur_last)).SpecialCells(xlCellTypeVisible)` หรือดู `ws.AutoFilterMode` / `ListObjects(...).ShowAutoFilter` / `Outline`) ควรเพิ่ม test ที่ fake End ให้ข้ามแถวซ่อนเสมอ และทดสอบกับ Excel จริงหนึ่งครั้ง (ทั้งยุบ group และใช้ Table filter)

## 🟢 Suggestion (ทางเลือก)

- **[src/uat_tracker/excel/writer.py:440-455] ควรทดสอบ smoke test บน Excel จริง** เพราะ test T8 ทั้งหมดใช้ fake COM ที่จำลองพฤติกรรมของ `End`, `FilterMode` และ `Range.Value` ตามที่เราเข้าใจ (เช่น End ในคอลัมน์ App ID ที่ถูกซ่อนทั้งคอลัมน์ หรือ `FilterMode` ตอนมี AutoFilter แต่ยังไม่ได้เลือกเงื่อนไข) ควรรัน `init_ids` + flush กับไฟล์ทดสอบใน Excel จริงหนึ่งรอบก่อนส่งงาน
- **[src/uat_tracker/excel/writer.py:443-445, 457-465] ตอน `FilterMode` ถ้า UsedRange ยาวถึงแถวล้าน `column_values` จะอ่าน `Range.Value` ขนาด ~1M cell ใน COM call เดียว** ใช้หน่วยความจำและเวลาหลายวินาที แต่ยังเร็วกว่าเดิมมาก (เดิม 1M round-trip) ยอมรับได้ ถ้าแก้ตาม 🟡 ทาง (ก) ข้อนี้จะหายไปด้วย
- **[src/uat_tracker/excel/sheets.py:111-121] พฤติกรรมเปลี่ยนที่ควรแจ้งผู้ใช้**: รายการเดิมที่แถวใน Excel ถูกล้างเนื้อหาจนเหลือแค่ App ID (ID ที่ระบบรู้จัก) ก่อนแก้ยังนับเป็น active ตอนนี้ import รอบถัดไปจะตั้งเป็น removed/inactive และ field ที่ค้างคิวจะ fail (`test_t7_qa_cleared_row_with_known_id_removed_like_deleted`) ถูกตาม tasks.md T7 แต่เป็นผลที่ผู้ใช้เห็นได้ ควรแจ้งไว้ในสรุปงาน
- **[src/uat_tracker/excel/sheets.py:90, 95, 100, 106, 111; writer.py:452 `-> list`] type hints**: `values: dict` / `-> dict` / `-> list` ควรระบุให้ชัดเป็น `Mapping[str, object]`, `dict[str, object]`, `list[object]` (ต่อเนื่องจากข้อ 🟢 รอบก่อน) และพารามิเตอร์ `ws` ใน `_moved_rows` / `_holds_item` ยังไม่มีชนิด ส่วนเมธอดใหม่ `last_data_row` / `column_values` มีชนิดแล้ว
- **[tests/test_bugfixes.py:1503-1505, 1497] PEP 8 ใน test**: ระหว่าง docstring กับ `def __init__` ของ `_QaComError` มีบรรทัดว่าง 2 บรรทัด (E303 ในคลาสใช้ได้ 1) และ `from uat_tracker.excel.sheets import APP_ID_KEY  # noqa: E402` อยู่กลางไฟล์ ควรย้ายไปไว้กับ import ด้านบน (บรรทัด 87 ที่ยาว 121 ตัวยังค้างจากรอบก่อน) นอกจากนี้ `normalize.py:75` ยังคำนวณ `tno` ก่อนเช็ก group header ซึ่งยังใช้ต่อที่บรรทัด 86-94 จึงไม่ใช่ตัวแปรที่ไม่ได้ใช้ ไม่ต้องแก้

## ความครอบคลุมของ test ต่อ acceptance (T7–T8)
| Task | Acceptance ตาม tasks.md | Test |
|---|---|---|
| T7 | แถวที่มีแค่ App ID / No. + App ID ไม่ถูก import, ไม่เกิด RowsMoved ค้าง, sync เขียนได้ | ครบ: `test_t7_rules_ignore_app_id`, `test_t7_normalize_skips_rows_with_only_app_id`, `test_t7_rows_with_only_unknown_app_id_not_imported`, `test_t7_sync_writes_with_only_unknown_app_id_rows` (flush 2 รอบ) + QA: orphan กลาง group, ย้ายแถวแล้ว retry, แถวที่ล้างแต่ ID ยังอยู่, Test Progress |
| T8 | column_values = อ่านทีละ cell; UsedRange 1M แถวไม่เรียก get ทีละ cell; ผล RowsMoved เหมือนเดิม | ครบตาม AC: `test_t8_column_values_same_as_cell_by_cell_*`, `test_t8_last_data_row_com`, `test_t8_com_scan_same_result_as_before_and_few_com_calls` (นับ COM call), QA: filter, แถวเดียว, แถว 1,048,576, busy retry — ขาดกรณีแถวที่ซ่อนโดยไม่ใช้ filter (ดู 🟡) |

## ความปลอดภัย
- ไม่พบ secret, อีเมล หรือชื่อจริงของบริษัทในโค้ดหรือ test ที่เปลี่ยนรอบนี้ ไม่มี input จากภายนอกเข้าไปใน COM call ใหม่ (มีแค่เลขแถว/คอลัมน์)

---

# รอบ T9 (24-Nov-2026)

ขอบเขต: T9 `src/uat_tracker/excel/writer.py` (ค่าคงที่ `XL_FORMULAS`/`XL_PART`/`XL_BY_ROWS`/`XL_PREVIOUS` :41-45, `ComSheet.last_data_row` :443-456 ที่ใช้ใน `_moved_rows` :116) และ test ใน `tests/test_bugfixes.py` ทั้งส่วน T8 ที่ปรับ (:1335-1840) และ `test_t9_*` / T9 QA (:1841 ถึงท้ายไฟล์) ไม่ได้รัน pytest ซ้ำ เพราะผู้สั่งงานแจ้งว่าผ่านแล้ว 466 passed ตรวจด้วย `git diff` ซึ่งยังปนกับงานรอบก่อนที่ยังไม่ commit

## สรุป
**ผ่าน (ต้อง smoke test กับ Excel จริงก่อน merge)**: 🔴 0, 🟡 1, 🟢 5

สถานะของ 🟡 จากรอบ T7–T8:
- **🟡 "End(xlUp) ข้ามแถวที่ซ่อนอยู่ทุกแบบ" ปิดแล้วสำหรับกรณีซ่อนแถวเองและยุบ Group/Outline** ตอนนี้ `last_data_row` (`writer.py:450-456`) ใช้ `Columns(col).Find("*", Cells(1,col), xlFormulas, xlPart, xlByRows, xlPrevious)` ตามทาง (ก) ที่แนะนำไว้ สิ่งที่ตรวจแล้วว่าถูกต้อง:
  - ลำดับ argument แบบ positional ตรงกับลายเซ็นของ `Range.Find` (What, After, LookIn, LookAt, SearchOrder, SearchDirection) และค่าคงที่ถูกต้อง (-4123, 2, 1, 2)
  - `After` คือ cell แรกของคอลัมน์และค้นแบบ xlPrevious จึงวนไปเริ่มที่แถว 1,048,576 ก่อน แถวสุดท้ายของชีตจึงถูกนับด้วย ไม่ต้องมีกรณีพิเศษแบบตอนใช้ End อีก ถ้าคอลัมน์มีแค่ header จะได้ row 1 เพราะ Excel ตรวจ cell `After` เป็นตัวสุดท้าย
  - ถ้าไม่พบอะไรเลย COM คืน `Nothing` ซึ่ง pywin32 แปลงเป็น `None` และโค้ดคืน 1 ให้ นอกจากนี้ `_moved_rows` ยังใช้ `max(..., max(expected))` (`:116`) จึงอ่านครอบทุกแถวที่คาดไว้เสมอ แม้ Find จะคืนค่าต่ำเกินจริง
  - ห่อ `_busy_retry` แล้ว การ retry ทั้งเมธอด (อ่าน FilterMode แล้วเรียก Find) ทำซ้ำได้โดยไม่มีผลข้างเคียง
  - ยังคงทางสำรอง FilterMode -> UsedRange ไว้ ซึ่งจำเป็น (ดู 🟡 ด้านล่าง) ลบ `XL_UP` และกรณีแถวสุดท้ายแยกออกครบแล้ว ไม่มี `.End(` เหลือใน writer
  - test ครอบคลุมตาม AC ของ T9 ครบ: ซ่อนแถวเอง/ยุบ group, sort หรือ insert จนรายการไปอยู่แถวที่ซ่อน, ไม่มีการแจ้งเตือนผิดเมื่อแถวอยู่ที่เดิม, ผล RowsMoved เท่ากับทาง openpyxl ทุก mutation, busy retry, คอลัมน์ App ID ถูกซ่อนทั้งคอลัมน์ และสูตรที่คืนค่า ""
- ส่วนที่ยังค้างคือ filter ของ Table (ListObject) ซึ่งยังไม่ได้ยืนยันกับ Excel จริง (ดู 🟡 ด้านล่าง)

**test T8 ที่ปรับไม่ได้อ่อนลง**
- `test_t8_com_scan_same_result_as_before_and_few_com_calls` (`tests/test_bugfixes.py:1510-1513`) ยังยืนยันว่า `used_range == 0`, อ่าน Range 2 ครั้งต่อชีต และไม่มีการอ่านทีละ cell ที่เปลี่ยนคือ `end_calls == sheets` กลายเป็น `find_calls == sheets and end_calls == 0` ซึ่งเข้มขึ้น
- `test_t8_last_data_row_com` (`:1461-1481`) ยังตรวจคอลัมน์ว่าง -> 1, FilterMode -> UsedRange และแถว 1,048,576 ครบ
- test QA ของ filter (`:1647-1690`) ยังตรวจว่าใช้ UsedRange และไม่เรียก End
- `_fake_find` (`:1361-1371`) assert What/LookAt/SearchOrder/SearchDirection/After ทุกครั้งที่ถูกเรียก ถ้าส่ง argument ผิดลำดับ test จะล้ม

**Excel จำค่าจาก Find ครั้งก่อน และต้องส่ง MatchCase ไหม**: ไม่ต้อง เพราะ `What="*"` เป็น wildcard ที่ตรงกับทุก cell ที่ไม่ว่าง ไม่ว่า MatchCase/MatchByte จะตั้งไว้อย่างไร LookIn/LookAt/SearchOrder/SearchDirection ก็ส่งเองครบ ค่าที่ Excel จำไว้จึงเปลี่ยนผลไม่ได้ ส่วน `SearchFormat` ที่ไม่ได้ส่งจะใช้ค่า default False รูปแบบที่ผู้ใช้ตั้งไว้ใน Find dialog จึงไม่มีผล ผลกระทบในทางกลับกัน (ค่าที่เราส่งไปค้างอยู่ใน dialog ของผู้ใช้) ดู 🟢 ข้อ 1

**pywin32 dynamic dispatch กับ positional args**: ส่งแบบ positional ปลอดภัยกว่า เพราะ dynamic dispatch (`Dispatch(ObjectFromLresult(...))` ที่ `writer.py:336`) รองรับ keyword argument ได้ไม่แน่นอน ขณะที่ positional ใช้ได้ทั้งแบบ dynamic และ early-bound (makepy) การเรียก `ws.Columns(col)` แบบมี parameter ก็ใช้อยู่แล้วใน `hide_column` (`:470`) และ `Nothing` ถูกแปลงเป็น `None` ทั้งสองแบบ

## 🔴 Critical (ต้องแก้ก่อน merge)
- ไม่มี

## 🟡 Warning (ควรแก้)

- **[src/uat_tracker/excel/writer.py:42, 450-456; tests/test_bugfixes.py:2047-2055, 2058-2069, 2072-2095] test สมมติพฤติกรรมของ Table (ListObject) filter ในแง่ดีที่สุด และยังไม่ได้ยืนยันกับ Excel จริง**
  - ปัญหา: `T9QaWs.find` สมมติว่า Find แบบ `LookIn=xlFormulas` มองเห็นแถวที่ filter ของ Table ซ่อนไว้ (`table_hidden` ถูกรวมเข้า `hidden` ซึ่ง fake ให้ xlFormulas มองเห็นทั้งหมด) แต่พฤติกรรมที่มีรายงานกันใน Excel คือ xlFormulas เห็นแถวที่ซ่อนเองหรือยุบ group ได้ ส่วนแถวที่**ถูก filter กรองออก** Find จะข้ามไปแม้ใช้ xlFormulas ซึ่งตัวโค้ดเองก็ยอมรับข้อนี้ จึงคงทางสำรอง FilterMode -> UsedRange ไว้สำหรับ AutoFilter ของชีต ถ้า filter ของ Table ทำงานแบบเดียวกัน **และ** `Worksheet.FilterMode` ไม่เป็น True ตอนกรองใน Table (ซึ่ง `test_t9_qa_table_filter_with_filter_mode_false_uses_find_only` ก็สมมติไว้แบบนี้) Find จะคืนแถวสุดท้ายที่มองเห็น และบั๊กของรอบ T7–T8 จะยังเหลือในกรณีนี้ คอมเมนต์ที่ `:42` ("also searches hidden rows") จึงเขียนเกินจริงสำหรับแถวที่ถูกกรองออก
  - สถานการณ์จริง: ข้อมูล UAT Plan/Defect อยู่ใน Table และผู้ใช้กรองจนแถวท้ายถูกซ่อน จากนั้น sort หรือ insert จนรายการ X ไปอยู่ในแถวที่ถูกกรองออกและต่ำกว่า `max(expected)` ขณะที่แถวเดิมของ X กลายเป็นรายการใหม่ที่ยังไม่มี ID ผลคือไม่เกิด `RowsMoved` และ `apply_writes` จะเขียน ID ของ X ลงแถวผิด (แบบ T1) ทั้งนี้ test 3 ตัวที่อ้างไว้ข้างต้นจะยังผ่าน เพราะ fake ทำตามสมมติฐานนี้ จึงจับกรณีนี้ไม่ได้
  - แนวทางแก้: (1) smoke test กับ Excel จริงตามรายการด้านล่างก่อน ถ้า Find เห็นแถวที่ Table กรองออก หรือ `ws.FilterMode` เป็น True ตอนกรอง Table ก็ไม่ต้องแก้โค้ด แก้แค่คอมเมนต์ (2) ถ้าไม่เป็นอย่างนั้น ให้ขยายเงื่อนไขทางสำรองเป็น `self.ws.FilterMode or any(lo.AutoFilter is not None and lo.AutoFilter.FilterMode for lo in self.ws.ListObjects)` แล้วใช้ UsedRange และปรับ fake ให้ Find มองไม่เห็นแถวที่ถูกกรองออกตามพฤติกรรมจริง พร้อมเพิ่ม test สำหรับทางสำรองนี้

## 🟢 Suggestion (ทางเลือก)

- **[src/uat_tracker/excel/writer.py:454-455] Find เปลี่ยนค่าที่ Find dialog (Ctrl+F) ของผู้ใช้จำไว้** ตามเอกสาร Excel ค่า LookIn/LookAt/SearchOrder ที่ส่งให้ `Range.Find` จะถูกบันทึกเป็นค่าตั้งต้นของ dialog และ ComWorkbook เกาะกับ Excel ที่ผู้ใช้เปิดอยู่ (`:336`) หลัง sync ทุกครั้ง dialog จึงกลับเป็น Look in = Formulas, ไม่ติ๊ก Match entire cell, Search = By Rows ซึ่งตรงกับค่า default ของ Excel อยู่แล้ว ผู้ใช้ส่วนใหญ่จึงไม่รู้สึกอะไร แต่ใครที่ตั้ง "Look in: Values" หรือ "Match entire cell contents" ไว้จะเห็นค่าถูกรีเซ็ตโดยไม่รู้ตัว ยอมรับได้ ถ้าต้องการให้แจ้งไว้ในสรุปงาน ไม่ต้องส่ง MatchCase เพิ่ม (ดูข้างบน)
- **Smoke test กับ Excel จริง (จำเป็นก่อน merge เพราะ test ทั้งหมดใช้ fake COM)** ใช้ไฟล์สำเนา รัน `init_ids` แล้ว flush ผ่าน ComWorkbook โดย sort ให้ X ไปอยู่ในแถวที่ซ่อนใต้ `max(expected)` และให้แถวเดิมของ X เป็นรายการใหม่ที่ยังไม่มี ID ผลที่ต้องได้คือ `RowsMoved` และไฟล์ไม่ถูกแก้ ทำทีละกรณี:
  1. **Table (ListObject) filter** เสี่ยงที่สุด (ดู 🟡) ให้ดูทั้งค่า `ws.FilterMode` และแถวที่ Find คืนมา
  2. **AutoFilter ของชีตที่กำลังกรองอยู่** ต้องใช้ UsedRange เพราะ FilterMode เป็น True
  3. **AutoFilter ที่เปิดไว้แต่ยังไม่ได้เลือกเงื่อนไข** ต้องใช้ Find และได้แถวสุดท้ายจริง
  4. **ยุบ Group/Outline ช่วงท้าย และซ่อนแถวเอง** สองกรณีที่ T9 ตั้งใจแก้
  5. **คอลัมน์ App ID ถูกซ่อนทั้งคอลัมน์** (`hide_column` ซ่อนเสมอ) ร่วมกับข้อ 4
  6. **ID อยู่แถว 1,048,576 / คอลัมน์มีแค่ header / คอลัมน์ว่างทั้งหมด**
  7. **ชีตที่ protect ไว้และ cell ติ๊ก Hidden** (xlFormulas อาจมองไม่เห็น cell ที่ FormulaHidden) ถ้าระบบไม่รองรับชีตที่ protect อยู่แล้ว ข้ามข้อนี้ได้
- **[src/uat_tracker/excel/writer.py:42, 445-449] คอมเมนต์ของ `XL_FORMULAS` และ docstring** ควรเขียนให้ชัดว่า xlFormulas เห็นแถวที่ซ่อนเองหรือยุบ group แต่ไม่รับประกันแถวที่ filter กรองออก (เป็นเหตุผลที่ต้องมีทางสำรอง FilterMode) เพื่อกันไม่ให้ใครมาลบทางสำรองนี้ภายหลัง docstring ที่เขียนว่า Find ไม่ข้ามแถวที่ "any filter" ซ่อน ก็ขัดกับเหตุผลนี้
- **[tests/test_bugfixes.py:1353-1358, 1934-1936] test ที่ไม่ได้ใช้แล้วหรือเปราะ**: `_FakeComCell.End` ไม่มีที่ไหนเรียกแล้ว (ที่ `:1857` ใช้ `_QaCell.End` ของ QA) จึงลบได้ ส่วน `test_t9_no_end_xlup_left_in_writer` ตรวจด้วยการค้นข้อความใน source ซึ่งเปราะ (แค่มีคอมเมนต์ที่มี `.End(` ก็ล้ม) และ test อื่นที่ assert `end_calls == 0` ครอบคลุมเรื่องนี้อยู่แล้ว
- **[src/uat_tracker/excel/writer.py:459, 608] type hints / PEP 8**: `-> list` ควรเป็น `list[object]` (ค้างจากรอบก่อน) ส่วนโค้ดใหม่ใน writer.py และ test T9 ไม่มีบรรทัดเกิน 120 ตัว การตั้งชื่อและการเว้นบรรทัดถูกต้องตาม PEP 8 ค่าคงที่ `XL_*_DOC` ใน test (`:1941`) ซ้ำกับใน writer แต่ตั้งใจให้เป็นค่าอ้างอิงอิสระตามเอกสาร Excel จึงยอมรับได้

## ความครอบคลุมของ test ต่อ acceptance (T9)
| AC ตาม tasks.md | Test |
|---|---|
| แถวสุดท้ายรวมแถวที่ซ่อนทุกแบบ, ไม่พบ -> 1 | `test_t9_last_data_row_counts_hidden_rows`, `test_t9_last_data_row_all_rows_hidden_and_empty_column`, QA กรณีซ่อนทั้งคอลัมน์และสูตรที่คืน "" |
| อ่าน COM จำนวนน้อย (ไม่สแกน UsedRange) | นับ call ได้ `{"filter": 1, "find": 1}`, `used == 0`, อ่าน Range ครั้งเดียว |
| ห่อ `_busy_retry` | `test_t9_find_busy_retried_and_wrapped`, `test_t8_qa_busy_*` (เพิ่ม "find") |
| ผล RowsMoved เหมือนเดิมทุกกรณี | `test_t9_hidden_rows_same_rows_moved_result`, `test_t9_qa_hidden_column_and_table_filter_same_rows_moved_result` × MUTATIONS |
| sort/insert เข้าแถวที่ซ่อน -> RowsMoved | `test_t9_sorted_into_hidden_row_detected`, `test_t9_row_inserted_pushes_item_into_hidden_row` |
| test T1/T7/T8 เดิมผ่าน | ผู้สั่งงานแจ้งว่า 466 passed |
| ยังไม่ครอบคลุม | พฤติกรรมของ Excel จริงเมื่อมี Table filter (ดู 🟡 และรายการ smoke test) |

## ความปลอดภัย
- ไม่พบ secret, อีเมล หรือชื่อจริงของบริษัทในโค้ดหรือ test ที่เปลี่ยนรอบนี้ argument ที่ส่งเข้า COM มีแค่ค่าคงที่และเลขคอลัมน์ ไม่มี input จากผู้ใช้

---

# รอบ T10–T12 (24-Nov-2026)

ขอบเขต: การเปลี่ยนแปลงหลังรอบ T9 ได้แก่ `src/uat_tracker/excel/writer.py` (T10: ทางสำรอง `FilterMode or ProtectContents` -> UsedRange ที่ `:503` พร้อมคอมเมนต์ `:43-46` และ docstring `:494-502`; T11: `_row_holds_item` `:141-148`, `_moved_rows(..., written)` `:151-180`, `_check_rows_unmoved(..., writes)` `:183-202`, `row_values` ใน ComSheet `:522-529` และ OpenpyxlSheet `:674-675`; T12: `SheetLike`/`BookLike` `:90-115` และ type hints), `sheets.py` (`Mapping`), `normalize.py`/`sync.py` (PEP 8) และ `tests/test_bugfixes.py` (test T10/T11/T12/QA, fake ที่ปรับ, test ที่ลบหรือเปลี่ยนชื่อ) ตรวจจาก `git diff` ซึ่งยังปนกับงานรอบก่อนที่ยังไม่ commit ไม่ได้รัน pytest หรือ smoke test ซ้ำ เพราะผู้สั่งงานแจ้งว่า pytest ผ่าน 548 passed, flake8 ไม่มีข้อผิดพลาด และ smoke test กับ Excel จริงผ่าน 24/24 และ 4/4

## สรุป
**ผ่าน พร้อม commit**: 🔴 0, 🟡 0, 🟢 4

**🟡 Table filter จากรอบ T9 ปิดแล้ว**
- smoke test กับ Excel จริงยืนยันว่าตอน Table (ListObject) กรองอยู่ Excel ตั้ง `ws.FilterMode=True` ด้วย จึงเข้าทางสำรอง UsedRange (`writer.py:503-505`) ที่ครอบทุกแถวที่มีค่า กรณีที่ห่วงไว้ในรอบ T9 (Find ข้ามแถวที่ Table กรองออก โดยที่ FilterMode เป็น False) จึงไม่เกิดใน Excel จริง ไม่ต้องตรวจ `ListObjects(...).AutoFilter.FilterMode` เพิ่ม
- fake ปรับให้ตรงกับ Excel จริงแล้ว คือภายใต้ filter `Find` คืนแถว 1 ซึ่งเชื่อถือไม่ได้ (`tests/test_bugfixes.py:1594-1597`) ดังนั้นถ้ามีใครลบทางสำรอง FilterMode ภายหลัง test จะล้ม ส่วน `test_t9_qa_filtering_table_sets_filter_mode_uses_used_range` (`:2110`) ยืนยันว่าไม่ได้เรียก Find และอ่านได้ถึงแถว 12 ของ UsedRange
- test ที่เคยสมมติว่า "Table กรองอยู่แต่ FilterMode=False" เปลี่ยนเป็น `test_t9_qa_table_filter_buttons_without_criteria_uses_find_only` (`:2099`) ซึ่งตรงกับพฤติกรรมจริง (แสดงปุ่ม filter แต่ยังไม่ได้เลือกเงื่อนไข) และยังซ่อนแถว 10 ด้วยการยุบ group แล้วยืนยันว่า Find หาเจอ ไม่ได้ทำให้ test อ่อนลง
- กรณีที่ smoke test พบเพิ่ม (ชีตที่ protect ไว้และ cell ตั้ง FormulaHidden ทำให้ Find คืน None) แก้ด้วย `or self.ws.ProtectContents` แล้ว มี test ครบทั้งกรณี Find คืน None, sort จนรายการไปอยู่แถวที่ซ่อน, ไม่แจ้งเตือนผิดเมื่อแถวอยู่ที่เดิม, ชีตที่ไม่ได้ protect ยังใช้ Find, busy retry และ COM error อื่นที่ไม่ควร retry (`:2594-2700`) คอมเมนต์ที่ `:43-46` และ docstring ที่ `:494-502` เขียนตรงกับพฤติกรรมจริงแล้ว และมีคำเตือนว่าห้ามลบทางสำรอง ซึ่งปิดข้อ 🟢 เรื่องคอมเมนต์จากรอบ T9 ด้วย

**T11: writer กับ importer ตัดสินตรงกัน ไม่พบกรณีที่ทำให้ RowsMoved ค้างถาวร**
- RowsMoved จะค้างถาวรได้ก็ต่อเมื่อ writer บอกว่าแถวนี้ "ไม่ใช่รายการ" แต่ importer หลังอ่านไฟล์ใหม่ยังถือว่าเป็นรายการที่ active อยู่ ไล่ดูแล้วทั้งสองฝั่งใช้ `is_item_row` ตัวเดียวกัน (`sheets.py`) และไม่นับ App ID ทั้งคู่ ฝั่ง importer มีเงื่อนไขที่เข้มกว่า writer แค่สองอย่าง คือข้ามแถวที่อยู่เหนือ group A และกรณีที่ openpyxl อ่าน cached value ของสูตรได้ None ทั้งสองอย่างทำให้ importer ตัดสินว่า "ไม่ใช่รายการ" บ่อยกว่า writer ซึ่งเป็นทิศทางที่ไม่ทำให้วนค้าง ค่าที่ COM อ่านได้ (datetime ของ pywintypes, error เป็นเลข int, 0.0, False) กับค่าที่ openpyxl อ่านได้ (data_only) ให้ผล `cell_text` ว่าง/ไม่ว่างตรงกันทุกชนิด ส่วน merged cell ทั้งสองฝั่งคืน None ที่ cell ที่ไม่ใช่มุมบนซ้ายเหมือนกัน และชุดคอลัมน์ก็มาจาก `map_headers` กับ `header_label` ชุดเดียวกัน (มีข้อจำกัดเดิมที่ writer อ่าน header แค่ 80 คอลัมน์ แต่เป็นข้อจำกัดที่มีอยู่ก่อนแล้ว)
- มี `test_t11_qa_writer_decision_matches_importer` (`tests/test_bugfixes.py:2397-2418`) ที่บันทึกแถวที่เหลือค่าค้างแต่ละแบบลงไฟล์ แล้วยืนยันว่า `writer_ok is item_row` และ `active is item_row` ซึ่งเป็น test ที่ตรงกับความเสี่ยงนี้โดยตรง ส่วน `test_t11_qa_sync_new_app_and_edit_on_cleared_anchor` ยืนยันผลจนจบทั้งรอบ ได้แก่ retry, ไม่บันทึกไฟล์, หลังอ่านไฟล์ใหม่ TP-35 เป็น inactive, edit fail และแถวใหม่ถูกเขียน
- ผลต่อ ROW_FIELD/anchor: `written` ไม่รวม ROW_FIELD (`writer.py:196`) ดังนั้น ROW write อย่างเดียวจะไม่ทำให้ anchor ถูกตรวจ (ดู 🟢 ข้อ 2) ส่วนถ้า anchor มี cell edit อยู่ในรอบเดียวกันด้วย จะเกิด RowsMoved ก่อนแตะ workbook และแถวใหม่ก็จะไม่ถูกแทรกด้วย (`test_t11_qa_edit_and_row_write_anchor_cleared_nothing_written`) ถูกต้อง
- จำนวน COM call เป็นไปตาม AC คืออ่าน `row_values` 1 ครั้งต่อรายการที่มีการเขียน และอ่านเฉพาะแถวที่ App ID ตรง (`:171`) ไม่ได้สแกนทุกแถว (`test_t11_com_one_row_read_per_written_item`) `row_values` ห่อ `_busy_retry` แล้ว และจัดการกรณีอ่าน 1 cell แล้วได้ scalar ถูกต้อง

**T12 ไม่เปลี่ยนพฤติกรรม**: `Protocol` ใช้แค่ตอนตรวจชนิด (ไม่ได้ใช้ `runtime_checkable`) และไฟล์มี `from __future__ import annotations` การเปลี่ยน `dict` เป็น `Mapping` ใน `sheets.py` และการแยกบรรทัดใน `normalize.py:168-170` / `sync.py:341` ก็เป็นแค่เรื่องชนิดและรูปแบบ เพิ่ม `_id_text` เพื่อลดโค้ดซ้ำโดยให้ผลเหมือนเดิมทุกประการ

**test ไม่อ่อนลง**: ลบ `_FakeComCell.End` และ `test_t9_no_end_xlup_left_in_writer` ตามที่แนะนำไว้ (test อื่นยัง assert ว่า `end_calls == 0` / `calls` ไม่มี "end") ลบ assertion ที่ตรวจ root drive ใน T4 แล้ว (ยังมี assertion ที่ตรวจ `tmp_path` กับ `attachment_dir/..` อยู่) ไม่เหลือ `try/except RowsMoved: pass` แล้ว (`:1524` เป็น helper ที่คืนข้อความเพื่อนำไปเทียบ ส่วน `:2414` ใช้ตั้ง flag แล้วนำไป assert) เพิ่ม `test_t2_interrupted_row_writing_already_in_file_not_inserted_twice` (`:1098`) ครบตาม 🟢 ของรอบแรก fake `FakeComWs` ของ T8 เพิ่ม `ProtectContents=False` เป็น attribute ธรรมดา ทำให้ test เดิมยังเข้าทาง Find เหมือนเดิม

## 🔴 Critical (ต้องแก้ก่อน merge)
- ไม่มี

## 🟡 Warning (ควรแก้)
- ไม่มี

## 🟢 Suggestion (ทางเลือก)

- **[src/uat_tracker/excel/writer.py:503-505] ชีตที่ protect ไว้ใช้ UsedRange เสมอ และในทางปฏิบัติก็เขียนไม่ได้อยู่แล้ว** ถ้า UsedRange ยาวถึงแถวล้าน `column_values` จะอ่านราว 1M cell ใน COM call เดียว (ข้อเดียวกับกรณี filter ในรอบ T7–T8 ซึ่งยอมรับได้) นอกจากนี้ถ้าชีตที่ protect มี cell ที่ lock ไว้ `set_text` / `insert_row` / `hide_column` จะล้มด้วย COM error แล้วกลายเป็น `WorkbookBusy` ที่ retry จนครบรอบแล้วแจ้ง error ข้อนี้มีมาก่อน T10 ไม่ได้เกิดจากรอบนี้ ควรแจ้งผู้ใช้ไว้ในสรุปงานว่าระบบยังไม่รองรับการเขียนลงชีตที่ protect ไว้ (ยกเว้น cell ที่ unlock หรือใช้ UserInterfaceOnly)
- **[src/uat_tracker/excel/writer.py:196, 266-285; tests/test_bugfixes.py:2460] ROW write ไม่ได้ตรวจว่า anchor ยังเป็นรายการอยู่หรือไม่** ถ้าแถวของ anchor ถูกล้างเนื้อหาหลังการอ่าน (App ID ยังค้างอยู่) และรอบนั้นมีแค่ ROW write แถวใหม่จะถูกแทรกใต้แถวว่างนั้น หลังอ่านไฟล์ใหม่ anchor จะเป็น inactive และแถวใหม่จะ active ตามปกติ ข้อมูลไม่ผิดและไม่ค้าง แค่ตำแหน่งของแถวใหม่อาจไม่สวย ส่วนกรณี spec ที่ไม่ได้ insert (`:278-283`) ก็อาจเติม ID ใหม่ทับ App ID เก่าที่ค้างอยู่ในแถวที่ถูกล้างไปแล้ว ซึ่งผลหลังอ่านไฟล์ใหม่ก็ยังถูกต้อง พฤติกรรมนี้ตั้งใจไว้และมี test กำกับ (`test_t11_qa_row_write_only_is_not_a_cell_write`) ยอมรับได้ ถ้าต้องการเข้มขึ้น ให้ใส่ `w.excel_row` ของ ROW write ที่เป็น anchor ลงในชุดที่ต้องตรวจด้วย
- **[src/uat_tracker/excel/writer.py:134-148] `_holds_item` กับ `_row_holds_item` ใช้กฎเดียวกันแต่อ่าน cell คนละวิธี** (อ่านทีละ cell กับอ่านเป็น Range ครั้งเดียว) ให้ `_holds_item` เรียก `_row_holds_item` ได้เลย จะเหลือโค้ดชุดเดียวและลด COM round-trip ในกรณีแถวที่ ID ไม่ตรงด้วย ผลการตัดสินเหมือนเดิม (test T7/T11 ครอบคลุมอยู่แล้ว)
- **[src/uat_tracker/excel/writer.py:532-560, 677-700; tests/test_bugfixes.py:2398] type hints / test**: เมธอดเดิมของ `ComSheet` / `OpenpyxlSheet` (`hide_column`, `set_text`, `insert_column` ฯลฯ) ยังไม่มี type hints แม้ `SheetLike` จะระบุชนิดไว้แล้ว ไม่กระทบการทำงาน และ `test_t11_qa_writer_decision_matches_importer` ข้ามกรณี `XL_DIV0` ซึ่งทำได้โดยบันทึก error เป็น string `"#DIV/0!"` ลงไฟล์ ไม่จำเป็นต้องทำ เพราะ `cell_text` ของทั้ง int และ str เป็นค่าที่ไม่ว่างอยู่แล้ว

## ความครอบคลุมของ test ต่อ acceptance (T10–T12)
| Task | Acceptance ตาม tasks.md | Test |
|---|---|---|
| T10 | ทุกกรณี last_data_row ครอบแถวที่มี ID จริง; sort/insert เข้าแถวซ่อน -> RowsMoved, ไฟล์ไม่ถูกแก้; Table filter | smoke test กับ Excel จริง 24/24 + 4/4 (ผู้สั่งงานแจ้ง), fake: `test_t10_*` (protect + FormulaHidden, filter ที่ Find เชื่อไม่ได้, busy), `test_t9_qa_filtering_table_sets_filter_mode_uses_used_range`, `test_t10_protected_formula_hidden_same_rows_moved_result` × MUTATIONS |
| T11 | ล้างแถว TP-12 หลังอ่าน -> ไม่เขียน; หลังอ่านไฟล์ใหม่ TP-12 inactive, edit fail, รายการอื่น written; อ่าน COM เฉพาะแถวที่เขียน | ครบ: `test_t11_cleared_after_read_*`, `test_t11_sync_cleared_after_read_retry_then_inactive`, `test_t11_com_one_row_read_per_written_item`, QA: leftover ทุกแบบ × 3 ชีต, writer ตรงกับ importer, ROW write/anchor, busy |
| T12 | PEP 8, type hints, คอมเมนต์, ลบ test เปราะ, test ROW 'writing' | ครบ (flake8 ไม่มีข้อผิดพลาดตามที่ผู้สั่งงานแจ้ง ตรวจเองแล้วไม่มีบรรทัดเกิน 120 ตัว) |

## ภาพรวมรอบ T1–T12 — พร้อม commit
- 🔴 ทุกรอบ: 0 / 🟡 ที่เคยพบทั้ง 4 ข้อ (app_id ไม่ตรงกัน, สแกน UsedRange, End(xlUp) ข้ามแถวซ่อน, Table filter) ปิดครบแล้วพร้อม test และ smoke test กับ Excel จริง ส่วน 🟢 ที่เหลือเป็นเรื่องที่ควรแจ้งผู้ใช้หรือทำเพิ่มได้ภายหลัง
- ความปลอดภัย: ไม่พบ secret, อีเมล หรือชื่อจริงของบริษัทในโค้ด, test, `docs/` และ `.claude/` ที่ยังไม่ได้ track (test ใช้ `@example.test` ส่วนชื่อแอปอย่าง "COLLINS" มาจาก `tests/conftest.py` ที่ commit ไปแล้ว) ไม่มี input จากภายนอกเข้าไปใน COM call ใหม่ (มีแค่เลขแถว/คอลัมน์และค่าคงที่) และไม่มี `settings.local.json`
- ก่อน commit: (1) ควรตัดสินใจว่าจะ commit `.claude/` (รวม `.claude/ai-office/` ที่เป็นสำเนาแม่แบบ) และ `docs/` ไปด้วยหรือไม่ (2) git เตือนเรื่องแปลง LF -> CRLF ใน 9 ไฟล์ ไม่กระทบการทำงาน (3) ควรแจ้งผู้ใช้ถึงพฤติกรรมที่ผู้ใช้จะเห็นได้: แถวที่ถูกล้างจนเหลือแค่ App ID จะกลายเป็น inactive และ edit ที่ค้างอยู่จะ fail (T7/T11), Find dialog ของ Excel จะถูกรีเซ็ตเป็นค่า default (T9) และยังไม่รองรับการเขียนลงชีตที่ protect ไว้
