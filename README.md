# EDC Search — วิธีติดตั้งและใช้งาน

## โครงสร้างไฟล์
```
edc-search/
├── convert.py     ← script แปลง Excel → JSON (รันทุกวันจันทร์)
├── index.html     ← หน้าเว็บค้นหา (deploy ขึ้น GitHub Pages)
├── data.json      ← ไฟล์ข้อมูล (สร้างจาก convert.py อัตโนมัติ)
└── README.md
```

---

## วิธีใช้งาน (ทุกวันจันทร์)

### ขั้นตอนที่ 1 — เตรียมไฟล์ Excel
วางไฟล์ Excel ทุกแบงค์ในโฟลเดอร์เดียวกับ convert.py
- ชื่อไฟล์ต้องมีชื่อแบงค์ เช่น: `KBANK_Jun2026.xlsx`, `SCB_Report.xlsx`
- รองรับ: KBANK, SCB, BBL, BAY

### ขั้นตอนที่ 2 — รัน convert script
```bash
pip install openpyxl
python convert.py
```
จะได้ไฟล์ `data.json` ใหม่

### ขั้นตอนที่ 3 — Push ขึ้น GitHub
```bash
git add data.json
git commit -m "Update data $(date +%Y-%m-%d)"
git push
```
GitHub Pages จะอัพเดทอัตโนมัติภายใน 1-2 นาที

---

## วิธี Deploy ครั้งแรก (ทำครั้งเดียว)

1. สร้าง GitHub repository ใหม่ (ชื่ออะไรก็ได้)
2. อัพโหลดไฟล์ทั้งหมด: `index.html`, `convert.py`, `data.json`, `README.md`
3. ไปที่ **Settings → Pages → Source: main branch → / (root)**
4. กด Save → รอ 2-3 นาที
5. URL จะเป็น: `https://[username].github.io/[repo-name]/`
6. แชร์ URL ให้ทีม — เข้าได้เลย ไม่ต้อง login

---

## ปรับ Config ให้ตรงกับ Column ในไฟล์จริง

เปิด `convert.py` แล้วแก้ส่วน `BANK_CONFIGS`:

```python
"KBANK": {
    "tid_col": "TID",          # ชื่อ column TID ในไฟล์ Excel
    "mid_col": "MID",          # ชื่อ column MID
    "sn_col": "S/N EDC",       # ชื่อ column Serial Number
    "slip_col": "Slip Line_1", # ชื่อ column ชื่อร้าน
    ...
}
```

---

## Q&A

**Q: ไฟล์ data.json ขนาดเท่าไหร่?**
A: ประมาณ 5-15 MB ขึ้นอยู่กับจำนวน record — GitHub Pages รองรับได้สบาย

**Q: คนเข้าพร้อมกันได้กี่คน?**
A: ไม่จำกัด เพราะเป็น static file ไม่มี server

**Q: ถ้า column ไม่ตรงทำยังไง?**
A: รัน convert.py แล้วดู error message — จะบอกว่า column ไหนหาไม่เจอ แล้วแก้ใน BANK_CONFIGS

**Q: เพิ่มแบงค์ใหม่ได้ไหม?**
A: ได้ เพิ่ม entry ใหม่ใน BANK_CONFIGS พร้อม pattern ชื่อไฟล์และ column names
