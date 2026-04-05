import requests
import csv
import re
from datetime import datetime
from io import StringIO
import os

spreadsheet_id = os.environ.get('SPREADSHEET_ID_PRODUCTEXPIRED')
context = ssl._create_unverified_context()
# เอา ID จาก URL ของ Google Sheet
SHEET_NAME = "ExpiredDrug"  # ชื่อชีทที่สอง
SHEET_URL = "https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, SHEET_NAME)
# ====== ตั้งค่า ======
# วาง URL ของ Google Sheet ที่นี่ (ลิงค์ปกติที่ได้จากการแชร์)
OUTPUT_FILE = "result.csv"

# ====== แปลง URL เป็น CSV Export URL ======
def get_csv_export_url(sheet_url, sheet_name):
    """ดึง Spreadsheet ID จาก URL แล้วสร้าง CSV export link"""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("ไม่พบ Spreadsheet ID ใน URL — กรุณาตรวจสอบลิงค์อีกครั้ง")
    spreadsheet_id = match.group(1)
    return (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
        f"/gviz/tq?tqx=out:csv&sheet={requests.utils.quote(sheet_name)}"
    )

# ====== ดึงข้อมูลจาก Google Sheets ======
csv_url = get_csv_export_url(SHEET_URL, SHEET_NAME)
print(f"กำลังดึงข้อมูลจาก: {SHEET_NAME} ...")

response = requests.get(csv_url)
response.raise_for_status()
response.encoding = "utf-8"

# แปลง response เป็น list of rows
reader = csv.reader(StringIO(response.text))
all_rows = list(reader)

if not all_rows:
    print("ไม่พบข้อมูลในชีท")
    exit()

header    = all_rows[0]   # แถวหัวตาราง
data_rows = all_rows[1:]  # แถวข้อมูล

print(f"ดึงข้อมูลสำเร็จ: {len(data_rows)} แถว (ไม่นับหัวตาราง)")
print(f"คอลัมน์ที่พบ: {header}")

# ====== วันที่วันนี้ ======
today = datetime.today().date()
print(f"วันที่วันนี้: {today.strftime('%Y-%m-%d')}")

# รูปแบบวันที่ที่อาจพบในคอลัมน์ G
DATE_FORMATS = [
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%m/%d/%Y",
    "%d/%m/%y",
    "%Y/%m/%d",
]

def parse_date(date_str):
    """แปลงสตริงวันที่เป็น date object"""
    date_str = date_str.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None

# ====== กรองแถวที่คอลัมน์ G ตรงกับวันนี้ ======
# คอลัมน์ B = index 1 | คอลัมน์ G = index 6
matched_rows = []

for i, row in enumerate(data_rows, start=2):  # start=2 เพราะแถวที่ 1 คือ header
    if len(row) < 7:
        continue  # ข้ามแถวที่ข้อมูลไม่ถึงคอลัมน์ G

    col_b = row[1]  # คอลัมน์ B
    col_g = row[6]  # คอลัมน์ G (วันที่)

    if not col_g.strip():
        continue  # ข้ามถ้าคอลัมน์ G ว่าง

    row_date = parse_date(col_g)

    if row_date is None:
        print(f"  ⚠️  แถว {i}: แปลงวันที่ไม่ได้ → '{col_g}'")
        continue

    if row_date == today:
        matched_rows.append(row)
        print(f"  ✅ แถว {i}: คอลัมน์ B = '{col_b}' | วันที่ = '{col_g}'")

print(f"\nพบข้อมูลที่ตรงกับวันนี้: {len(matched_rows)} แถว")

# ====== บันทึกลง result.csv ======
with open(OUTPUT_FILE, "w", newline="", encoding="utf-8-sig") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(header)          # เขียนหัวตาราง
    writer.writerows(matched_rows)   # เขียนแถวที่ตรงกับวันนี้

print(f"บันทึกข้อมูลลงไฟล์ '{OUTPUT_FILE}' เรียบร้อยแล้ว ✅")
