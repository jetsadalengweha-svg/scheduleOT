import gspread
from google.oauth2.service_account import Credentials

# Auth พร้อม Drive scope
scopes = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# เปิด Google Sheet ด้วย ID จาก URL
# https://docs.google.com/spreadsheets/d/xxxxxxx/edit
sheet = client.open_by_key("1ziUEgviiyYNZgOJ8UEr_iqnAKAo4FzCtoQJVwnzhsVo").get_worksheet(1)

# ดึงข้อมูลทั้งหมด
data = sheet.get_all_records()

# สร้างไฟล์ txt
with open("output.txt", "w", encoding="utf-8") as f:
    for row in data:
        f.write(str(row) + "\n")

print("บันทึกไฟล์ output.txt เรียบร้อยแล้ว")
