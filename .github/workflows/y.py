import pandas as pd

# เอา ID จาก URL ของ Google Sheet
# https://docs.google.com/spreadsheets/d/xxxxxxx/edit
sheet_id = "1ziUEgviiyYNZgOJ8UEr_iqnAKAo4FzCtoQJVwnzhsVo"
sheet_name = "STOCK"  # ชื่อชีทที่สอง

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

# ดึงข้อมูล
df = pd.read_csv(url)

# บันทึกเป็น txt
df.to_csv("output.txt", index=False, encoding="utf-8")

print("บันทึกไฟล์ output.txt เรียบร้อยแล้ว")
