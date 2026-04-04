import pandas as pd
import csv
import ssl
import urllib.request
import sqlite3
import pandas as pd
import os

spreadsheet_id = os.environ.get('SPREADSHEET_ID_PRODUCTEXPIRED')
context = ssl._create_unverified_context()
# เอา ID จาก URL ของ Google Sheet
sheet_name = "ExpiredDrug"  # ชื่อชีทที่สอง
url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
datainput_cvs_format(url,'product.csv')
       
