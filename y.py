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
       
def datainput_cvs_format(url,csvfile):
       response = urllib.request.urlopen(url, context=context)
       data = response.read().decode('utf-8')
       rows = list(csv.reader(data.splitlines()))
       header = rows[0]
       data_rows = rows[1:]
       f = open(csvfile, 'w' )
       for row in data_rows:
           name = row[1]
           for i in range(4, len(header)):
               label = header[i]
               value = row[i]
               if value.strip()  != "":          
                   f.write ( '%s,%s,%s\n' % (label, name, value.lower()) )
       f.close()
