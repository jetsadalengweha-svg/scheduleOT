import csv
import ssl
import urllib.request
import sqlite3
import pandas as pd
import os

spreadsheet_id = os.environ.get('SPREADSHEET_ID')

context = ssl._create_unverified_context()

def processmatching(filename, filename2, outputfilename):
       with open(filename, "r") as f:
               text = f.read()
       with open(filename2, "r") as f2:
               text2 = f2.read()

       conn = sqlite3.connect(':memory:')
       cursor = conn.cursor()
       cursor.execute('CREATE TABLE sheet1 ( DateJob TEXT NOT NULL , Name TEXT NOT NULL, Job TEXT NOT NULL )')
       data = []
       holder = ""
       for row in text:
               if row == "," or row =="\n":
                       data.append(holder)
                       holder = ""
               else:
                       holder = holder + row

       for i in range(len(data)):
               if not i % 3:
                       sqlstatement = "INSERT INTO sheet1 ( DateJob, Name, Job )  VALUES ('" +  data[i] + "', '" + data[i+1] + "', '" + data[i+2] + "')"
                       cursor.execute(sqlstatement)
                       conn.commit()

       cursor.execute('CREATE TABLE sheet2 ( DateJob TEXT NOT NULL , Name TEXT NOT NULL, Job TEXT NOT NULL )')
       data2 = []
       holder2 = ""
       for row2 in text2:
               if row2 == "," or row2 =="\n":
                       data2.append(holder2)
                       holder2 = ""
               else:
                       holder2 = holder2 + row2

       for i2 in range(len(data2)):
               if not i2 % 3:
                       sqlstatement2 = "INSERT INTO sheet2 ( DateJob, Name, Job )  VALUES ('" +  data2[i2] + "', '" + data2[i2+1] + "', '" + data2[i2+2] + "')"
                       cursor.execute(sqlstatement2)
                       conn.commit()

       sqlstatementF = "SELECT s1.DateJob, s1.Name, s1.Job, s2.Job, s2.Name FROM sheet1 as s1 inner join sheet2 as s2 on s1.DateJob = s2.DateJob and s1.Job = s2.Job ORDER BY s1.DateJob ASC"
       cursor.execute( sqlstatementF )
       results = cursor.fetchall()
       
       f = open( outputfilename, 'w')
       for row in results:
               if row[1] == row[4]:
                       f.write ('%s,%s,%s,%s,%s,%s,%s,%s\n' % (row[0], row[1],'','', row[2],'', row[3], '') )
               else:
                       f.write ('%s,%s,%s,%s,%s,%s,%s,%s\n' % (row[0], row[1],'','', row[2],'', row[3], row[4]) )
       f.close()
       conn.close
       print("Done")

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
       

def runPharmOT():
       sheet_name = 'tblOrigin'
       url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
       datainput_cvs_format(url,'pharmO.csv')
       
def runPharmFOT():
       sheet_name = 'tblFinal'
       url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
       datainput_cvs_format(url,'pharmF.csv')
       
def runOfficerOT():
       sheet_name = 'tblOO'
       url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
       datainput_cvs_format(url,'officerO.csv')
       
def runOfficerFOT():
       sheet_name = 'tblOF'
       url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
       datainput_cvs_format(url,'officerF.csv')
       
def runHelperOT():
       sheet_name = 'tblHO'
       url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
       datainput_cvs_format(url,'helperO.csv')
       
def runHelperFOT():
       sheet_name = 'tblHF'
       url = 'https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}'.format(spreadsheet_id, sheet_name)
       datainput_cvs_format(url,'helperF.csv')
       
runPharmOT()
runPharmFOT()
runOfficerOT()
runOfficerFOT()
runHelperOT()
runHelperFOT()
print("Done")
processmatching("pharmO.csv", "pharmF.csv", "P.csv")
processmatching("officerO.csv", "officerF.csv", "O.csv")
processmatching("helperO.csv", "helperF.csv", "H.csv")
print("Done Processing")

