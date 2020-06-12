
import pyodbc
import xlsxwriter
from datetime import datetime,timedelta

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.193.194.40;'
                      'Database=MantraDB;'
                      'Trusted_Connection=no;')
print("Connected")
csr = conn.cursor()
print(datetime.today())
dt=datetime.today()
print(dt)
dt1=( dt).strftime('%Y-%m-%d')
print(dt1)

csr.execute(" Select um.name,max(f),max(r) from users_man um join (Select u.name,0 as f,0 as r from users_man u where u.Permissions like '%DF,DR%' union select ae.accexe,count(lastfollowup) as f,count(DateCompleted) as r from ae where cast(lastfollowup as date) = ? or cast(DateCompleted as date) = ? group by ae.accexe) a on um.name=a.name and cast(createddate as date)<= ? group by um.name ",dt1,dt1,dt1 )
rcrds = csr.fetchall()
print(rcrds)
AErecords = len(rcrds)
print(AErecords)
#dte = (datetime.today()).strftime('%d%b%y')
workbook = xlsxwriter.Workbook('C:/Mantra/Reports/ConsolidatedActivity/''ConsolidatedActivityReport.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'User')
worksheet.write('B1', 'User  Type')
worksheet.write('C1', 'Date')
worksheet.write('D1', 'Follow ups Completed')
worksheet.write('E1', 'Reductions Completed')
aeName = 'AE'
row = 1
col = 0
for r in (rcrds):
    worksheet.write(row, col, r[0])
    worksheet.write(row, col + 1, aeName)
    worksheet.write(row, col + 2, dt1)
    worksheet.write(row, col + 3, r[1])
    worksheet.write(row, col + 4, r[2])
    row += 1
    print(r[0])
    print(r[1])

#NONAE DATA - Followups

csr = conn.cursor()
csr.execute(" Select um.name,'Account',max(f),max(r) from users_man um join (Select u.name,0 as f,0 as r from users_man u where u.Permissions like 'DF' union select nonae.username,count(followups) as f,0 as r from NONAE where cast(DateCompleted as date) = ? and followups=1 group by nonae.Username) a on um.name=a.name and cast(createddate as date) <= ? group by um.name",dt1,dt1 )
rcrds = csr.fetchall()
print(rcrds)
NONAEfrecords = len(rcrds) + AErecords
# if AErecords == 0 :
#     row = 1
# else:
#     row = row
# col = 0
for r in (rcrds):
    worksheet.write(row, col, r[0])
    worksheet.write(row, col + 1, r[1])
    worksheet.write(row, col + 2, dt1)
    worksheet.write(row, col + 3, r[2])
    worksheet.write(row, col + 4, r[3])
    row += 1

#NONAE DATA - Reductions

csr = conn.cursor()
csr.execute("Select um.name,'CS',max(f),max(r) from users_man um join (Select u.name,0 as f,0 as r from users_man u where u.Permissions like 'DR' union select nonae.username,0 as f,count(ReductionCheckBox) as r from NONAE where cast(DateCompleted as date) =? and ReductionCheckBox=1 group by nonae.Username) a on um.name=a.name and cast(createddate as date) <= ? group by um.name",dt1,dt1 )
rcrds = csr.fetchall()
print(rcrds)
NONAERrecords = len(rcrds) + NONAEfrecords
# if NONAEfrecords == 0 :
#     row = 1
# else:
#     row = row
# col = 0
for r in (rcrds):
    worksheet.write(row, col, r[0])
    worksheet.write(row, col + 1, r[1])
    worksheet.write(row, col + 2, dt1)
    worksheet.write(row, col + 3, r[2])
    worksheet.write(row, col + 4, r[3])
    row += 1


workbook.close()

