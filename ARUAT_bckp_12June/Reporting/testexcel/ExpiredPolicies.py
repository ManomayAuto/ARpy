
import pyodbc
import xlsxwriter
from datetime import datetime, timedelta

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.193.194.40;'
                      'Database=MantraDB;'
                      'Trusted_Connection=no;')
print("Connected")

dt = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')
print(dt)
csr = conn.cursor()
csr.execute("Select Branch,Username,Clientname,EmailId,PhnNum,PolicyNum,PremDue,EquityDate from ae  where cast(equitydate as date)= ? and ReductionCheckBox = 0 order by Branch",dt)
rcrds = csr.fetchall()
print(rcrds)
workbook = xlsxwriter.Workbook('ExiredPolicies.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Branch')
worksheet.write('B1', 'User')
worksheet.write('C1', 'Client Name')
worksheet.write('D1', 'Email Id')
worksheet.write('E1', 'Phone Number')
worksheet.write('F1', 'Policy Number')
worksheet.write('G1', 'Premium Due')
worksheet.write('H1', 'Equity Date')
row = 1
col = 0

AErecords = len(rcrds)
print(AErecords)

for r in (rcrds):
    worksheet.write(row, col, r[0])
    worksheet.write(row, col + 1, r[1])
    worksheet.write(row, col + 2, r[2])
    worksheet.write(row, col + 3, r[3])
    worksheet.write(row, col + 4, r[4])
    worksheet.write(row, col + 5, r[5])
    worksheet.write(row, col + 6, r[6])
    worksheet.write(row, col + 7, r[7])
    row += 1


#NONAE

csr = conn.cursor()
csr.execute("Select Branch,Username,Clientname,EmailId,PhnNum,PolicyNum,PremDue,EquityDate from nonae where cast(Equitydate as date)= ? and ReductionCheckBox = 0 order by branch",dt)
rcrds = csr.fetchall()
print(rcrds)
if AErecords == 0:
    row = 1
else:
    row = row

for r in (rcrds):
    worksheet.write(row, col, r[0])
    worksheet.write(row, col + 1, r[1])
    worksheet.write(row, col + 2, r[2])
    worksheet.write(row, col + 3, r[3])
    worksheet.write(row, col + 4, r[4])
    worksheet.write(row, col + 5, r[5])
    worksheet.write(row, col + 6, r[6])
    worksheet.write(row, col + 7, r[7])
    row += 1

workbook.close()


