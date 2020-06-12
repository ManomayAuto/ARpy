
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
csr.execute("select username,'AE',Branch,ClientName,Policynum,premdue,ReductionTransid from AE where cast(datecompleted as date)= ? and ReductionCheckBox = 1 order by username",dt)
rcrds = csr.fetchall()
print(rcrds)
workbook = xlsxwriter.Workbook('PendingWriteoff(AE&NONAE)Report.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Processed by User')
worksheet.write('B1', 'User Type')
worksheet.write('C1', 'Branch')
worksheet.write('D1', 'Customer Name')
worksheet.write('E1', 'Policy Number')
worksheet.write('F1', 'Premium Due')
worksheet.write('G1', 'ReductionTransid')
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
    row += 1


#NONAE

csr = conn.cursor()
csr.execute("select Username,'CS',Branch,ClientName,Policynum,premdue,ReductionTransid from NONAE where cast(datecompleted as date)= ? and ReductionCheckBox = 1 order by Username",dt)
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
    row += 1

workbook.close()


