
import pyodbc
import xlsxwriter
from datetime import datetime

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.193.194.40;'
                      'Database=MantraDB;'
                      'Trusted_Connection=no;')
print("Connected")
csr = conn.cursor()
csr.execute("select nonae.branch,max(p1),max(p2),max(p3) from nonae join (Select Branch,count(policynum) as p1,0 as p2,0 as p3  from NONAE where DTE < 0 and reductioncheckbox = 0  group by branch union Select Branch,0 as p1,count(policynum) as p2,0  as p3  from NONAE where DTE = 0 and reductioncheckbox = 0 group by branch union Select Branch,0 as p1,0 as p2,count(policynum) as p3  from NONAE where DTE between 35 and 37 and followups = 0 group by branch) p  on p.branch = nonae.Branch group by nonae.branch")
rcrds = csr.fetchall()
print(rcrds)
dt = (datetime.today()).strftime('%d%b%y')
workbook = xlsxwriter.Workbook('C:/Mantra/Reports/''NONAEEscalationReport'+'_'+dt+'.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Branch')
worksheet.write('B1', 'Expired')
worksheet.write('C1', 'Expiring Today')
worksheet.write('D1', 'Collection Call Over Due')
row = 1
col = 0
for r in (rcrds):
    worksheet.write(row, col, r[0])
    worksheet.write(row, col + 1, r[1])
    worksheet.write(row, col + 2, r[2])
    worksheet.write(row, col + 3, r[3])
    row += 1

workbook.close()

