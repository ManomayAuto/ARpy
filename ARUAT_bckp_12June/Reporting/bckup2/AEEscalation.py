
import pyodbc
import xlsxwriter

from datetime import datetime, timedelta

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.193.194.40;'
                      'Database=MantraDB;'
                      'Trusted_Connection=no;')
print("Connected")

csr = conn.cursor()
csr.execute("select users.name,max(p1),max(p2),max(p3) from users_man users left join (  Select name ,0 as p1,0 as p2,0 as p3 from users_man  where Permissions like '%DF,DR%'  union Select Accexe,count(policynum) as p1,0 as p2,0 as p3  from AE where DTE < 0 and  reductioncheckbox = 0  group by Accexe union Select Accexe,0 as p1,count(policynum) as p2,0  as p3  from AE where DTE = 0 and reductioncheckbox = 0  group by Accexe union  Select ACcexe,0 as p1,0 as p2,count(policynum) as p3  from AE where DTE = 15 and Followups = 0  group by Accexe)p on p.name = users.name where  users.Permissions like '%DF,DR%' group by users.name")
rcrds = csr.fetchall()
print(rcrds)
dt = (datetime.today()).strftime('%d%b%y')
workbook = xlsxwriter.Workbook('C:/Mantra/Reports/AEEscalation/''AEEscalationReport'+'_'+dt+'.xlsx')
worksheet = workbook.add_worksheet()
header = workbook.add_format({'border': 2,'bold': True, 'align': "center", 'bg_color': "#d9e1f2" })
fields = workbook.add_format({'border': 1, 'align': "left" })
worksheet.freeze_panes(1,0)
worksheet.write('A1', 'User', header)
worksheet.write('B1', 'Expired', header)
worksheet.write('C1', 'Expiring Today', header)
worksheet.write('D1', 'Not Followed up', header)
row = 1
col = 0
for r in (rcrds):
    worksheet.write(row, col, r[0], fields)
    worksheet.write(row, col + 1, r[1], fields)
    worksheet.write(row, col + 2, r[2], fields)
    worksheet.write(row, col + 3, r[3], fields)
    row += 1

workbook.close()

