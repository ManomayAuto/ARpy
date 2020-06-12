import pyodbc
#import xlsxwriter
from openpyxl import Workbook

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
print(type(rcrds))

dt = (datetime.today()).strftime('%d%b%y')

file_name = 'Dataae1.xlsx'
header = ('User', 'Expired', 'Expiring Today', 'Not Followed up')
wb = Workbook()
ws = wb.active
ws.append(header)


for r in rcrds:
    print("r is", type(r))
    r1 = tuple(r)
    print("r1 is", type(r1))
    ws.append(r1)

wb.save(file_name)
wb.close()




