import pyodbc
from openpyxl import Workbook
from datetime import datetime
import os
from openpyxl import *

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.193.194.40;'
                      'Database=MantraDB;'
                      'Trusted_Connection=no;')
print("Connected")


def close(self):
    """
    Close workbook file if open. Only affects read-only and write-only modes.
    """
    print("in close function")
    if hasattr(self, '_archive'):
        self._archive.close()
csr = conn.cursor()
csr.execute("select nonae.branch,max(p1),max(p2),max(p3) from nonae join (Select Branch,count(policynum) as p1,0 as p2,0 as p3  from NONAE where DTE < 0 and reductioncheckbox = 0  group by branch union Select Branch,0 as p1,count(policynum) as p2,0  as p3  from NONAE where DTE = 0 and reductioncheckbox = 0 group by branch union Select Branch,0 as p1,0 as p2,count(policynum) as p3  from NONAE where DTE between 35 and 37 and followups = 0 group by branch) p  on p.branch = nonae.Branch group by nonae.branch")
rcrds = csr.fetchall()
print(rcrds)
dt = (datetime.today()).strftime('%d%b%y')
file_name = 'Datanonae1.xlsx'
t = os.path.exists(file_name)
if t:
    wb = load_workbook('Datanonae1.xlsx')
    print("In true statement")
    wb.close(self)

header = ('Branch', 'Expired', 'Expiring Today', 'Collection Call Over Due')
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



