
import pyodbc
import xlsxwriter
from datetime import datetime,timedelta
print(datetime.today().strftime('%d'))
dt = (datetime.today()).strftime('%d%b%y')
pfile_name = 'C:/Mantra/Reports/AEEscalation/''AEEscalationReport' + '_' + dt + '.xlsx'
print(pfile_name)
