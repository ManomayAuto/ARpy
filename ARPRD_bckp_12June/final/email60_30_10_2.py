from configparser import ConfigParser
from email.mime.multipart import MIMEMultipart
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
##from Crypto.Cipher import AES
##from Crypto.Util.Padding import pad, unpad
##from base64 import b64encode, b64decode
import pyodbc
import os
import xlsxwriter
from datetime import date
import time
import datetime
from datetime import datetime,timedelta




# get emailid and password from config file
def get_email_data():
    config = ConfigParser()
    config.read('email_config.ini')
    email_details = config["emaildata"]
    if email_details["encrypted"] == "No":
        email = email_details["emailid"]
        password = email_details["password"]
        return email, password
    elif email_details["encrypted"] == "Yes":
        return email_details["emailid"], Encryption().decrypt(email_details["password"])


cnxn = pyodbc.connect("Driver={SQL Server};"
                      "Server=Inbroker4;"
                      "Database=MantraDB;"
                      "Trusted_Connection=yes;")


def sixty():

    cur = cnxn.cursor()
    cur.execute("""

Select  c.IdEnte as ClientID,c.BuscarComo as ClientName,
case when c.Telefono1 is not null and c.Telefono1 != 'NONE' and c.Telefono1 != ' '
	then case when c.CodTipoTelefono1 = 'PAR' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'CEL' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'LEG' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'COM' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'FAX' then c.CodTipoTelefono1+' - '+c.Telefono1

			 end 
			 end as Phnnum1,
		CASE WHEN   sgsp.DescSubProducto like '%-%' then
substring(sgsp.descsubproducto,1,PATINDEX('%-%',sgsp.DescSubProducto)-1)
else sgsp.DescSubProducto
 end as ProdType,
sgo.codpoliza as PolicyNum,
sum(sgo.ImpSaldoOperacion) as PremDue,
cast(DATEADD(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza) as date)
as EquityDate, 
DATEDIFF(day,GETDATE(),Dateadd(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza)) 
as DTE,
0 as Followups,
case when c.EMail = 'NULL' then NULL
when c.EMail = '' then NULL
else c.email end as EmailID,case when sgc.CodSegmentacion5 is null then sgs1.descsegmentacion
when sgc.codsegmentacion5 is not null then sgs2.descsegmentacion
end as Branch,case	when c.Telefono2 is not null and c.Telefono2 != 'NONE' and c.Telefono2 != ' '
	then case when c.CodTipoTelefono2 = 'PAR' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'LEG' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'CEL' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'FAX' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'COM' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 is NULL then 'No Type -'+c.Telefono2
			end 
			end as Phnnum2,

	case when c.Telefono3 is not null and c.Telefono3 != 'NONE' and c.Telefono3 != ' '
	then case when c.CodTipoTelefono3 = 'CEL' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'PAR' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'LEG' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'FAX' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'COM' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 is NULL then 'No Type -'+c.Telefono3
			end
		end as PhnNum3

from Inbroker3.InbkProd.dbo.Sg_Operacion sgo
inner join (select codpoliza c,idoperacion m from Inbroker3.InbkProd.dbo.sg_operacion where FecVigenciaInicialOperacion < getdate() and
  idpoliza in(select max(idpoliza) from Inbroker3.InbkProd.dbo.sg_operacion where FecVigenciaFinalPoliza >= GEtdate() and 
  FecVigenciaInicialPoliza <=GETDATE()  group by codpoliza) group by codpoliza,IdOperacion)sgo1 on  
sgo.idoperacion=sgo1.m and sgo.CodPoliza=sgo1.c
left join Inbroker3.InbkProd.dbo.Sg_SubProducto sgsp on sgsp.CodSubProducto=sgo.CodSubProducto
left join Inbroker3.InbkProd.dbo.tg_ente c
on c.IdEnte=sgo.IdCliente 
join Inbroker3.InbkProd.dbo.sg_cliente sgc
on sgc.idcliente=c.idente
left join Inbroker3.InbkProd.dbo.Sg_SegmentacionCartera sgs1
on sgs1.CodSegmentacion=sgc.CodSegmentacion5
left join Inbroker3.InbkProd.dbo.Sg_SegmentacionCartera sgs2
on sgs2.CodSegmentacion=sgo.CodSegmentacion5
where   sgo.FecVigenciaFinalPoliza >= GEtdate() and sgo.FecVigenciaInicialPoliza <=GETDATE() and (sgo.CodSegmentacion3 is NULL and sgc.CodSegmentacion3 is null )and


sgo.CodPoliza not in (select CodPoliza from inbroker3.inbkprod.dbo.sg_operacion where idoperacion in (
 select max(idoperacion) from inbroker3.inbkprod.dbo.sg_operacion sgo
 where ((sgo.codtipoendoso  in ('P')) or (sgo.CodTipoEndoso in ('A') and sgo.FecVigenciaInicialOperacion <= getdate()))  and IdPoliza in (select max(idpoliza) from inbroker3.inbkprod.dbo.sg_operacion where FecVigenciaFinalPoliza >= GEtdate() and 
  FecVigenciaInicialPoliza <=GETDATE()  group by codpoliza)group by codpoliza ))

group by C.IdEnte,C.ApellidoRazonSocial,C.NomEnte,C.Telefono1,sgo.CodPoliza,sgo.FecVigenciaInicialPoliza,c.Telefono2,c.Telefono3,c.CodTipoTelefono1,c.CodTipoTelefono2,c.CodTipoTelefono3,
sgo.FecVigenciaFinalPoliza,c.EMail,sgsp.DescSubProducto,BuscarComo,sgs1.DescSegmentacion,sgs2.DescSegmentacion,sgc.CodSegmentacion5
having  sum(sgo.ImpSaldoOperacion) >0.1 and SUM(sgo.DP_imppremioIVA) >0.1  and 
(case 
	when (sum(sgo.DP_imppremioIVA) >0.1) and sum(ImpSaldoOperacion) >0.1
	then DATEDIFF(day,GETDATE(),Dateadd(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza)) 
	end) =60

or

(case 
	when (sum(sgo.DP_imppremioIVA) >0.1) and sum(ImpSaldoOperacion) >0.1
	then DATEDIFF(day,GETDATE(),Dateadd(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza)) 
	end) =60



""")

    records = cur.fetchall()

    dt = (datetime.today()).strftime('%d%b%y')
    workbook = xlsxwriter.Workbook("C:/Users/administrator.JSJHO/Desktop/Consultant Files/Mantra/Emailfollowup/EmailFollowUps"+"_"+dt+".xlsx")
    worksheet = workbook.add_worksheet()
    header = workbook.add_format({'border': 1, 'bold': True, 'align': "center", 'bg_color': "#d9e1f2"})
    fields = workbook.add_format({'border': 1, 'align': "left"})
    worksheet.freeze_panes(0, 0)
    worksheet.write(0, 0, 'Branch', header)
    worksheet.write(0, 1, 'Client Id', header)
    worksheet.write(0, 2, 'Client Name', header)
    worksheet.write(0, 3, 'Policy Number', header)
    worksheet.write(0, 4, 'Reminder Type', header)
    worksheet.write(0, 5, 'Premium Due', header)
    worksheet.write(0, 6, 'Phone Number 1', header)
    worksheet.write(0, 7, 'Phone Number 2', header)
    worksheet.write(0, 8, 'Phone Number 3', header)
    worksheet.write(0, 9, 'Sent Status', header)
    row = 1
    for record in records:
        clId = record[0]
        brnch = record[10]
        clNam = record[1]
        prodType = record[3]
        polNum = record[4]
        premDue = record[5]
        premDue = round(premDue, 2)
        print(premDue)
        emailid = record[9]
        dte = record[7]
        phnNum = record[2]
        phnNum1 = record[11]
        phnNum2 = record[12]
        if emailid is None:
            emailid = ''
            print(emailid)
        elif "," in emailid:
            emailid = emailid.split(",")
            print(emailid)
        elif ";" in emailid:
            emailid = emailid.split(";")
            print(emailid)
        # email for 60 days
        time1 = datetime.now()
        email, password = get_email_data()
        sender = email
        receivers = emailid

        msg = MIMEMultipart()
        msg['subject'] = "Overdue - Instalment Payment"
        message = """<p>Dear """ + str(clNam) + """</p>

       <p>The payment on your """ + str(prodType) + """ Policy-""" + str(
            polNum) + """ is overdue. The total outstanding balance is """ + str(premDue) + """.</p>

       <p>Kindly make the payment as soon as possible to ensure the full coverage continuity.</p>

       <p>You can make the payment by visiting any of our offices, online banking through Scotiabank, Royal </p>

       <p>Bank of Canada and Commonwealth Bank or via our online service centre by clicking on the link <a href="https://jsjohnsononline.com/osc/osc/">https://jsjohnsononline.com/osc/osc/</a></p>

       <p>Should you have any questions, you can contact us at (242) 397-2100 or visit any nearest office.</p>


        <p>Regards</p>
        <p>Team</p>
        <p>J.S JOHNSON</p>
        <p>Web: www.jsjohnson.com</p>
        <p>Tel: (242) 397-2100|Fax: (242) 323-3720</p>

        <p style="color:red;">**This is an auto-generated email, please do not respond to this</p>
        """
        msg.attach(MIMEText(message, 'html'))
        try:
            smtpObj = smtplib.SMTP('mail.jsjemail.com')
            smtpObj.login(email, password)
            smtpObj.sendmail(sender, receivers, msg.as_string())
            cur = cnxn.cursor()
            cur.execute(
                """insert into emailsBackup(ClientName,PolicyNum,DTE,Status,Date_time) values (?,?,?,'Success',?)""",
                (str(clNam), str(polNum), str(dte), time1))
            cur.commit()
            worksheet.write(row, 0, brnch, fields)
            worksheet.write(row, 1, clId, fields)
            worksheet.write(row, 2, clNam, fields)
            worksheet.write(row, 3, polNum, fields)
            worksheet.write(row, 4, dte, fields)
            worksheet.write(row, 5, premDue, fields)
            worksheet.write(row, 6, phnNum, fields)
            worksheet.write(row, 7, phnNum1, fields)
            worksheet.write(row, 8, phnNum2, fields)
            worksheet.write(row, 9, "Sent", fields)
            row += 1
            print("Successfully sent email")

        except smtplib.SMTPException as e:
            print("Error: unable to send email")
            print(e)
            cur = cnxn.cursor()
            cur.execute(
                """insert into emailsBackup(ClientName,PolicyNum,DTE,Status,Date_time) values (?,?,?,'Fail',?)""",
                (str(clNam), str(polNum), str(dte), time1))
            cur.commit()
            worksheet.write(row, 0, brnch, fields)
            worksheet.write(row, 1, clId, fields)
            worksheet.write(row, 2, clNam, fields)
            worksheet.write(row, 3, polNum, fields)
            worksheet.write(row, 4, dte, fields)
            worksheet.write(row, 5, premDue, fields)
            worksheet.write(row, 6, phnNum, fields)
            worksheet.write(row, 7, phnNum1, fields)
            worksheet.write(row, 8, phnNum2, fields)
            worksheet.write(row, 9, "Fail",fields)
            row += 1

    return row, worksheet, workbook,header, fields


def thirty():
    row, worksheet, workbook,header, fields = sixty()
    cur = cnxn.cursor()
    cur.execute("""
Select  c.IdEnte as ClientID,c.BuscarComo as ClientName,
case when c.Telefono1 is not null and c.Telefono1 != 'NONE' and c.Telefono1 != ' '
	then case when c.CodTipoTelefono1 = 'PAR' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'CEL' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'LEG' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'COM' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'FAX' then c.CodTipoTelefono1+' - '+c.Telefono1

			 end 
			 end as Phnnum1,
		CASE WHEN   sgsp.DescSubProducto like '%-%' then
substring(sgsp.descsubproducto,1,PATINDEX('%-%',sgsp.DescSubProducto)-1)
else sgsp.DescSubProducto
 end as ProdType,
sgo.codpoliza as PolicyNum,
sum(sgo.ImpSaldoOperacion) as PremDue,
cast(DATEADD(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza) as date)
as EquityDate, 
DATEDIFF(day,GETDATE(),Dateadd(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza)) 
as DTE,
0 as Followups,
case when c.EMail = 'NULL' then NULL
when c.EMail = '' then NULL
else c.email end as EmailID,case when sgc.CodSegmentacion5 is null then sgs1.descsegmentacion
when sgc.codsegmentacion5 is not null then sgs2.descsegmentacion
end as Branch,case	when c.Telefono2 is not null and c.Telefono2 != 'NONE' and c.Telefono2 != ' '
	then case when c.CodTipoTelefono2 = 'PAR' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'LEG' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'CEL' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'FAX' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'COM' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 is NULL then 'No Type -'+c.Telefono2
			end 
			end as Phnnum2,

	case when c.Telefono3 is not null and c.Telefono3 != 'NONE' and c.Telefono3 != ' '
	then case when c.CodTipoTelefono3 = 'CEL' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'PAR' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'LEG' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'FAX' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'COM' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 is NULL then 'No Type -'+c.Telefono3
			end
		end as PhnNum3

from Inbroker3.InbkProd.dbo.Sg_Operacion sgo
inner join (select codpoliza c,idoperacion m from Inbroker3.InbkProd.dbo.sg_operacion where FecVigenciaInicialOperacion < getdate() and
  idpoliza in(select max(idpoliza) from Inbroker3.InbkProd.dbo.sg_operacion where FecVigenciaFinalPoliza >= GEtdate() and 
  FecVigenciaInicialPoliza <=GETDATE()  group by codpoliza) group by codpoliza,IdOperacion)sgo1 on  
sgo.idoperacion=sgo1.m and sgo.CodPoliza=sgo1.c
left join Inbroker3.InbkProd.dbo.Sg_SubProducto sgsp on sgsp.CodSubProducto=sgo.CodSubProducto
left join Inbroker3.InbkProd.dbo.tg_ente c
on c.IdEnte=sgo.IdCliente 
join Inbroker3.InbkProd.dbo.sg_cliente sgc
on sgc.idcliente=c.idente
left join Inbroker3.InbkProd.dbo.Sg_SegmentacionCartera sgs1
on sgs1.CodSegmentacion=sgc.CodSegmentacion5
left join Inbroker3.InbkProd.dbo.Sg_SegmentacionCartera sgs2
on sgs2.CodSegmentacion=sgo.CodSegmentacion5
where   sgo.FecVigenciaFinalPoliza >= GEtdate() and sgo.FecVigenciaInicialPoliza <=GETDATE() and (sgo.CodSegmentacion3 is NULL and sgc.CodSegmentacion3 is null )and

sgo.CodPoliza not in (select CodPoliza from inbroker3.inbkprod.dbo.sg_operacion where idoperacion in (
 select max(idoperacion) from inbroker3.inbkprod.dbo.sg_operacion sgo
 where ((sgo.codtipoendoso  in ('P')) or (sgo.CodTipoEndoso in ('A') and sgo.FecVigenciaInicialOperacion <= getdate()))  and IdPoliza in (select max(idpoliza) from inbroker3.inbkprod.dbo.sg_operacion where FecVigenciaFinalPoliza >= GEtdate() and 
  FecVigenciaInicialPoliza <=GETDATE()  group by codpoliza)group by codpoliza ))

group by C.IdEnte,C.ApellidoRazonSocial,C.NomEnte,C.Telefono1,sgo.CodPoliza,sgo.FecVigenciaInicialPoliza,c.Telefono2,c.Telefono3,c.CodTipoTelefono1,c.CodTipoTelefono2,c.CodTipoTelefono3,
sgo.FecVigenciaFinalPoliza,c.EMail,sgsp.DescSubProducto,BuscarComo,sgs1.DescSegmentacion,sgs2.DescSegmentacion,sgc.CodSegmentacion5
having  sum(sgo.ImpSaldoOperacion) >0.1 and SUM(sgo.DP_imppremioIVA) >0.1  and 
(case 
	when (sum(sgo.DP_imppremioIVA) >0.1) and sum(ImpSaldoOperacion) >0.1
	then DATEDIFF(day,GETDATE(),Dateadd(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza)) 
	end) =30

or

(case 
	when (sum(sgo.DP_imppremioIVA) >0.1) and sum(ImpSaldoOperacion) >0.1
	then DATEDIFF(day,GETDATE(),Dateadd(day,ceiling(round((sum(sgo.DP_imppremioIVA)-sum(sgo.ImpSaldoOperacion))/(sum(sgo.DP_ImpPremioIVA)/datediff(day,sgo.FecVigenciaInicialPoliza,sgo.FecVigenciaFinalPoliza)),1)),sgo.FecVigenciaInicialPoliza)) 
	end) =30
""")
    records = cur.fetchall()

    for record in records:
        clId = record[0]
        brnch = record[10]
        clNam = record[1]
        prodType = record[3]
        polNum = record[4]
        premDue = record[5]
        premDue = round(premDue, 2)
        emailid = record[9]
        eqdt = record[6]
        dte = record[7]
        phnNum = record[2]
        phnNum1 = record[11]
        phnNum2 = record[12]

        if emailid is None:
            emailid = ''
            print(emailid)
        elif "," in emailid:
            emailid = emailid.split(",")
            print(emailid)
        elif ";" in emailid:
            emailid = emailid.split(";")
            print(emailid)
        # email for 30 days
        time1 = datetime.now()
        email, password = get_email_data()
        sender = email
        receivers = emailid

        msg = MIMEMultipart()
        msg['subject'] = "Warning – Policy Cover Reduction"
        message = """<p>Dear """ + str(clNam) + """</p>

        <p>The payment on your """ + str(prodType) + """ Policy-""" + str(
            polNum) + """ has been overdue. The total outstanding balance is """ + str(premDue) + """.</p>

        <p style="color:red;"><strong>Should we not receive an immediate payment, your policy period would be reduced to """ + str(
            eqdt) + """</strong></p>

        <p>You can make the payment by visiting any of our offices, online banking through Scotiabank, Royal </p>

        <p>Bank of Canada and Commonwealth Bank or via our online service centre by clicking on the link <a href="https://jsjohnsononline.com/osc/osc/">https://jsjohnsononline.com/osc/osc/</a></p>

        <p>Should you have any questions, you can contact us at (242) 397-2100 or visit any nearest office.</p>


        <p>Regards</p>
        <p>Team</p>
        <p>J.S JOHNSON</p>
        <p>Web: www.jsjohnson.com</p>
        <p>Tel: (242) 397-2100|Fax: (242) 323-3720</p>

        <p style="color:red;">**This is an auto-generated email, please do not respond to this</p>
        """
        msg.attach(MIMEText(message, 'html'))
        try:
            smtpObj = smtplib.SMTP('mail.jsjemail.com')
            smtpObj.login(email, password)
            smtpObj.sendmail(sender, receivers, msg.as_string())
            cur = cnxn.cursor()
            cur.execute(
                """insert into emailsBackup(ClientName,PolicyNum,DTE,Status,Date_time) values (?,?,?,'Success',?)""",
                (str(clNam), str(polNum), str(dte), time1))
            cur.commit()
            worksheet.write(row, 0, brnch, fields)
            worksheet.write(row, 1, clId, fields)
            worksheet.write(row, 2, clNam, fields)
            worksheet.write(row, 3, polNum, fields)
            worksheet.write(row, 4, dte, fields)
            worksheet.write(row, 5, premDue, fields)
            worksheet.write(row, 6, phnNum, fields)
            worksheet.write(row, 7, phnNum1, fields)
            worksheet.write(row, 8, phnNum2, fields)
            worksheet.write(row, 9, "Sent", fields)
            row += 1
            print("Successfully sent email")

        except smtplib.SMTPException as e:
            print("Error: unable to send email")
            print(e)
            cur = cnxn.cursor()
            cur.execute(
                """insert into emailsBackup(ClientName,PolicyNum,DTE,Status,Date_time) values (?,?,?,'Fail',?)""",
                (str(clNam), str(polNum), str(dte), time1))
            cur.commit()
            worksheet.write(row, 0, brnch, fields)
            worksheet.write(row, 1, clId, fields)
            worksheet.write(row, 2, clNam, fields)
            worksheet.write(row, 3, polNum, fields)
            worksheet.write(row, 4, dte, fields)
            worksheet.write(row, 5, premDue, fields)
            worksheet.write(row, 6, phnNum, fields)
            worksheet.write(row, 7, phnNum1, fields)
            worksheet.write(row, 8, phnNum2, fields)
            worksheet.write(row, 9, "Fail", fields)
            row += 1
    workbook.close()


thirty()


def tenTwo():
    cnxn = pyodbc.connect("Driver={SQL Server};"
                          "Server=Inbroker3;"
                          "Database=InbkProd;"
                          "Trusted_Connection=yes;")
    cur = cnxn.cursor()
    cur.execute("""


Select  c.IdEnte as ClientID,c.BuscarComo as ClientName,
case when c.Telefono1 is not null and c.Telefono1 != 'NONE' and c.Telefono1 != ' '
	then case when c.CodTipoTelefono1 = 'PAR' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'CEL' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'LEG' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'COM' then c.CodTipoTelefono1+' - '+c.Telefono1
			  when c.CodTipoTelefono1 = 'FAX' then c.CodTipoTelefono1+' - '+c.Telefono1

			 end 
			 end as Phnnum1,


		CASE WHEN   sgsp.DescSubProducto like '%-%' then
substring(sgsp.descsubproducto,1,PATINDEX('%-%',sgsp.DescSubProducto)-1)
else sgsp.DescSubProducto
 end as ProdType,
sgo.codpoliza as PolicyNum,
sgi.ImpCuota as InstalmentDue,sgi.FecVencimiento as InstalmentDueDate,3 as DTE,0 as Followups,
case when c.EMail = 'NULL' then NULL
when c.EMail = '' then NULL
else c.email end as EmailID,sgi.NroCuota as InstalmentNumber,
case when sgc.CodSegmentacion5 is null then sgs1.descsegmentacion
when sgc.codsegmentacion5 is not null then sgs2.descsegmentacion
end as Branch,

case	when c.Telefono2 is not null and c.Telefono2 != 'NONE' and c.Telefono2 != ' '
	then case when c.CodTipoTelefono2 = 'PAR' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'LEG' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'CEL' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'FAX' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 = 'COM' then c.CodTipoTelefono2+' - '+c.Telefono2
			  when c.CodTipoTelefono2 is NULL then 'No Type -'+c.Telefono2
			end 
			end as Phnnum2,

	case when c.Telefono3 is not null and c.Telefono3 != 'NONE' and c.Telefono3 != ' '
	then case when c.CodTipoTelefono3 = 'CEL' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'PAR' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'LEG' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'FAX' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 = 'COM' then c.CodTipoTelefono3+' - '+c.Telefono3
			  when c.CodTipoTelefono3 is NULL then 'No Type -'+c.Telefono3
			end
		end as PhnNum3


from Inbroker3.InbkProd.dbo.Sg_Operacion sgo
inner join (select codpoliza c,idoperacion m from Inbroker3.InbkProd.dbo.sg_operacion where FecVigenciaInicialOperacion < getdate() and
  idpoliza in(select max(idpoliza) from Inbroker3.InbkProd.dbo.sg_operacion where FecVigenciaFinalPoliza >= GEtdate() and 
  FecVigenciaInicialPoliza <=GETDATE()  group by codpoliza) group by codpoliza,IdOperacion)sgo1 on  
sgo.idoperacion=sgo1.m and sgo.CodPoliza=sgo1.c
left join Inbroker3.InbkProd.dbo.Sg_SubProducto sgsp on sgsp.CodSubProducto=sgo.CodSubProducto
left join Inbroker3.InbkProd.dbo.tg_ente c
on c.IdEnte=sgo.IdCliente 
join Inbroker3.InbkProd.dbo.sg_cliente sgc
on sgc.idcliente=c.idente
left join Inbroker3.InbkProd.dbo.Sg_SegmentacionCartera sgs1
on sgs1.CodSegmentacion=sgc.CodSegmentacion5
left join Inbroker3.InbkProd.dbo.Sg_SegmentacionCartera sgs2
on sgs2.CodSegmentacion=sgo.CodSegmentacion5
left join Sg_ParametroPlanPago sgp
on sgp.IdParametroPlanPago=sgo.IdParametroPlanPago 
left Join Sg_CuentaCorrientePremio sgi
on sgi.idpoliza=sgo.idpoliza
where   sgo.FecVigenciaFinalPoliza >= GEtdate() and sgo.FecVigenciaInicialPoliza <=GETDATE() and (sgo.CodSegmentacion3 is NULL and sgc.CodSegmentacion3 is null ) and sgo.IdParametroPlanPago!=24 and
 sgi.IdCobranza is NULL and

sgo.CodPoliza not in (select CodPoliza from inbroker3.inbkprod.dbo.sg_operacion where idoperacion in (
 select max(idoperacion) from inbroker3.inbkprod.dbo.sg_operacion sgo
 where ((sgo.codtipoendoso  in ('P')) or (sgo.CodTipoEndoso in ('A') and sgo.FecVigenciaInicialOperacion <= getdate()))  and IdPoliza in (select max(idpoliza) from inbroker3.inbkprod.dbo.sg_operacion where FecVigenciaFinalPoliza >= GEtdate() and 
  FecVigenciaInicialPoliza <=GETDATE()  group by codpoliza)group by codpoliza ))

group by C.IdEnte,C.Telefono1,sgo.CodPoliza,sgo.FecVigenciaInicialPoliza,c.Telefono2,c.Telefono3,c.CodTipoTelefono1,c.CodTipoTelefono2,c.CodTipoTelefono3,
sgo.FecVigenciaFinalPoliza,c.EMail,sgsp.DescSubProducto,BuscarComo,sgi.ImpCuota,sgi.FecVencimiento,sgi.NroCuota,sgc.CodSegmentacion5,sgs1.DescSegmentacion,sgs2.DescSegmentacion
having  sum(sgo.ImpSaldoOperacion) >0.1 and SUM(sgo.DP_imppremioIVA) >0.1  and cast(DATEADD(D,3,GETDATE()) as date)= cast(sgi.FecVencimiento as date)

""")

    records = cur.fetchall()
    dt = (datetime.today()).strftime('%d%b%y')
    workbook = xlsxwriter.Workbook('C:/Users/administrator.JSJHO/Desktop/Consultant Files/Mantra/Instalmentreminder/InstalmentReminderFailedReport'+'_'+dt+'.xlsx')
    worksheet = workbook.add_worksheet()
    header = workbook.add_format({'border': 1, 'bold': True, 'align': "center", 'bg_color': "#d9e1f2"})
    fields = workbook.add_format({'border': 1, 'align': "left"})
    worksheet.freeze_panes(0, 0)
    worksheet.write(0, 0, 'Branch', header)
    worksheet.write(0, 1, 'Client Id', header)
    worksheet.write(0, 2, 'Client Name', header)
    worksheet.write(0, 3, 'Policy Number', header)
    worksheet.write(0, 4, ' Instalment Number', header)
    worksheet.write(0, 5, ' Instalment Amount', header)
    worksheet.write(0, 6, 'Phone Number 1', header)
    worksheet.write(0, 7, 'Phone Number 2', header)
    worksheet.write(0, 8, 'Phone Number 3', header)

    row = 1
    for record in records:
        clId = record[0]
        brnch = record[11]
        clNam = record[1]
        prodType = record[3]
        polNum = record[4]
        premDue = record[5]
        premDue = round(premDue, 2)
        emailid = record[9]
        dte = record[7]
        phnNum = record[2]
        phnNum2 = record[12]
        phnNum3 = record[13]
        instdueDate = record[6]
        instdueDate = str(instdueDate)
        instdueDate = instdueDate[:10]
        instNo = record[10]
        if emailid is None:
            emailid = ''
            print(emailid)
        elif "," in emailid:
            emailid = emailid.split(",")
            print(emailid)
        elif ";" in emailid:
            emailid = emailid.split(";")
            print(emailid)

        email, password = get_email_data()
        sender = email
        receivers = emailid

        msg = MIMEMultipart()
        msg['subject'] = "Reminder - Instalment Payment Due"
        message = """<p>Dear """ + str(clNam) + """</p>

        <p>The instalment payment of """ + str(premDue) + """ for the """ + str(prodType) + """ Policy-""" + str(
            polNum) + """ is due on """ + str(instdueDate) + """.</p>

        <p>Kindly make the payment on or before the above date to avoid the possible decrease of the policy’s coverage period.</p>

        <p>You can make the payment by visiting any of our offices, online banking through Scotiabank, Royal </p>

        <p>Bank of Canada and Commonwealth Bank or via our online service centre by clicking on the link <a href="https://jsjohnsononline.com/osc/osc/">https://jsjohnsononline.com/osc/osc/</a></p>

        <p style="color:red;"><b>Kindly, disregard this email if already paid</b></p>

        <p>Should you have any questions, you can contact us at (242) 397-2100 or visit any nearest office.</p>



        <p>Regards</p>
        <p>Team</p>
        <p>J.S JOHNSON</p>
        <p>Web: www.jsjohnson.com</p>
        <p>Tel: (242) 397-2100|Fax: (242) 323-3720</p>

        <p style="color:red;">**This is an auto-generated email, please do not respond to this</p>
        """
        msg.attach(MIMEText(message, 'html'))
        try:
            smtpObj = smtplib.SMTP('mail.jsjemail.com')
            smtpObj.login(email, password)
            smtpObj.sendmail(sender, receivers, msg.as_string())

            print("Successfully sent email")

        except smtplib.SMTPException as e:
            print("Error: unable to send email")
            print(e)
            worksheet.write(row, 0, brnch, fields)
            worksheet.write(row, 1, clId, fields)
            worksheet.write(row, 2, clNam, fields)
            worksheet.write(row, 3, polNum, fields)
            worksheet.write(row, 4, instNo, fields)
            worksheet.write(row, 5, premDue, fields)
            worksheet.write(row, 6, phnNum, fields)
            worksheet.write(row, 7, phnNum2, fields)
            worksheet.write(row, 8, phnNum3, fields)

            row += 1
    workbook.close()


tenTwo()




























