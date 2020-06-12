dt = (datetime.today()).strftime('%d%b%y')
file_name = 'C:/Users/administrator.JSJHO/Desktop/Consultant Files/Mantra/TempFolder/''AEEscalationReport' + '_' + dt + '.xlsx'
wb.save(
    'C:/Users/administrator.JSJHO/Desktop/Consultant Files/Mantra/TempFolder/''AEEscalationReport' + '_' + dt + '.xlsx')
return send_file(file_name)