from pathlib import *
import pyodbc
from datetime import *
import csv

#def GetTask(Invoice, DevType):
#    source_dir = str(Path.cwd())
    
#    TaskCSVFile = open(source_dir + '\\' + DevType+'Task.csv', 'w', newline='\n', )
#    writer = csv.DictWriter(TaskCSVFile, fieldnames = ['Item','ID','BQ','Size','Vol','BN','ED','UID'], delimiter=',')
#    writer.writeheader()
#    writer.writerow({'Item': prs['Item'], 'ID': prs['ID'], 'BQ': prs['BQ'], 'Size': prs['Size'], 'Vol': prs['Vol'], 'BN': prs['BN'], 'ED': prs['ED'], 'UID': prs['UID']})
#     csvnewfile.close()
#        win32api.WinExec('C:\\Users\\AM\\Desktop\\LVWIN60\\lv.exe C:\\Users\\AM\\Desktop\\Bars\\Bar.lbl')
#    os.remove(str(source_dir) + HOSP+'.csv')

def DictFromMDB(invoice):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=X:\trade_database\HOSP_OFFICE_be.mdb;')
    cursor = conn.cursor() 
    cursor.execute("select KOD, PRIM_NAKL, QT from EXCEL_DATA WHERE NOM_SH='"+invoice+"'" )

    columns = ['REF','ED','BQ']
    results = []
    for row in cursor.fetchall():
        results.append(dict(zip(columns, row)))
    return(results)

def GetInvoices():
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=X:\trade_database\HOSP_OFFICE_be.mdb;')
    cursor = conn.cursor() 

    nowdate = datetime.now()
    daysago = nowdate - timedelta(days=14)
    print(nowdate.date(),daysago.date())
    cursor.execute("select DISTINCT NOM_SH from EXCEL_SHET WHERE DATA BETWEEN {d'%(da)s'} AND {d'%(nwd)s'}" %{'da':daysago.date(), 'nwd':nowdate.date()})
    results = []
    for row in cursor.fetchall():
        results.append(row[0])
    cursor.close()
    return(results)

def GetCustomer(invoiceNumber):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=X:\trade_database\HOSP_OFFICE_be.mdb;')
    cursor = conn.cursor() 

    cursor.execute("select distinct GRUZOPOL from EXCEL_DATA WHERE NOM_SH = '%(invc)s'" %{'invc':invoiceNumber})
    results = []
    for row in cursor.fetchall():
        results.append(row[0])
    cursor.close()
    return(results)

def GetReagentByREF(REF):
    ReagentFile = open('C:\\Users\SK\Desktop\Bars\REAGENT.csv', newline='\n')
    reader = csv.DictReader(ReagentFile, delimiter=';')
    i=-1
    for r in reader:
        if r['REF'] == REF:
            i=1
            break
    ReagentFile.close()
    if i == -1:
        return i
    else:
        return r
    

#res=DictFromMDB('0191-21')
#res=GetCustomer('0191-21')
#res = GetInvoices()
#res = GetReagentByREF('4001210LR')
#print(res)