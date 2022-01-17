from datetime import *
from configparser import *

import pyodbc

# Reads path to DB from Settings.ini by db_name (can be 'Trade' or 'Goods') and connects to it, returns cursor
def _db_cursor(db_name): 
    conf = ConfigParser()
    conf.read("Settings.ini")
    path_to_db = conf['PathTo'][db_name+'Database']
    connection = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
                                f"DBQ={path_to_db};")
    return connection.cursor() 

# Returns only biochemical reagents invoice positions 
def get_invoice(invoiceN): 
    trade_cursor = _db_cursor('Trade')
    goods_cursor = _db_cursor('Goods')
    trade_cursor.execute(f"select KOD, PRIM_NAKL, QT from EXCEL_DATA where NOM_SH='{invoiceN}'")

    columns = ['ITEM',
               'R1Vol',
               'R2Vol',
               'URIT_ID', 
               'URIT_SIZE',
               'TECOM_ID',
               'REF',
               'ED',
               'BQ'
              ]
    results = []
    for row in trade_cursor.fetchall():
        goods_cursor.execute("select ITEM, R1_VOL, R2_VOL, URIT_ID, URIT_SIZE, TECOM_ID from EXCEL_KART WHERE KOD = '%s'" %row[0])
        params = goods_cursor.fetchone()
        if params[0] is None:
            continue
        else:
            results.append(dict(zip(columns, list(params)+list(row))))
    trade_cursor.close()
    goods_cursor.close()
    return results

# Returns list of invoices for the last 14 days Returns customer name by invoice number
def get_from_db(invoice_number=None):
    cursor = _db_cursor('Trade') 
    if invoice_number is not None:
        request = f"select POLU from EXCEL_SHET WHERE NOM_SH = '{invoice_number}'"
    else:
        nowdate = datetime.now()
        daysago = nowdate - timedelta(days=14)
        request = ("select NOM_SH from EXCEL_SHET WHERE DATA BETWEEN {d'%(da)s'} AND {d'%(nwd)s'}"
                  %{'da':daysago.date(), 'nwd':nowdate.date()})
    cursor.execute(request)
    results = []
    for row in cursor.fetchall():
        results.append(row[0])
    cursor.close()
    return(results)

#res=get_from_db()
res = get_invoice('1554-20')
print(res)