from configparser import *
from datetime import *

import pyodbc

class generator:
    def __init__(self, dev_name):
        self.name = dev_name

    def __repr__(self):
        return self.name

    @staticmethod
    def expiry_date(self.date):
        Y = '20'+Date[2:]
        M = Date[:2]
        D = calendar.monthrange(int(Y), int(M))[1]
        return str(D)+M+Y

class HospitexDB:

    def __init__(self, db_name):
        conf = ConfigParser()
        conf.read("Settings.ini")
        self.path_to_db = conf['PathTo'][db_name + 'Database']

    def db_request(self, request):
        connection = pyodbc.connect(
            "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.path_to_db};")
        cursor = connection.cursor()
        cursor.execute(request)
        results = []
        for row in cursor.fetchall():
            results.append(row[0])
        cursor.close()
        return results

doc = docx.Document('Коды приборов.docx')

def get_hosps(customer_text):

        words = re.findall('\w+', customer_text)
        newtable = []
        DevTypes = ['Tecom','BioELab','3 dif','Urit']

        for word in words:
            if len(word)<4:
                continue
            print('Поиск по "'+word+'"')
            for t in range(4):
                table = self.doc.tables[t]
                for r in range(1,len(table.rows)):
                    flag=table.cell(r,1).text.find(word)
                    if flag != -1:
                        newtable.append([DevTypes[t],table.cell(r,0).text,table.cell(r,2).text,table.cell(r,3).text])
        print('Поиск завершен\n')
        for r in range(len(newtable)):
            print(newtable[r])
        return newtable

def get_invoice(invoice_n):
    trade_cursor = _db_cursor('Trade')
    goods_cursor = _db_cursor('Goods')
    trade_cursor.execute(
        f"select KOD, PRIM_NAKL, QT from EXCEL_DATA where NOM_SH='{invoice_n}'"
    )

    columns = [
        'ITEM',
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
        goods_cursor.execute(
            "select ITEM, R1_VOL, R2_VOL, URIT_ID, URIT_SIZE, TECOM_ID "
            "from EXCEL_KART where KOD = '%s'" % row[0])
        params = goods_cursor.fetchone()
        if params[0] is None:
            continue
        else:
            results.append(dict(zip(columns, list(params) + list(row))))
    trade_cursor.close()
    goods_cursor.close()
    return results


# Returns list of invoices for the last 14 days or
# customer name by invoice number
def get_from_db(invoice_number=None):
    cursor = _db_cursor('Trade')
    if invoice_number is not None:
        request = "select POLU from EXCEL_SHET " \
                  f"where NOM_SH = '{invoice_number}'"
    else:
        now_date = datetime.now()
        days_ago = now_date - timedelta(days=14)
        request = ("select NOM_SH from EXCEL_SHET "
                   "where DATA between {d'%(da)s'} and {d'%(nwd)s'}"
                   % {'da': days_ago.date(), 'nwd': now_date.date()}
                   )
    cursor.execute(request)
    results = []
    for row in cursor.fetchall():
        results.append(row[0])
    cursor.close()
    return results


# res=get_from_db()
res = get_invoice('1554-20')
print(res)
