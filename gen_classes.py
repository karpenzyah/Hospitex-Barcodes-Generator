from configparser import *
from datetime import *
from pyOpenRPA.Robot import UIDesktop
import calendar

import pyodbc

class Generator:
    def __init__(self, dev_name, window_ui=None):
        self.dev_name = dev_name
        self.window_ui = window_ui

    def __repr__(self):
        return self.name

    def ui_select(self, *ui_indexes):
        ui_select_args = [self.window_ui]
        for ui_index in ui_indexes:
            ui_select_args.append(dict(ctrl_index=ui_index))
        return UIDesktop.UIOSelector_Get_UIO(ui_select_args)

    @classmethod
    def expiry_date(cls,date):
        y = '20'+date[2:]
        m  = date[:2]
        d = calendar.monthrange(int(y), int(m))[1]
        return str(d)+m+y

class HospitexDB:

    def __init__(self, db_name):
        conf = ConfigParser()
        conf.read("Settings.ini")
        self.path_to_db = conf['PathTo']['{0}Database'.format(db_name)]

    def db_request(self, request):
        connection = pyodbc.connect(
            "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.path_to_db};")
        cursor = connection.cursor()
        cursor.execute(request)
        results = []
        for row in cursor.fetchall():
            results.append(row)
        cursor.close()
        return results

#doc = docx.Document('Коды приборов.docx')

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
    trade_db = HospitexDB("Trade")
    goods_db = HospitexDB("Goods")
    columns = ['ITEM', 'R1Vol', 'R2Vol', 'URIT_ID', 'URIT_SIZE', 'TECOM_ID',
               'REF', 'ED', 'BQ']
    results = []
    res = trade_db.db_request(
        f"select KOD, PRIM_NAKL, QT from EXCEL_DATA where NOM_SH='{invoice_n}'"
        )
    params = []
    for row in res:
        params = goods_db.db_request(
            "SELECT ITEM, R1_VOL, R2_VOL, URIT_ID, URIT_SIZE, TECOM_ID "
            "FROM EXCEL_KART INNER JOIN BARCODE "
            "ON EXCEL_KART.KOD = BARCODE.KOD "
            "WHERE BARCODE.KOD = '%s'" % row[0])
        if len(params) == 0:
            continue
        else:
            results.append(dict(zip(columns, list(*params) + list(row))))
    return results

# # Returns list of invoices for the last 14 days or
# # customer name by invoice number
# def get_from_db(invoice_number=None):
#     cursor = _db_cursor('Trade')
#     if invoice_number is not None:
#         request = "select POLU from EXCEL_SHET " \
#                   f"where NOM_SH = '{invoice_number}'"
#     else:
#         now_date = datetime.now()
#         days_ago = now_date - timedelta(days=14)
#         request = ("select NOM_SH from EXCEL_SHET "
#                    "where DATA between {d'%(da)s'} and {d'%(nwd)s'}"
#                    % {'da': days_ago.date(), 'nwd': now_date.date()}
#                    )
#     cursor.execute(request)
#     results = []
#     for row in cursor.fetchall():
#         results.append(row[0])
#     cursor.close()
#     return results
#
#
# res=get_from_db()
res = get_invoice('1554-20')
print(res)

