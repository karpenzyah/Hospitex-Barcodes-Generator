from configparser import *
from datetime import *
from pyOpenRPA.Robot import UIDesktop
import calendar
import subprocess
import time
import os
import csv
import dbf

import pyodbc


# def IsRunning():
#     while True:
#         for proc in psutil.process_iter():
#             if proc.name() == "Reagent serial number generator_RUS.exe":
#                 return
#         subprocess.Popen(r'\\Server\work\technical_support\Баркоды\Генератор бар-кода 100 (новый)\Reagent serial number generator_RUS.exe')
#         time.sleep(3)


class Generator:
    def __init__(self, window_ui=None):
        self.conf = ConfigParser()
        self.conf.read("Settings.ini", encoding="utf-8")
        self.dev_name = self.conf['Parameters']['Device']
        self.hosp = self.conf['Parameters']['Hospital']
        self.sn = self.conf['Parameters']['DeviceSN']
        if self.dev_name != 'Bioelab':
            self.uid = self.conf['Parameters']['DeviceUID']
        if window_ui is not None:
            self.window_ui = window_ui
            subprocess.Popen(self.conf['PathTo'][f'{self.dev_name}Generator'])
            time.sleep(3)
        self.barcodes = []
        
        

    def __repr__(self):
        return self.dev_name

    def generate_barcode(self, item, ref, ed, bq):
        print('Метод генерации не определен')

    def ui_select(self, *ui_indexes):
        ui_select_args = [self.window_ui]
        for ui_index in ui_indexes:
            ui_select_args.append(dict(ctrl_index=ui_index))
        return UIDesktop.UIOSelector_Get_UIO(ui_select_args)

    @classmethod
    def expiry_date(cls, dt):
        y = '20' + dt[2:]
        m = dt[:2]
        d = calendar.monthrange(int(y), int(m))[1]
        return y + '.' + str(d) + '.' + m

    @classmethod
    def bn_gen(cls):
        now_date = datetime.today()
        return '{:02}'.format(now_date.day) + \
               '{:02}'.format(now_date.month + datetime.weekday(now_date))

    def write_to_dbf(self, path_to_file):
        dictrow = self.barcodes[0]
        bc_len = len(dictrow['bcs'][0])
        hosp_len = len(self.hosp)
        sn_len = len(self.sn)
        outfile = dbf.Table(f'{path_to_file}',
                            'ITEM C(5); '
                            f'BC C({bc_len+2}); '
                            'REF C(9); '
                            'ED C(5); '
                            f'HOSP C({hosp_len});  '
                            f'SN C({sn_len})',
                            codepage='cp1251')
        outfile.open(dbf.READ_WRITE)
        for prms in self.barcodes:
            for bc_r in prms['bcs']:
                ed = prms['ed'][:2] + '/' + prms['ed'][2:]
                outfile.append(
                    {'ITEM': prms['item'],
                     'BC': bc_r,
                     'REF': prms['ref'],
                     'ED': ed,
                     'HOSP': self.hosp,
                     'SN': self.sn}
                )
        outfile.close()
        os.system(path_to_file)

    def gen_from_taskfile(self):
        task_file = open('Task.csv', newline='\n')
        for prs in csv.DictReader(task_file):
            if prs['bq'] == '0':
                continue
            else:
                self.generate_barcode(prs['item'],
                                      prs['ref'],
                                      prs['ed'],
                                      int(prs['bq']))

    def gen_from_invoice(self, invoice_n):

        trade_db = HospitexDB("Trade")
        goods_db = HospitexDB("Goods")
        res = trade_db.db_request(
            f"select KOD, PRIM_NAKL, QT from EXCEL_DATA where NOM_SH='{invoice_n}' "
        )
        params = []
        for row in res:
            ref = row[0]
            params = goods_db.db_request(
                "SELECT ITEM "
                "FROM EXCEL_KART INNER JOIN BARCODE "
                "ON EXCEL_KART.KOD = BARCODE.KOD "
                "WHERE BARCODE.KOD = '%s'" % ref)
            if len(params) == 0:
                continue
            else:
                item = params[0][0]
                i = row[1].find('/')
                if i != -1:
                    ed = row[1][i - 2:i] + row[1][i + 3:i + 5]
                else:
                    print('Не удалось получить срок годности')
                bq = int(row[2])
                self.generate_barcode(item,
                                      ref,
                                      ed,
                                      bq)



class HospitexDB:

    def __init__(self, db_name):
        self.conf = ConfigParser()
        self.conf.read("Settings.ini")
        self.path_to_db = self.conf['PathTo']['{0}Database'.format(db_name)]

    def db_request(self, request):
        db = ("Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
              "DBQ=%s;" % self.path_to_db)
        connection = pyodbc.connect(db)
        cursor = connection.cursor()
        cursor.execute(request)
        results = []
        for row in cursor.fetchall():
            results.append(row)
        cursor.close()
        return results

# doc = docx.Document('Коды приборов.docx')

# def get_hosps(customer_text):
#     words = re.findall('\w+', customer_text)
#     newtable = []
#     DevTypes = ['Tecom', 'BioELab', '3 dif', 'Urit']
#
#     for word in words:
#         if len(word) < 4:
#             continue
#         print('Поиск по "' + word + '"')
#         for t in range(4):
#             table = self.doc.tables[t]
#             for r in range(1, len(table.rows)):
#                 flag = table.cell(r, 1).text.find(word)
#                 if flag != -1:
#                     newtable.append([DevTypes[t], table.cell(r, 0).text,
#                                      table.cell(r, 2).text,
#                                      table.cell(r, 3).text])
#     print('Поиск завершен\n')
#     for r in range(len(newtable)):
#         print(newtable[r])
#     return newtable
#
#
# def get_invoice(invoice_n):
#     trade_db = HospitexDB("Trade")
#     goods_db = HospitexDB("Goods")
#     columns = ['ITEM', 'R1Vol', 'R2Vol', 'URIT_ID', 'URIT_SIZE', 'TECOM_ID',
#                'REF', 'ED', 'BQ']
#     results = []
#     res = trade_db.db_request(
#         f"select KOD, PRIM_NAKL, QT from EXCEL_DATA where NOM_SH='{invoice_n}'"
#     )
#     params = []
#     for row in res:
#         params = goods_db.db_request(
#             "SELECT ITEM, R1_VOL, R2_VOL, URIT_ID, URIT_SIZE, TECOM_ID "
#             "FROM EXCEL_KART INNER JOIN BARCODE "
#             "ON EXCEL_KART.KOD = BARCODE.KOD "
#             "WHERE BARCODE.KOD = '%s'" % row[0])
#         if len(params) == 0:
#             continue
#         else:
#             results.append(dict(zip(columns, list(*params) + list(row))))
#     return results


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

# trade_db = HospitexDB("Trade")
# res = trade_db.db_request(
#         f"select KOD, PRIM_NAKL, QT from EXCEL_DATA where NOM_SH='2338-21'"
#     )
# goods_db = HospitexDB("Goods")
# params = goods_db.db_request(
#             "SELECT ITEM "
#             "FROM EXCEL_KART INNER JOIN BARCODE "
#             "ON EXCEL_KART.KOD = BARCODE.KOD "
#             "WHERE BARCODE.KOD = '4001008NR'"
# )
# #res = get_invoice('2396-21')
# print(res)
# print(params)
