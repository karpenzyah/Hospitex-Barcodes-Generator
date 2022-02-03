from configparser import *
from datetime import *
from pyOpenRPA.Robot import UIDesktop
import calendar
import subprocess
import time
import os
import csv
import dbf
import re

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
        self.dev_name = self.conf['Parameters']['device']
        self.hosp = self.conf['Parameters']['hospital']
        self.sn = self.conf['Parameters']['device_sn']
        self.uid = self.conf['Parameters']['device_uid']
        if window_ui is not None:
            self.window_ui = window_ui
            subprocess.Popen(self.conf['PathTo'][f'{self.dev_name.lower()}_generator'])
            time.sleep(2)

        self.bn = self.bn_gen()

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

    def write_to_dbf(self, path_to_file):
        dictrow = self.barcodes[0]
        bc_len = len(dictrow['bcs'][1][0])+1
        hosp_len = len(self.hosp)
        sn_len = len(self.sn)
        outfile = dbf.Table(f'{path_to_file}',
                            'ITEM C(5); '
                            f'BCR1 C({bc_len}); '
                            f'BCR2 C({bc_len}); '
                            'REF C(9); '
                            'ED C(5); '
                            f'HOSP C({hosp_len});  '
                            f'SN C({sn_len})',
                            codepage='cp1251')
        outfile.open(dbf.READ_WRITE)
        cnt = 0
        for prms in self.barcodes:
            ed = prms['ed'][:2] + '/' + prms['ed'][2:]
            r_flag = prms['bcs'][0]
            bcs = prms['bcs'][1]
            bc_r1 = [''] * r_flag.count('R1')
            bc_r2 = [''] * r_flag.count('R1')
            j = 0
            for i in range(len(bcs)):
                if r_flag[i] == "R1":
                    bc_r1[j] = bcs[i]
                    j += 1
                else:
                    bc_r2[j - 1] = bcs[i]
            for i in range(len(bc_r1)):
                outfile.append(
                    {'ITEM': prms['item'],
                     'BCR1': bc_r1[i],
                     'BCR2': bc_r2[i],
                     'REF': prms['ref'],
                     'ED': ed,
                     'HOSP': self.hosp,
                     'SN': self.sn})
                cnt += 1
        outfile.close()
        print(f'Cтрок записано в файл {path_to_file}: {cnt}')
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

    def gen_from_invoice(self):
        invoice_ns = self.conf['Parameters']['invoice_n'].split(',')
        trade_db = HospitexDB("Trade")
        goods_db = HospitexDB("Goods")
        parameters = []
        for invoice_n in invoice_ns:
            print('Подготовка задания по счету №', invoice_n)
            res = trade_db.db_request(
                "select KOD, PRIM_NAKL, QT, VIBOR "
                "from EXCEL_DATA "
                f"where NOM_SH='{invoice_n}' AND VIBOR = TRUE "
            )
            for row in res:
                # print(row)
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

                    if row[1] is None:
                        row[1] = input(
                            f'Введите срок годности для {ref, item} '
                            'в формате ММ/ГГГГ:  \n')
                    search = re.search('\d\d/\d{4}', row[1])
                    if search is None:
                        ed_r = input(
                            f'Введите срок годности для {ref, item} '
                            'в формате ММ/ГГГГ:  \n')
                    else:
                        ed_r = search.group(0)
                    ed = ed_r[:2] + ed_r[5:7]
                    bq = int(row[2])

                    parameters.append([item, ref, ed, bq])
        print(f'Выполнение задания, кол-во позиций: {parameters.__len__()}')
        for prs in parameters:
            self.generate_barcode(*prs)

    def start(self):
        task_from_invoice = self.conf['Parameters']['task_from_invoice']
        if task_from_invoice == 'Yes':
            self.gen_from_invoice()
        else:
            self.gen_from_taskfile()
        self.write_to_dbf(f'Bases\\out{self.dev_name}.dbf')

    @staticmethod
    def expiry_date(dt):
        y = '20' + dt[2:]
        m = dt[:2]
        d = calendar.monthrange(int(y), int(m))[1]
        return y + '.' + str(d) + '.' + m

    @staticmethod
    def bn_gen():
        now_date = datetime.today()
        return '{:02}'.format(now_date.day) + \
               '{:02}'.format(now_date.month + datetime.weekday(now_date))


class HospitexDB:

    def __init__(self, db_name):
        self.conf = ConfigParser()
        self.conf.read("Settings.ini")
        self.path_to_db = self.conf['PathTo']['{0}_database'.format(db_name)]

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
