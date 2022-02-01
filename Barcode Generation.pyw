from PyQt5 import QtWidgets, uic
from datetime import *
import sys
import os
from gen_classes import HospitexDB
from configparser import *
import docx


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('bars.ui', self)
        self.conf = ConfigParser()
        self.conf.read("Settings.ini", encoding="utf-8")
        self.params = self.conf['Parameters']

        self.dev_type_box.addItems(['Tecom', 'Urit', 'Bioelab'])
        self.dev_type_box.setCurrentText(self.params['Device'])
        self.invoices = self.last_invoices()

        for invoice in reversed(self.invoices):
            self.invoice_box.addItem(*invoice)
        self.invoice_box.setCurrentText("Выберите счет")

        self.get_hosps()

        self.invoice_rb.toggled.connect(self.by_invoice)
        self.manual_rb.toggled.connect(self.from_file)
        self.invoice_box.activated.connect(self.add_btn_enable)
        self.hosp_box.activated.connect(self.ok_btn_enable)
        self.add_btn.clicked.connect(self.add_invoice)
        self.dev_type_box.activated.connect(self.get_hosps)
        self.ok_btn.clicked.connect(self.start)
        self.show()

    def by_invoice(self):
        self.invoice_box.setEnabled(True)

    def from_file(self):
        self.invoice_box.setEnabled(False)
        self.ok_btn.setEnabled(False)
        self.add_btn.setEnabled(False)
        self.invoice_list.clear()
        self.invoice_list.setEnabled(False)

    def add_btn_enable(self):
        self.add_btn.setEnabled(True)

    def ok_btn_enable(self):
        self.ok_btn.setEnabled(True)

    def add_invoice(self):
        self.invoice_list.setEnabled(True)
        self.invoice_list.addItem(self.invoice_box.currentText())

    def start(self):
        # invoices = self.invoice_list.items()
        # self.params.set('invoice_n', f'{invoices}')
        # with open("Settings.ini", "w", encoding="utf-8") as config_file:
        #     self.conf.write(config_file)
        print(f'{self.dev_type_box.currentText().lower()}.py')
        os.system(f'python {self.dev_type_box.currentText().lower()}.py')
        self.close()

    def get_hosps(self):
        device = self.dev_type_box.currentText()
        self.hosp_box.clear()
        print(device)
        doc = docx.Document('Коды приборов.docx')
        dev_tables = {'Tecom': 0, 'Urit': 3, 'Bioelab': 1}

        table = doc.tables[dev_tables[device]]
        result = []
        for r in range(1, len(table.rows)):
            result.append(table.cell(r,0).text)

        self.hosp_box.addItems(result)
        self.hosp_box.setCurrentText("Выберите прибор")

    def last_invoices(self):
        trade_db = HospitexDB("trade")
        now_date = datetime.now()
        days_ago = now_date - timedelta(days=14)

        results = trade_db.db_request("select NOM_SH from EXCEL_SHET "
                                      "where DATA between {d'%(da)s'} and {d'%(nwd)s'}"
                                      % {'da': days_ago.date(),
                                         'nwd': now_date.date()}
                                      )
        return results


app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()
