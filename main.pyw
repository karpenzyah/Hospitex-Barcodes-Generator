from PyQt5 import QtWidgets, uic
import sys
import docx
import re
import csv
import subprocess
from TecomRobot import *
from UritGen import UritGen
from DBFunc import *

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('bars.ui', self)
        self.WhichGenerator = 100
        self.doc = docx.Document('Коды приборов.docx')
        self.ROW = 0

        self.invoices=GetInvoices()
        for invoice in reversed(self.invoices):
            #print(invoice)
            self.invoiceBox.addItem(invoice)

        self.invoiceBox.activated.connect(self.SetCustomer)
        self.searchBtn.clicked.connect(self.GetHosps)
        self.taskBtn.clicked.connect(self.GenTask)
        self.OkBtn.clicked.connect(self.BarcodeGeneration)
        self.show()


    def BarcodeGeneration(self):
        result = UritGen('ООО "БелАнта"', '2337069235', 'LICG900V013021A-81697')
        if result != 'UritGen - Нечего выполнять':
            os.system(str(source_dir)+'\Bases\out.dbf')
        else: print(result)
        #self.close()
        #source_dir = Path.cwd()
        
#        HOSP =table.cell(self.ROW,0).text
#        UID = table.cell(self.ROW,2).text
#        SN = table.cell(self.ROW,3).text
#        print(HOSP,UID,SN)
#        if self.WhichGenerator == 0:
#            strcode = TecomGen(HOSP,UID,SN)
#            print(strcode)
#            if strcode != 'TecomRobot - Нечего выполнять!':
#                   os.system(str(source_dir)+'\Bases\\outTecom.dbf')
###                    subprocess.call('C:\\Users\\AM\\Desktop\\LVWIN60\\Lvdbed.exe C:\\Users\\AM\\Desktop\\Bars\\outblab.dbf')
###                    subprocess.call('C:\\Users\\AM\\Desktop\\LVWIN60\\lv.exe')      
#        elif self.WhichGenerator == 3:
#            strcode = UritGen(HOSP,UID,SN)
#            print(strcode)
#            if strcode != 'UritGen - Нечего выполнять!':
#                   os.system(str(source_dir)+'\Bases\out.dbf')
#        else:
#            print('Непонятно что делать')
            
        
    
    def GetHosps(self):
        words = re.findall('\w+',self.customerEdit.text())
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

        self.HospsTable.setColumnCount(3)                               # Устанавливаем кол-во колонок
        self.HospsTable.setRowCount(len(newtable))                      # и строк в таблице
        self.HospsTable.setHorizontalHeaderLabels(['Тип прибора','Клиника','SN']) # заголовок таблицы

        for r in range(len(newtable)):
            for c in range(3):
                #print(r,c-1)
                item = QtWidgets.QTableWidgetItem(newtable[r][c], 0)
                self.HospsTable.setItem(r, c, item)
        self.HospsTable.resizeColumnsToContents()
        #self.HospsTable.cellClicked.connect(self.on_click)

    def SetCustomer(self):
        if GetCustomer(self.invoiceBox.currentText())[0] is None:
            self.customerEdit.setText('Отсутств.')
        else:    
            st = GetCustomer(self.invoiceBox.currentText())[0]
            i = st.find(',')
            if i != -1:
                stri = st[:i]
            self.customerEdit.setText(stri)  

    def on_click(self):
        self.ROW = self.HospsTable.currentRow()+1

    def GenTask(self):
        DevType = self.HospsTable.item(self.HospsTable.currentRow(),0).text()
        source_dir = str(Path.cwd())+'\\'+ DevType + ' task'

        TemplTaskFile = open(source_dir + ' template.csv', 'r', newline='\n', )
        TaskFile = open(source_dir + '.csv', 'w', newline='\n', )
        TemplReader = csv.DictReader(TemplTaskFile)
        writer = csv.DictWriter(TaskFile, fieldnames = TemplReader.fieldnames, delimiter=',')
        writer.writeheader()        

        Templ = []
        for row in TemplReader:
            Templ.append(row) 

        ReagDictRows = DictFromMDB(self.invoiceBox.currentText())
        for row in ReagDictRows:
            print(row)
            i = row['ED'].find('/')
            if i != -1:
                ED = row['ED'][i-2:i]+row['ED'][i+3:i+5]
            ID = ''
            Size=''
            reag = []

            reag = GetReagentByREF(row['REF'])
            if reag!=-1:
                print(reag)
                for k in range(len(Templ)):
                    if Templ[k]['Item']==reag['Item']:
                        ID = Templ[k]['ID']
                        Size = Templ[k]['Size']
            else: 
                continue

            if DevType == 'Tecom':
                writer.writerow({'Item': reag['Item'], 'ID': ID, 'BQ': int(row['BQ']), 'Vol': int(reag['VolR1'])+int(reag['VolR2']), 'ED': ED, 'REF': row['REF']})
            else:
                writer.writerow({ 'ID': ID, 'Item': reag['Item'], 'BQ': '{:04}'.format(int(row['BQ'])), 'Size': Size, 'Vol': '{:04}'.format(int(reag['VolR1'])+int(reag['VolR2'])), 'ED': ED, 'REF': row['REF']})
                #print(Trow['ID'], reag['Item'], int(row['BQ']), int(reag['VolR1'])+int(reag['VolR2']), row['ED'][i-2:i+5], row['REF'])

        TaskFile.close()
        TemplTaskFile.close()


        

  
app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()
