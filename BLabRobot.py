from pathlib import *
from pyOpenRPA.Robot import UIDesktop
import csv
from datetime import datetime, date
import os
import dbf
import win32api
import calendar
import subprocess
import time
import psutil
import warnings
import keyboard
#warnings.simplefilter('ignore', category=UserWarning)


def IsRunning():
    while True:
        for proc in psutil.process_iter():
            if proc.name() == "Reagent serial number generator_RUS.exe":
                return
        subprocess.Popen(r'\\Server\work\technical_support\Баркоды\Генератор бар-кода 100 (новый)\Reagent serial number generator_RUS.exe')
        time.sleep(3)
        
def EDgenBLab(Date):
    Y = '20'+Date[2:]
    M = Date[:2]
    D = calendar.monthrange(int(Y), int(M))[1]
    return Y+'.'+str(D)+'.'+M  

#def BCgenBLab(BQ, ID, RT, Vol, BN, ED):
    #IsRunning()
GENERATOR_UI = UIDesktop.UIOSelector_Get_UIO([{"title":"Reagent Barcode Generator","class_name":"#32770","backend":"win32"}])
GENERATOR_UI.set_focus()

Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":0}]) # Объем
Generate_UI.click()

Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":13}]) # Объем
Generate_UI.click()

Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":7}]) # Объем
Generate_UI.click()

Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":2}]) # Объем
Generate_UI.click()

ID_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":1}]) # Поле ввода кода реагента
ID_UI.select(20) # Выбор реагента из списка

Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":0}]) # Объем
Generate_UI.click()

Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":16}]) # Объем
Generate_UI.click()

keyboard.send("enter")
keyboard.send("enter")

time.sleep(1)

BNF_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":8}]) #Начальн. партия
BNF_UI.set_text('2001')

BNL_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":9}]) # Конечн. партия
BNL_UI.set_text('2050')

ED_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":5}]) # Срок годности
ED_UI.click_input()
ED_UI.type_keys(EDgenBLab('0224'))

##RT_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":10}]) # Тип реагента R1 или R2
##RT_UI.select(0)
##
##Vol_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":25}]) # Объем
##Vol_UI.set_text('200')


##     BarCodes = []
##    for i in range(0,len(Code)-1):
##        BarCodes.append(Code[i][19:])
##    return BarCodes

##def BNgenTecom():
##    nowdate = datetime.today()
##    return int(str(nowdate.day)+str(nowdate.month+datetime.weekday(nowdate)))
##
  
##
##
##def TecomGen(HOSP, UID, SN):
##    source_dir = Path.cwd()
##    outfile = dbf.Table(str(source_dir)+'\Bases\\outTecom.dbf', 'ITEM C(4); BC C(13); REF C(9); ED C(5); HOSP C(60);  SN C(21)', codepage='cp1251')
##    outfile.open(dbf.READ_WRITE)

##    csvfile = open(str(source_dir)+'\Tecom job.csv', newline='\n')
##    #csvnewfile = open(str(source_dir) + HOSP+'.new.csv', 'w', newline='\n', )
##    reader = csv.DictReader(csvfile)
##    #writer = csv.DictWriter(csvnewfile, fieldnames = ['Item','ID','BQ','Size','Vol','BN','ED','UID'], delimiter=',')
##    #writer.writeheader()
##    BN = BNgenTecom()
##    for prs in reader:
##        if prs['BQ'] == '0000':
##            #writer.writerow({'Item': prs['Item'], 'ID': prs['ID'], 'BQ': prs['BQ'], 'Size': prs['Size'], 'Vol': prs['Vol'], 'BN': prs['BN'], 'ED': prs['ED'], 'UID': prs['UID']})
##            continue
##        else:
##            CurrentItemBarcodes = BCgenTecom(int(prs['BQ']), prs['ID'], int(prs['Vol']), BN, EDgenTecom(prs['ED']), UID)
##            print(CurrentItemBarcodes)
##            #time.sleep(1)
##            for iBC in range(int(prs['BQ'])):
##                #outfile.append({'BC': CurrentItemBarcodes[iBC], 'Item': prs['Item'], 'SN': prs['SN'], 'REF': prs['REF'], 'HOSP': HOSP })
##                outfile.append({'Item': prs['Item'], 'BC': CurrentItemBarcodes[iBC], 'REF': prs['REF'], 'ED': prs['ED'][:2]+'/'+prs['ED'][2:], 'HOSP': HOSP, 'SN': SN})
##         #   #writer.writerow({'Item': prs['Item'], 'ID': prs['ID'], 'BQ': '0000', 'Size': prs['Size'], 'Vol': prs['Vol'], 'BN': prs['BN'], 'ED': prs['ED'], 'UID': prs['UID']})    
##    outfile.close()
##    #csvnewfile.close()
##    csvfile.close()
##    
##    if (os.stat(str(source_dir)+'\Bases\\outTecom.dbf').st_size)==0:
##        return('TecomRobot - Нечего выполнять')
##    else:
##        return('TecomRobot - Дело сделано')
##
##    #win32api.WinExec('C:\\Users\\AM\\Desktop\\LVWIN60\\lv.exe C:\\Users\\AM\\Desktop\\Bars\\Bar.lbl')
##    #os.remove(str(source_dir) + HOSP+'.csv')
##    #os.rename(str(source_dir) + HOSP+'.new.csv', str(source_dir) + HOSP+'.csv')
##
##
##HOSP = 'НУЗ "Отделенческая"'
##UID = 'LICG900V03593059470366C61FAD'
##SN = 'LICG900V035930594703'
##
##source_dir = Path.cwd()
##if __name__ == "__main__":
##    result = TecomGen(HOSP, UID, SN)
##    if result != 'TecomRobot - Нечего выполнять':
##        os.system(str(source_dir)+'\Bases\\outTecom.dbf')
##    else: print(result)
##

