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
#import warnings
#warnings.simplefilter('ignore', category=UserWarning)


def IsRunning():
    for proc in psutil.process_iter():
        try:
            if proc.name() == "Reagent serial number generator_RUS.exe":
                return
                break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    subprocess.Popen(r'\\Server\work\technical_support\Баркоды\Генератор бар-кода 100 (новый)\Reagent serial number generator_RUS.exe')
    time.sleep(3)
    

    #while True:
    #    for proc in psutil.process_iter():
    #        if proc.name() == "Reagent serial number generator_RUS.exe":
    #            return
    #    subprocess.Popen(r'\\Server\work\technical_support\Баркоды\Генератор бар-кода 100 (новый)\Reagent serial number generator_RUS.exe')
    #    time.sleep(3)



def BCgenTecom(BQ, ID, Vol, BN, ED, UID):
    IsRunning()

    GENERATOR_UI = UIDesktop.UIOSelector_Get_UIO([{"title":"To generate Reagent number verification","class_name":"ThunderRT6FormDC","backend":"win32"}])
    GENERATOR_UI.set_focus()
    
    ID_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":22}]) # Поле ввода кода реагента
    ID_UI.set_text(ID)
    
    BNF_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":0},{"ctrl_index":17}]) # Поле ввода начальн. серийн. нормера
    BNF_UI.set_text(str(BN))
   
    BNL_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":0},{"ctrl_index":9}]) # Поле ввода конечн. серийн. номера
    BNL_UI.set_text(str(BN+BQ-1))
    
    ED_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":13}])
    ED_UI.click_input()
    ED_UI.type_keys(ED)

    UID_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":0},{"ctrl_index":0}])
    UID_UI.set_text(UID)
    
    Vol_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":0},{"ctrl_index":19}])
    VolValues = Vol_UI.item_texts()
    
    iVals = [0]*len(VolValues)                                      # Вспомогательный целочисленный массив для поиска ближайшего значения объема 
    for val in VolValues:
        iVals[VolValues.index(val)] =  int(val[4:len(val)-2])       # Целочисленный массив объемов из выпадающего списка генератора
    SortedValues = sorted(enumerate(iVals), key=lambda i: i[1])     # Сортировка по возрастанию нумерованного кортежа значений объемов
    for i in range(len(SortedValues)):
        if SortedValues[i][1] >= Vol:                               # Поиск первого большего либо равного значения объема
            break
    Vol_UI.select(SortedValues[i][0])                               # Установить элемент выпадающего списка по старому индексу

    Generate_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":0},{"ctrl_index":13}])
    Generate_UI.click()

    Codes_UI = UIDesktop.UIOSelector_Get_UIO([{"class_name":"ThunderRT6FormDC","backend":"win32"},{"ctrl_index":0},{"ctrl_index":14},{"ctrl_index":1}])
    Code = (Codes_UI.window_text()).split('             Passed\r\n')

    BarCodes = []
    for i in range(0,len(Code)-1):
        BarCodes.append(Code[i][19:])
    return BarCodes

def BNgenTecom():
    nowdate = datetime.today()
    return int(str(nowdate.day)+str(nowdate.month+datetime.weekday(nowdate)))

def EDgenTecom(Date):
    Y = '20'+Date[2:]
    M = Date[:2]
    D = calendar.monthrange(int(Y), int(M))[1]
    return Y+'.'+str(D)+'.'+M    


def TecomGen(HOSP, UID, SN):
    source_dir = Path.cwd()
    outfile = dbf.Table(str(source_dir)+'\Bases\\outTecom.dbf', 'ITEM C(4); BC C(13); REF C(9); ED C(5); HOSP C(60);  SN C(21)', codepage='cp1251')
    outfile.open(dbf.READ_WRITE)

    csvfile = open(str(source_dir)+'\Tecom task.csv', newline='\n')
    reader = csv.DictReader(csvfile)
    
    BN = BNgenTecom()
    for prs in reader:
        if prs['BQ'] == '0000':
            continue
        else:
            CurrentItemBarcodes = BCgenTecom(int(prs['BQ']), prs['ID'], int(prs['Vol']), BN, EDgenTecom(prs['ED']), UID)
            print(CurrentItemBarcodes)
            for iBC in range(int(prs['BQ'])):
                outfile.append({'Item': prs['Item'], 'BC': CurrentItemBarcodes[iBC], 'REF': prs['REF'], 'ED': prs['ED'][:2]+'/'+prs['ED'][2:], 'HOSP': HOSP, 'SN': SN})    
    outfile.close()
    csvfile.close()
    
    if (os.stat(str(source_dir)+'\Bases\\outTecom.dbf').st_size)==0:
        return('TecomRobot - Нечего выполнять')
    else:
        return('TecomRobot - Дело сделано')

HOSP = 'ГБУЗ РХ "Абазинская городская больница'
UID = 'LICG900V03593059470366C61FAD'
SN = 'LICG900V035930594703'

source_dir = Path.cwd()
if __name__ == "__main__":
    result = TecomGen(HOSP, UID, SN)
    if result != 'TecomRobot - Нечего выполнять':
        os.system(str(source_dir)+'\Bases\\outTecom.dbf')
    else: print(result)

