from pathlib import *
import csv
import datetime
import os
import dbf
import win32api
import calendar


def _bn_urit():
    nowdate = datetime.datetime.today()
    #print(nowdate)
    return '{:06}'.format(int(str(nowdate.day)+str(nowdate.month)+str(nowdate.year)[2:]))

def BCgenUrit(BQ, Size, ID, Vol, BN, ED, UID):
    D=[0]*31
    D[1] = int(Size)
    D[2] = int(str(BQ)[-1])
    D[3] = int(ID)//10
    D[4] = int(str(BQ//10)[-1])
    D[5] = int(ID)%10
    D[6] = int(str(BQ//100)[-1])
    D[7] = int(ED[-1])
    D[8] = int(str(BQ//1000)[-1])

    now = datetime.datetime(int(ED[-4:]),int(ED[2:4]),int(ED[:2]))
    then = datetime.datetime(int(ED[-4:])-1,12,31)
    D9_10 = (now - then).days//7+25
    D[9] = D9_10//10
    D[10] = D9_10%10

    D[13] = int(BN[0])
    D[14] = int(BN[1])
    D[15] = int(BN[2])
    D[16] = int(BN[3])
    D[17] = int(BN[4])
    D[18] = int(BN[5])
    D[19] = int(ED[-2])
    D[20] = (now - then).days % 7
    D[21] = int(Vol[-1])
    D[22] = int(Vol[-2])

    D[23] = int(UID[2])
    D[24] = int(UID[1])
    D[25] = int(UID[0])
    D[26] = int(UID[3])
    D[27] = int(UID[5])
    D[28] = int(UID[4])

    D[29] = int(Vol[1])
    D[30] = int(Vol[0])
    SN1 = 4*(10*D[1]+D[2])+(10*D[3]+D[4])+8*(10*D[5]+D[6])+9*(10*D[7]+D[8])+5*(10*D[9]+D[10])\
          +7*(10*D[13]+D[14])+7*(10*D[15]+D[16])+8*(10*D[17]+D[18])+9*(10*D[19]+D[20])\
          +6*(10*D[21]+D[22])+(10*D[23]+D[24])+(10*D[25]+D[26])+(10*D[27]+D[28])+2*(10*D[29]+D[30])
    SN2 = str((SN1%103)/1000+0.0001)[2:5]
    SN = SN2[1:]
            
    D[11] = int(SN[0])
    D[12] = int(SN[1])

    BarCode = ''
    for k in D[1:]:
        BarCode += str(k)

    print(BarCode, '|||', SN, BQ, BQ//10, BQ//100, BQ//1000)

    return(BarCode)


def UritGen(HOSP, UID, SN):
    source_dir = Path.cwd()
    
    outfile = dbf.Table(str(source_dir)+'\Bases\outUrit.dbf', 'BC C(30); ITEM C(4); HOSP C(60); SN C(21); ED C(5); REF C(9)',codepage='cp1251')
    outfile.open(dbf.READ_WRITE)

    csvfile = open(str(source_dir)+'\\Urit task.csv', newline='\n')
    reader = csv.DictReader(csvfile)
    for prs in reader:    
        if prs['BQ'] == '0000':
            continue
        else:
            for j in range(1,int(prs['BQ'])+1):
                CurrentItemBacrode = BCgenUrit(j, prs['Size'], prs['ID'], prs['Vol'], _bn_urit(), _ed_urit(prs['ED']), UID)
                outfile.append({'BC': CurrentItemBacrode, 'Item': prs['Item'], 'HOSP': HOSP, 'SN': SN, 'ED': prs['ED'][:2]+'/'+prs['ED'][2:], 'REF': prs['REF']})                
    outfile.close()
    csvfile.close()

    if (os.stat(str(source_dir)+'\Bases\outUrit.dbf').st_size)==0:
        return('UritGen - Нечего выполнять')
    else: return('UritGen - Дело сделано')

HOSP = 'Пятигорский онкодиспансер'
UID = '1277346393'
SN = '8021A81042'

print(_bn_urit())

source_dir = Path.cwd()
if __name__ == "__main__":
    result = UritGen(HOSP, UID, SN)
    if result != 'UritGen - Нечего выполнять':
        os.system(str(source_dir)+'\Bases\outUrit.dbf')
    else: print(result)

