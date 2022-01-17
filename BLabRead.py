from pathlib import *
import csv
import dbf
import os

def BLabRead():
    source_dir = Path.cwd()/'BioelabBar'
    outfile = dbf.Table('outblab.dbf', 'BCR1 C(23); BCR2 C(23); Reagent C(5); Type C(2) ; Vol C(6); HOSP C(30); DevSN C(25)')
    outfile.open(dbf.READ_WRITE)

    for curr_csv in source_dir.iterdir():
        with open(curr_csv) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['Type']=='R1':
                    R1=row['Barcode']
                else:
                    outfile.append({'BCR1': R1, 'BCR2': row['Barcode'], 'Reagent': row['Reagent'], 'Type': row['Type'], 'Vol': row['Volume'], 'HOSP': 'RCB', 'DEVSN': 'LICG900V02AES2K0WH04053F'})
        if row['Type']=='R1':
            with open(curr_csv) as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    outfile.append({'BCR1': row['Barcode'], 'Reagent': row['Reagent'], 'Type': row['Type'], 'Vol': row['Volume'], 'HOSP': 'RCB', 'DEVSN': 'LICG900V02AES2K0WH04053F'}) 
        #os.remove(curr_csv)
    outfile.close()

    if (os.stat('outblab.dbf').st_size)<369:
        return('BLabRead - Нечего выполнять!')
    else:
        return('BLabRead - Дело сделано!')


if __name__ == "__main__":
    BLabRead()
    os.system('outblab.dbf')
