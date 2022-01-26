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
import keyboard

from gen_classes import Generator


# [{"title": "Reagent Barcode Generator", "class_name": "#32770",
#               "backend": "win32"}])
class BioelabGenerator(Generator):

    def _generate_files(self, bq, item, vol_r1, vol_r2, ed, uid, hosp, sn,
                         ref):
        generator_ui = self.ui_select()
        generator_ui.set_focus()

        reagent_list_ui = self.ui_select(1)
        reagent_list_ui.select(item)

        ed_ui = self.ui_select(5)
        ed_ui.click_input()
        ed_ui.type_keys(self.expiry_date(ed))

        bn = self.bn_gen()
        first_bottle_n_ui = self.ui_select(8)
        first_bottle_n_ui.set_text(str(bn))

        last_bottle_n_ui = self.ui_select(9)
        last_bottle_n_ui.set_text(str(bn + bq))

        type_ui = self.ui_select(10)
        type_ui.select(0)

        vol_ui = self.ui_select(25)
        vol_ui.set_text(vol_r1)

        add_to_list_ui = self.ui_select(0)
        add_to_list_ui.click()

        if vol_r2 != 0:
            type_ui.select(1)
            vol_ui.set_text(vol_r2)
            add_to_list_ui.click()

        # barcode_list_ui = UIDesktop.UIOSelector_Get_UIO(
        #    [{"class_name": "#32770", "backend": "win32"}, {"ctrl_index": 11}])
        # return barcode_list_ui.texts()
        export_ui = UIDesktop.UIOSelector_Get_UIO(
            [{"class_name": "#32770", "backend": "win32"}, {"ctrl_index": 16}])
        export_ui.click()
        keyboard.send('enter')
        keyboard.send('enter')


    def read_files():
        source_dir = Path.cwd() / 'BioelabBar'
        outfile = dbf.Table('outblab.dbf',
                            'BCR1 C(23); BCR2 C(23); Reagent C(5); Type C(2) ; Vol C(6); HOSP C(30); DevSN C(25)')
        outfile.open(dbf.READ_WRITE)

        for curr_csv in source_dir.iterdir():
            with open(curr_csv) as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if row['Type'] == 'R1':
                        R1 = row['Barcode']
                    else:
                        outfile.append({'BCR1': R1, 'BCR2': row['Barcode'],
                                        'Reagent': row['Reagent'],
                                        'Type': row['Type'], 'Vol': row['Volume'],
                                        'HOSP': 'RCB',
                                        'DEVSN': 'LICG900V02AES2K0WH04053F'})
            if row['Type'] == 'R1':
                with open(curr_csv) as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        outfile.append(
                            {'BCR1': row['Barcode'], 'Reagent': row['Reagent'],
                             'Type': row['Type'], 'Vol': row['Volume'],
                             'HOSP': 'RCB', 'DEVSN': 'LICG900V02AES2K0WH04053F'})
            # os.remove(curr_csv)
        outfile.close()

        if (os.stat('outblab.dbf').st_size) < 369:
            return ('BLabRead - Нечего выполнять!')
        else:
            return ('BLabRead - Дело сделано!')


if __name__ == "__main__":
    BLabRead()
    os.system('outblab.dbf')
