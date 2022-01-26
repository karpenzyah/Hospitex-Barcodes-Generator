import csv
import os
import subprocess
import time
from datetime import datetime
from configparser import *

import dbf


from gen_classes import Generator


class TecomGenerator(Generator):

    # Ввод параметров в окно генератора и генерация штрих-кодов
    def generate_barcode(self, bq, item, item_id, vol, ed, uid, ref):
        window_generator_ui = self.ui_select()
        window_generator_ui.set_focus()

        item_ui = self.ui_select(22)  # Reagent code field
        item_ui.set_text(item_id)

        bn = self.bn_gen()
        bn_f_ui = self.ui_select(0, 17)  # Fist batch number
        bn_f_ui.set_text(str(bn))

        bn_l_ui = self.ui_select(0, 9)  # Last batch number
        bn_l_ui.set_text(str(int(bn) + bq - 1))

        ed_ui = self.ui_select(13)  # Expiry date field
        ed_ui.click_input()
        ed = self.expiry_date(ed)
        ed_ui.type_keys(ed)

        uid_ui = self.ui_select(0, 0)  # Device UID field
        uid_ui.set_text(uid)

        vol_ui = self.ui_select(0, 19)  # Reagent volume

        vol_values = vol_ui.item_texts()
        i_vals = [0] * len(vol_values)
        for val in vol_values:
            i_vals[vol_values.index(val)] = int(val[4:len(val) - 2])
        sorted_values = sorted(enumerate(i_vals), key=lambda k: k[1])
        stop_sorting_index = []
        for row in sorted_values:
            if row[1] >= vol:  # Closest volume value
                stop_sorting_index = sorted_values.index(row)
                break
        vol_ui.select(
            sorted_values[stop_sorting_index][0]
        )  # Setting volume by original index

        generate_btn_ui = self.ui_select(0, 13)
        generate_btn_ui.click()

        list_with_codes_ui = self.ui_select(0, 14, 1)
        code = list(filter(None, (list_with_codes_ui.window_text()).split(
            '             Passed\r\n')))
        for i in range(bq):
            code[i] = code[i][19:]

        self.barcodes.append({'item': item,
                              'bcs': code,
                              'ref': ref,
                              'ed': ed})


if __name__ == "__main__":
    tecom = TecomGenerator('Tecom',
                               {
                                   "title": "To generate Reagent number verification",
                                   "class_name": "ThunderRT6FormDC",
                                   "backend": "win32"})
    _hosp = '"Ветпомощь Оберег", г. Москва'
    _uid = 'LICG900V03L2-E1634BAC241E6'
    _sn = 'LICG900V03L2-E1634'

    task_file = open(r'Tecom task.csv', newline='\n')
    reader = csv.DictReader(task_file)
    for prs in reader:
        if prs['BQ'] == '0000':
            continue
        else:
            tecom.generate_barcode(int(prs['BQ']),
                                       prs['ID'],
                                       prs['Item'],
                                       int(prs['Vol']),
                                       prs['ED'],
                                       _uid,
                                       _hosp,
                                       _sn,
                                       prs['REF'])
    task_file.close()
    # os.system(r'Bases\outTecom.dbf')
    res1 = tecom.barcodes
    tecom.write_to_dbf('outTecom')
    print(res1)
