import csv
import os
import subprocess
import time
from datetime import datetime
from configparser import *

import dbf
from pyOpenRPA.Robot import UIDesktop

from gen_classes import Generator


class TecomGenerator(Generator):

    def __init__(self, dev_name, window_ui):
        self.dev_name = dev_name
        self.window_ui = window_ui
        conf = ConfigParser()
        conf.read("Settings.ini", encoding="utf-8")
        subprocess.Popen(conf['PathTo'][f'{self.dev_name}Generator'])
        time.sleep(3)
        self.barcodes = []
        self.gen_params = []

    @classmethod
    def bn_gen(cls):
        now_date = datetime.today()
        return '{:02}'.format(now_date.day) + \
               '{:02}'.format(now_date.month + datetime.weekday(now_date))

    def generate_barcode(self, bq, id, item, vol, ed, uid, hosp, sn, ref):
        bn = TecomGenerator.bn_gen()
        window_generator_ui = self.ui_select()
        window_generator_ui.set_focus()

        item_ui = self.ui_select(22)  # Reagent code field
        item_ui.set_text(id)

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
        for i in range(len(sorted_values)):
            if sorted_values[i][1] >= vol:  # Closest volume value
                stop_sorting_index = i
                break
        vol_ui.select(
            sorted_values[stop_sorting_index][0]
        )  # Setting volume by original index

        generate_btn_ui = self.ui_select(0, 13)
        generate_btn_ui.click()

        list_with_codes_ui = self.ui_select(0, 14, 1)
        code = (list_with_codes_ui.window_text()).split(
            '             Passed\r\n')

        self.gen_params.append({'item': item,
                               'ref': ref,
                               'ed': ed,
                               'hosp': hosp,
                               'sn': sn})
        self.barcodes.append(code[:][19:])


    # def write_to_dbf(self, 'path_to_file')
    #     outfile = dbf.Table(f'{path_to_file}',
    #                         'ITEM C(4); '
    #                         'BC C(13); '
    #                         'REF C(9); '
    #                         'ED C(5); '
    #                         'HOSP C(60);  '
    #                         'SN C(21)',
    #                         codepage='cp1251')
    #     outfile.open(dbf.READ_WRITE)
    #
    #     for i_bc in range(int(bq)):
    #         outfile.append(
    #             {'ITEM': item,
    #              'BC': bar_codes[i_bc],
    #              'REF': ref,
    #              'ED': ed[:2] + '/' + ed[2:],
    #              'HOSP': hosp,
    #              'SN': sn}
    #         )
    #     outfile.close()

if __name__ == "__main__":
    tecom_gen = TecomGenerator('Tecom',
                               {"title": "To generate Reagent number verification",
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
            tecom_gen.generate_barcode(int(prs['BQ']),
                                       prs['ID'],
                                       prs['Item'],
                                       int(prs['Vol']),
                                       prs['ED'],
                                       _uid,
                                       _hosp,
                                       _sn,
                                       prs['REF'])
    task_file.close()
    #os.system(r'Bases\outTecom.dbf')
    res = tecom_gen.gen_params
    res1 = tecom_gen.barcodes
    print(res)
    print(res1)
