import calendar
import csv
import os
import subprocess
import time
from datetime import datetime
from configparser import *

import dbf
from pyOpenRPA.Robot import UIDesktop


def bn_gen_tecom():
    now_date = datetime.today()
    return '{:02}'.format(now_date.day) + \
           '{:02}'.format(now_date.month + datetime.weekday(now_date))


def ed_gen_tecom(date):
    y = '20' + date[2:]
    m = date[:2]
    d = calendar.monthrange(int(y), int(m))[1]
    return y + '.' + str(d) + '.' + m

# Как можно переписать это говно поизящнее?
def ui_select(ui_index_1=0, ui_index_2=None, ui_index_3=None):
    if ui_index_1 != 0:
        return UIDesktop.UIOSelector_Get_UIO(
            [{"class_name": "ThunderRT6FormDC", "backend": "win32"},
             {"ctrl_index": ui_index_1}])
    elif ui_index_2 is not None and ui_index_3 is None:
        return UIDesktop.UIOSelector_Get_UIO(
            [{"class_name": "ThunderRT6FormDC", "backend": "win32"},
             {"ctrl_index": ui_index_1}, {"ctrl_index": ui_index_2}])
    elif ui_index_3 is not None:
        return UIDesktop.UIOSelector_Get_UIO(
            [{"class_name": "ThunderRT6FormDC", "backend": "win32"},
             {"ctrl_index": ui_index_1}, {"ctrl_index": ui_index_2},
             {"ctrl_index": ui_index_3}])


def bc_gen_tecom(bq, id, item, vol, ed, uid, hosp, sn, ref):
    _conf = ConfigParser()
    _conf.read("Settings.ini", encoding="utf-8")
    subprocess.Popen(_conf['PathTo']['TecomGenerator'])
    time.sleep(3)
    bn = bn_gen_tecom()
    window_generator_ui = UIDesktop.UIOSelector_Get_UIO([{
        "title": "To generate Reagent number verification",
        "class_name": "ThunderRT6FormDC",
        "backend": "win32"}])
    window_generator_ui.set_focus()

    item_ui = ui_select(ui_index_1=22)  # Reagent code field
    item_ui.set_text(id)

    bn_f_ui = ui_select(ui_index_2=17)  # Fist batch number
    bn_f_ui.set_text(str(bn))

    bn_l_ui = ui_select(ui_index_2=9)  # Last batch number
    bn_l_ui.set_text(str(int(bn) + bq - 1))

    ed_ui = ui_select(ui_index_1=13)  # Expiry date field
    ed_ui.click_input()
    ed_ui.type_keys(ed_gen_tecom(ed))

    uid_ui = ui_select(ui_index_2=0)  # Device UID field
    uid_ui.set_text(uid)

    vol_ui = ui_select(ui_index_2=19)  # Reagent volume

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

    generate_btn_ui = ui_select(ui_index_2=13)
    generate_btn_ui.click()

    list_with_codes_ui = ui_select(ui_index_2=14, ui_index_3=1)
    code = (list_with_codes_ui.window_text()).split('             Passed\r\n')

    bar_codes = []
    for i in range(0, len(code) - 1):
        bar_codes.append(code[i][19:])
    print(bar_codes)
    outfile = dbf.Table(r'Bases\outTecom.dbf',
                        'ITEM C(4); '
                        'BC C(13); '
                        'REF C(9); '
                        'ED C(5); '
                        'HOSP C(60);  '
                        'SN C(21)',
                        codepage='cp1251')
    outfile.open(dbf.READ_WRITE)

    for i_bc in range(int(bq)):
        outfile.append(
            {'ITEM': item,
             'BC': bar_codes[i_bc],
             'REF': ref,
             'ED': ed[:2] + '/' + ed[2:],
             'HOSP': hosp,
             'SN': sn}
        )
    outfile.close()


if __name__ == "__main__":
    _hosp = 'ООО "Айболит", г. Иваново'
    _uid = 'LICG900V03L2-E1577B2CF180D'
    _sn = 'LICG900V03L2-E1577'

    task_file = open(r'Tecom task.csv', newline='\n')
    reader = csv.DictReader(task_file)
    for prs in reader:
        if prs['BQ'] == '0000':
            continue
        else:
            bc_gen_tecom(int(prs['BQ']),
                         prs['ID'],
                         prs['Item'],
                         int(prs['Vol']),
                         prs['ED'],
                         _uid,
                         _hosp,
                         _sn,
                         prs['REF'])
    task_file.close()
    os.system(r'Bases\outTecom.dbf')
