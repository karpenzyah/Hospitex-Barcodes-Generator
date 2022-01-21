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
import keyboard


def IsRunning():
    while True:
        for proc in psutil.process_iter():
            if proc.name() == "Reagent serial number generator_RUS.exe":
                return
        subprocess.Popen(r'\\Server\work\technical_support\Баркоды\Генератор бар-кода 100 (новый)\Reagent serial number generator_RUS.exe')
        time.sleep(3)
        
def ed_gen_bioelab(date):
    y = '20' + date[2:]
    m = date[:2]
    d = calendar.monthrange(int(y), int(m))[1]
    return y + '.' + str(d) + '.' + m

def export_to_file():
    export_ui = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name":"#32770","backend":"win32"},{"ctrl_index":16}])
    export_ui.click()
    keyboard.send('enter')
    keyboard.send('enter')

def bc_bioelab(bq, item, vol_r1, vol_r2, ed):
    #IsRunning()
    generator_ui = UIDesktop.UIOSelector_Get_UIO(
        [{"title":"Reagent Barcode Generator","class_name":"#32770","backend":"win32"}])
    generator_ui.set_focus()

    reagent_list_ui = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name":"#32770","backend":"win32"},{"ctrl_index":1}])
    reagent_list_ui.set_focus()
    #keyboard.press_and_release('down')
    reagent_list_ui.get_active()
    reagent_list_ui.select(item)


    ed_ui = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name":"#32770","backend":"win32"},{"ctrl_index":5}])
    ed_ui.click_input()
    ed_ui.type_keys(ed_gen_bioelab(ed))

    bottle_n_first = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name":"#32770","backend":"win32"},{"ctrl_index":8}])
    bottle_n_first.set_text('2222')

    bottle_n_last = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name":"#32770","backend":"win32"},{"ctrl_index":9}])
    bottle_n_last.set_text(str(2222+bq))

    type_ui = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name":"#32770","backend":"win32"},{"ctrl_index":10}])
    type_ui.select(0)

    vol_ui = UIDesktop.UIOSelector_Get_UIO(
        [{"class_name": "#32770", "backend": "win32"}, {"ctrl_index": 25}])
    vol_ui.set_text(vol_r1)

    add_to_list_ui = UIDesktop.UIOSelector_Get_UIO([{"class_name":"#32770","backend":"win32"},{"ctrl_index":0}])
    add_to_list_ui.click()

    if vol_r2 != 0:
        type_ui.select(1)
        vol_ui.set_text(vol_r2)
        add_to_list_ui.click()

    #barcode_list_ui = UIDesktop.UIOSelector_Get_UIO(
    #    [{"class_name": "#32770", "backend": "win32"}, {"ctrl_index": 11}])
    #return barcode_list_ui.texts()
    export_to_file()

res = bc_bioelab(bq=2, item=3, vol_r1=160, vol_r2=40, ed='1224')

res = bc_bioelab(bq=1, item=3, vol_r1=320, vol_r2=80, ed='1225')

