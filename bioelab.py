from pyOpenRPA.Robot import UIDesktop
from configparser import *
from gen_classes import Generator, HospitexDB


class BioelabGenerator(Generator):

    def generate_barcode(self, item, ref, ed, bq):
        goods_db = HospitexDB("Goods")
        prs = goods_db.db_request(
            "SELECT R1_VOL, R2_VOL, BIOELAB_ID "
            "FROM DEVICE_IDS INNER JOIN BARCODE "
            "ON DEVICE_IDS.ITEM = BARCODE.ITEM "
            "WHERE BARCODE.KOD = '%s'" % ref)

        item_id = prs[0][2]
        vol_r1 = prs[0][0]
        vol_r2 = prs[0][1]

        generator_ui = self.ui_select()
        generator_ui.set_focus()

        reagent_list_ui = self.ui_select(1)
        reagent_list_ui.select(item_id-1)

        ed_ui = self.ui_select(5)
        ed_ui.click_input()
        ed_ui.type_keys(self.expiry_date(ed))

        bn = self.bn_gen()
        first_bottle_n_ui = self.ui_select(8)
        first_bottle_n_ui.set_text(str(bn))

        last_bottle_n_ui = self.ui_select(9)
        last_bottle_n_ui.set_text(str(int(bn) + bq - 1))

        type_ui = self.ui_select(10)
        type_ui.select(0)

        vol_ui = self.ui_select(25)
        vol_ui.set_text(vol_r1)

        add_to_list_ui = self.ui_select(0)
        add_to_list_ui.click()

        if vol_r2 != '0':
            type_ui.select(1)
            vol_ui.set_text(vol_r2)
            add_to_list_ui.click()

        barcode_list_ui = self.ui_select(11)
        bcs = barcode_list_ui.texts()[2::8]
        r_flag = barcode_list_ui.texts()[4::8]
        bcs_r = [r_flag, bcs]
        print(item, ref, ed, bq, vol_r1, vol_r2, bcs)

        del_ui = self.ui_select(13)
        del_ui.click()
        del_win_ui = UIDesktop.UIOSelector_Get_UIO(
            [{"title":"Delete Barcode",
              "class_name":"#32770",
              "backend":"win32"},
             {"ctrl_index": 7}])
        del_win_ui.click()
        ok_win_ui = UIDesktop.UIOSelector_Get_UIO(
            [{"title":"Confirm","class_name":"#32770","backend":"win32"},
             {"ctrl_index":2}])
        ok_win_ui.click()

        self.barcodes.append({'item': item,
                              'bcs': bcs_r,
                              'ref': ref,
                              'ed': ed})


if __name__ == "__main__":
    gen = BioelabGenerator({"title": "Reagent Barcode Generator",
                            "class_name": "#32770",
                            "backend": "win32"})
    gen.start()
