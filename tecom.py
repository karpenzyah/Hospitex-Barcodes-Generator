from gen_classes import Generator, HospitexDB


class TecomGenerator(Generator):

    # Ввод параметров в окно генератора и генерация штрих-кодов
    def generate_barcode(self, item, ref, ed, bq):
        goods_db = HospitexDB("Goods")
        prs = goods_db.db_request(
            "SELECT R1_VOL, R2_VOL, TECOM_ID "
            "FROM DEVICE_IDS INNER JOIN BARCODE "
            "ON DEVICE_IDS.ITEM = BARCODE.ITEM "
            "WHERE BARCODE.KOD = '%s'" % ref)

        item_id = '{:02}'.format(prs[0][2])
        vol = int(prs[0][0])+int(prs[0][1])

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
        ed_ui.type_keys(self.expiry_date(ed))

        uid_ui = self.ui_select(0, 0)  # Device UID field
        uid_ui.set_text(self.uid)

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
    tecom = TecomGenerator({"title": "To generate Reagent number verification",
                            "class_name": "ThunderRT6FormDC",
                            "backend": "win32"})
    if tecom.dev_name != 'Tecom':
        print('Указан некорректный тип прибора')
    tecom.gen_from_taskfile()
    tecom.write_to_dbf('outTecomTest.dbf')
    c=1

