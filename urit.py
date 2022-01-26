import datetime
import win32api

from gen_classes import Generator


class UritGenerator(Generator):

    def _bn_urit(self):
        nowdate = datetime.datetime.today()
        # print(nowdate)
        return '{:06}'.format(int(str(nowdate.day) +
                                  str(nowdate.month) +
                                  str(nowdate.year)[2:]))

    def barcode_generation(self, bq, size, item, item_id, vol, ed, uid, ref):
        bn = self._bn_urit()
        barcodes = []
        for bottle in bq:
            d = [0] * 31
            d[1] = int(size)
            d[2] = int(str(bottle)[-1])
            d[3] = int(item_id) // 10
            d[4] = int(str(bottle // 10)[-1])
            d[5] = int(item_id) % 10
            d[6] = int(str(bottle // 100)[-1])
            d[7] = int(ed[-1])
            d[8] = int(str(bottle // 1000)[-1])

            now = datetime.datetime(int(ed[-4:]), int(ed[2:4]), int(ed[:2]))
            then = datetime.datetime(int(ed[-4:]) - 1, 12, 31)
            d9_10 = (now - then).days // 7 + 25
            d[9] = d9_10 // 10
            d[10] = d9_10 % 10

            d[13] = int(bn[0])
            d[14] = int(bn[1])
            d[15] = int(bn[2])
            d[16] = int(bn[3])
            d[17] = int(bn[4])
            d[18] = int(bn[5])
            d[19] = int(ed[-2])
            d[20] = (now - then).days % 7
            d[21] = int(vol[-1])
            d[22] = int(vol[-2])

            d[23] = int(uid[2])
            d[24] = int(uid[1])
            d[25] = int(uid[0])
            d[26] = int(uid[3])
            d[27] = int(uid[5])
            d[28] = int(uid[4])

            d[29] = int(vol[1])
            d[30] = int(vol[0])
            sn1 = 4 * (10 * d[1] + d[2]) + \
                  (10 * d[3] + d[4]) + \
                  8 * (10 * d[5] + d[6]) + \
                  9 * (10 * d[7] + d[8]) + \
                  5 * (10 * d[9] + d[10]) + \
                  7 * (10 * d[13] + d[14]) + \
                  7 * (10 * d[15] + d[16]) + \
                  8 * (10 * d[17] + d[18]) + \
                  9 * (10 * d[19] + d[20]) + \
                  6 * (10 * d[21] + d[22]) + \
                  (10 * d[23] + d[24]) + \
                  (10 * d[25] + d[26]) + \
                  (10 * d[27] + d[28]) + \
                  2 * (10 * d[29] + d[30])
            sn2 = str((sn1 % 103) / 1000 + 0.0001)[2:5]
            sn = sn2[1:]

            d[11] = int(sn[0])
            d[12] = int(sn[1])

            for _ in d[1:]:
                barcodes[int(bottle)] += str(_)

            print(barcodes[int(bottle)], '|||', sn, bottle, bottle // 10,
                  bottle // 100, bottle // 1000)

        self.barcodes.append({'item': item,
                              'bcs': barcodes,
                              'ref': ref,
                              'ed': ed})

# if __name__ == "__main__":
#     HOSP = 'Пятигорский онкодиспансер'
# UID = '1277346393'
# SN = '8021A81042'
#
#     result = UritGen(HOSP, UID, SN)
#     if result != 'UritGen - Нечего выполнять':
#         os.system(str(source_dir)+'\Bases\outUrit.dbf')
#     else:
#         print(result)
