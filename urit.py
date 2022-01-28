import datetime
import calendar

from gen_classes import Generator, HospitexDB


class UritGenerator(Generator):

    def __init__(self):
        super(UritGenerator, self).__init__()
        self.size = None
        self.item_id = None
        self.vol = None
        self.ed = None
        self.bn = self.bn_gen()
        self.ref = None
        self.item = None

    def init_params(self, ref, ed):
        goods_db = HospitexDB("Goods")
        prs = goods_db.db_request(
            "SELECT R1_VOL, R2_VOL, URIT_ID, URIT_SIZE, DEVICE_IDS.ITEM  "
            "FROM DEVICE_IDS INNER JOIN BARCODE "
            "ON DEVICE_IDS.ITEM = BARCODE.ITEM "
            "WHERE BARCODE.KOD = '%s'" % ref)
        self.size = prs[0][3]
        self.item_id = prs[0][2]
        self.vol = str(int(prs[0][0]) + int(prs[0][1]))
        self.ed = self.expiry_date(ed)
        self.bn = self.bn_gen()
        self.ref = ref
        self.item = prs[0][4]

    @classmethod
    def bn_gen(cls):
        nowdate = datetime.datetime.today()
        return '{:06}'.format(int(str(nowdate.day) +
                                  str(nowdate.month) +
                                  str(nowdate.year)[2:]))

    @classmethod
    def expiry_date(cls, dt):
        y = '20' + dt[2:]
        m = dt[:2]
        d = calendar.monthrange(int(y), int(m))[1]
        return str(d) + m + y

    def generate_barcode(self, bq):

        bcs = [''] * bq
        for b in range(1, bq + 1):
            d = [0] * 31
            d[1] = int(self.size)
            d[2] = int(str(b)[-1])
            d[3] = int(self.item_id) // 10
            d[4] = int(str(b // 10)[-1])
            d[5] = int(self.item_id) % 10
            d[6] = int(str(b // 100)[-1])
            d[7] = int(self.ed[-1])
            d[8] = int(str(b // 1000)[-1])

            now = datetime.datetime(int(self.ed[-4:]),
                                    int(self.ed[2:4]),
                                    int(self.ed[:2]))
            then = datetime.datetime(int(self.ed[-4:]) - 1, 12, 31)
            d9_10 = (now - then).days // 7 + 25
            d[9] = d9_10 // 10
            d[10] = d9_10 % 10

            d[13] = int(self.bn[0])
            d[14] = int(self.bn[1])
            d[15] = int(self.bn[2])
            d[16] = int(self.bn[3])
            d[17] = int(self.bn[4])
            d[18] = int(self.bn[5])
            d[19] = int(self.ed[-2])
            d[20] = (now - then).days % 7
            d[21] = int(self.vol[-1])
            d[22] = int(self.vol[-2])

            d[23] = int(self.uid[2])
            d[24] = int(self.uid[1])
            d[25] = int(self.uid[0])
            d[26] = int(self.uid[3])
            d[27] = int(self.uid[5])
            d[28] = int(self.uid[4])

            d[29] = int(self.vol[1])
            d[30] = int(self.vol[0])
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
                bcs[b - 1] += str(_)

            print(bcs[b - 1], '|||', sn, b, b // 10,
                  b // 100, b // 1000)

        self.barcodes.append({'item': self.item,
                              'bcs': bcs,
                              'ref': self.ref,
                              'ed': self.ed})


if __name__ == "__main__":
    urit = UritGenerator()
    if urit.dev_name != 'Urit':
        print('Указан некорректный тип прибора')
    urit.gen_from_taskfile()
    urit.write_to_dbf('outTecomTest.dbf')
    c = 1
