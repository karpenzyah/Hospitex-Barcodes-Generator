from urit import UritGenerator

def test_urit_method():
    test = UritGenerator()
    test.size = '1'
    test.item_id = '1'
    test.vol = '250'
    test.ed = '30102021'
    test.bn = '181019'
    test.ref = '40001NNNN'
    test.item = 'DB'
    test.generate_barcode(1)
    assert test.barcodes[0]['bcs'][0] == '110010106861181019220582117352'

test_urit_method()