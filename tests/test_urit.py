from urit import UritGenerator

def test_urit_method():
    test = UritGenerator()
    test.size = '1'
    test.item_id = '15'
    test.vol = '400'
    test.ed = '15122024'
    test.bn = '101235'
    test.ref = '40001NNNN'
    test.item = 'AST'
    test.generate_barcode(3)
    print(test.barcodes)
    assert test.barcodes[0][0] == '12035646464654615351'
