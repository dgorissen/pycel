import os


def test_connect(unconnected_excel):
    try:
        unconnected_excel.connect()
        connected = True
    except:
        connected = False
    assert connected


def test_save_as(excel, tmpdir):
    path_copy = os.path.join(str(tmpdir), "exampleCopy.xlsx")
    if os.path.exists(path_copy):
        os.remove(path_copy)
    excel.save_as(path_copy)
    assert os.path.exists(path_copy)


def test_set_and_get_active_sheet(excel):
    excel.set_sheet("Sheet3")
    assert excel.get_active_sheet() == 'Sheet3'


def test_get_range(excel):
    excel.set_sheet("Sheet2")
    excel_range = excel.get_range('Sheet2!A5:B7')
    assert sum(map(len, excel_range.cells)) == 6


def test_get_used_range(excel):
    excel.set_sheet("Sheet1")
    assert sum(map(len, excel.get_used_range())) == 72


def test_get_value(excel):
    excel.set_sheet("Sheet1")
    assert excel.get_value(2, 2) == 9


def test_get_formula(excel):
    excel.set_sheet("Sheet1")
    assert excel.get_formula(2, 2) == "=SUM(A2:A4)"
    assert excel.get_formula(3, 12) is None


def test_has_formula(excel):
    excel.set_sheet("Sheet1")
    assert excel.has_formula("Sheet1!C2:C5")
    assert not excel.has_formula("Sheet1!A2:A5")


def test_get_formula_from_range(excel):
    excel.set_sheet("Sheet1")
    formulas = excel.get_formula_from_range("Sheet1!C2:C5")
    assert len(formulas) == 4
    assert formulas[1] == "=SIN(B3*A3^2)"


def test_get_formula_or_value(excel):
    excel.set_sheet("Sheet1")
    result = excel.get_formula_or_value("Sheet1!A2:C2")
    assert (('2', '=SUM(A2:A4)', '=SIN(B2*A2^2)'),) == result
    result = excel.get_formula_or_value("Sheet1!A1:A3")
    assert (('1',), ('2',), ('3',)) == result


def test_get_row(excel):
    excel.set_sheet("Sheet1")
    assert len(excel.get_row(2)) == 4


def test_get_ranged_names(excel):
    assert sum(map(len, excel.rangednames)) == sum(
        map(len, [[(1, 'SINUS', 'Sheet1!$C$1:$C$18')]]))
