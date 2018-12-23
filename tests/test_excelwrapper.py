

def test_connect(unconnected_excel):
    try:
        unconnected_excel.connect()
        connected = True
    except:  # noqa: E722
        connected = False
    assert connected


def test_set_and_get_active_sheet(excel):
    excel.set_sheet("Sheet3")
    assert excel.get_active_sheet_name() == 'Sheet3'


def test_get_range(excel):
    excel.set_sheet("Sheet2")
    excel_range = excel.get_range('Sheet2!A5:B7')
    assert sum(map(len, excel_range.formulas)) == 6
    assert sum(map(len, excel_range.values)) == 6


def test_get_used_range(excel):
    excel.set_sheet("Sheet1")
    assert sum(map(len, excel.get_used_range())) == 72


def test_get_formula_from_range(excel):
    excel.set_sheet("Sheet1")
    formulas = excel.get_formula_from_range("Sheet1!C2:C5")
    assert len(formulas) == 4
    assert formulas[1] == "=SIN(B3*A3^2)"

    formulas = excel.get_formula_from_range("Sheet1!C600:C601")
    assert formulas is None

    formula = excel.get_formula_from_range("Sheet1!C3")
    assert formula == "=SIN(B3*A3^2)"


def test_get_formula_or_value(excel):
    result = excel.get_formula_or_value("Sheet1!A2:C2")
    assert (('2', '=SUM(A2:A4)', '=SIN(B2*A2^2)'),) == result

    result = excel.get_formula_or_value("Sheet1!A1:A3")
    assert (('1',), ('2',), ('3',)) == result


def test_get_range_formula(excel):
    result = excel.get_range("Sheet1!A2:C2").Formula
    assert (('2', '=SUM(A2:A4)', '=SIN(B2*A2^2)'),) == result

    result = excel.get_range("Sheet1!A1:A3").Formula
    assert (('1',), ('2',), ('3',)) == result

    result = excel.get_range("Sheet1!C2").Formula
    assert '=SIN(B2*A2^2)' == result

    excel.set_sheet('Sheet1')
    result = excel.get_range("C2").Formula
    assert '=SIN(B2*A2^2)' == result

    result = excel.get_range("Sheet1!AA1:AA3").Formula
    assert (('',), ('',), ('',)) == result

    result = excel.get_range("Sheet1!CC2").Formula
    assert '' == result


def test_get_value_with_formula(excel):
    result = excel.get_range("Sheet1!A2:C2").Value
    assert ((2, 9, -0.9917788534431158),) == result

    result = excel.get_range("Sheet1!A1:A3").Value
    assert ((1,), (2,), (3,)) == result

    result = excel.get_range("Sheet1!B2").Value
    assert 9 == result

    excel.set_sheet('Sheet1')
    result = excel.get_range("B2").Value
    assert 9 == result

    result = excel.get_range("Sheet1!AA1:AA3").Value
    assert ((None,), (None,), (None,)) == result

    result = excel.get_range("Sheet1!CC2").Value
    assert result is None


def test_get_range_value(excel):
    result = excel.get_range("Sheet1!A2:C2").Value
    assert ((2, 9, -0.9917788534431158),) == result

    result = excel.get_range("Sheet1!A1:A3").Value
    assert ((1,), (2,), (3,)) == result

    result = excel.get_range("Sheet1!A1").Value
    assert 1 == result

    result = excel.get_range("Sheet1!AA1:AA3").Value
    assert ((None,), (None,), (None,)) == result

    result = excel.get_range("Sheet1!CC2").Value
    assert result is None


def test_get_defined_names(excel):
    expected = {'SINUS': ('$C$1:$C$18', 'Sheet1')}
    assert expected == excel.defined_names

    assert excel.defined_names == excel.defined_names


def test_get_tables(excel):
    for table_name in ('Table1', 'tAbLe1'):
        table, sheet_name = excel.table(table_name)
        assert 'sref' == sheet_name
        assert 'D1:F4' == table.ref
        assert 'Table1' == table.name

    assert (None, None) == excel.table('JUNK')
