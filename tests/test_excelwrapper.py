import datetime as dt
import pytest


def test_connect(unconnected_excel):
    try:
        unconnected_excel.connect()
        connected = True
    except:  # noqa: E722
        connected = False
    assert connected


def test_set_and_get_active_sheet(excel):
    excel.set_sheet("Sheet2")
    assert excel.get_active_sheet_name() == 'Sheet2'

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


@pytest.mark.parametrize(
    'address, value',
    [
        ("Sheet1!A2", 2),
        ("Sheet1!B2", '=SUM(A2:A4)'),
        ("Sheet1!A2:C2", ((2, '=SUM(A2:A4)', '=SIN(B2*A2^2)'),)),
        ("Sheet1!A1:A3", ((1,), (2,), (3,))),
        ("Sheet1!1:2", (
            (1, '=SUM(A1:A3)', '=SIN(B1*A1^2)', '=LINEST(C1:C18,B1:B18)'),
            (2, '=SUM(A2:A4)', '=SIN(B2*A2^2)', None))),
    ]
)
def test_get_formula_or_value(excel, address, value):
    assert value == excel.get_formula_or_value(address)


def test_get_range_formula(excel):
    result = excel.get_range("Sheet1!A2:C2").formulas
    assert (('', '=SUM(A2:A4)', '=SIN(B2*A2^2)'),) == result

    result = excel.get_range("Sheet1!A1:A3").formulas
    assert (('',), ('',), ('',)) == result

    result = excel.get_range("Sheet1!C2").formulas
    assert '=SIN(B2*A2^2)' == result

    excel.set_sheet('Sheet1')
    result = excel.get_range("C2").formulas
    assert '=SIN(B2*A2^2)' == result

    result = excel.get_range("Sheet1!AA1:AA3").formulas
    assert (('',), ('',), ('',)) == result

    result = excel.get_range("Sheet1!CC2").formulas
    assert '' == result


@pytest.mark.parametrize(
    'address1, address2',
    [
        ("Sheet1!1:2", "Sheet1!A1:D2"),
        ("Sheet1!A:B", "Sheet1!A1:B18"),
        ("Sheet1!2:2", "Sheet1!A2:D2"),
        ("Sheet1!B:B", "Sheet1!B1:B18"),
    ]
)
def test_get_unbounded_range(excel, address1, address2):
    assert excel.get_range(address1) == excel.get_range(address2)


def test_get_value_with_formula(excel):
    result = excel.get_range("Sheet1!A2:C2").values
    assert ((2, 9, -0.9917788534431158),) == result

    result = excel.get_range("Sheet1!A1:A3").values
    assert ((1,), (2,), (3,)) == result

    result = excel.get_range("Sheet1!B2").values
    assert 9 == result

    excel.set_sheet('Sheet1')
    result = excel.get_range("B2").values
    assert 9 == result

    result = excel.get_range("Sheet1!AA1:AA3").values
    assert ((None,), (None,), (None,)) == result

    result = excel.get_range("Sheet1!CC2").values
    assert result is None


def test_get_range_value(excel):
    result = excel.get_range("Sheet1!A2:C2").values
    assert ((2, 9, -0.9917788534431158),) == result

    result = excel.get_range("Sheet1!A1:A3").values
    assert ((1,), (2,), (3,)) == result

    result = excel.get_range("Sheet1!A1").values
    assert 1 == result

    result = excel.get_range("Sheet1!AA1:AA3").values
    assert ((None,), (None,), (None,)) == result

    result = excel.get_range("Sheet1!CC2").values
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


def test_get_datetimes(excel):
    result = excel.get_range("datetime!A1:B12").values
    for row in result:
        if isinstance(row[1], (dt.date, dt.datetime)):
            assert row[0] == row[1]


@pytest.mark.parametrize(
    'result_range, expected_range',
    [
        ("Sheet1!C:C", "Sheet1!C1:C18"),
        ("Sheet1!2:2", "Sheet1!A2:D2"),
        ("Sheet1!B:C", "Sheet1!B1:C18"),
        ("Sheet1!2:3", "Sheet1!A2:D3"),
    ]
)
def test_get_entire_rows_columns(excel, result_range, expected_range):

    result = excel.get_range(result_range).values
    expected = excel.get_range(expected_range).values
    assert result == expected
