# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

from pycel.excelutil import AddressRange
from pycel.excelwrapper import (
    _OpxRange,
    ARRAY_FORMULA_FORMAT,
    ExcelOpxWrapperNoData,
)


def test_set_and_get_active_sheet(excel):
    excel.set_sheet("Sheet2")
    assert excel.get_active_sheet_name() == 'Sheet2'

    excel.set_sheet("Sheet3")
    assert excel.get_active_sheet_name() == 'Sheet3'


def test_get_range(excel):
    excel.set_sheet("Sheet2")
    excel_range = excel.get_range('Sheet2!A5:B7')
    assert excel_range.formula == (('', ''), ('', ''), ('', ''))
    assert sum(map(len, excel_range.values)) == 6


def test_get_used_range(excel):
    excel.set_sheet("Sheet1")
    assert sum(map(len, excel.get_used_range())) == 72


def test_get_formula_from_range(excel):
    excel.set_sheet("Sheet1")
    formulas = excel.get_formula_from_range("Sheet1!C2:C5")
    assert len(formulas) == 4
    assert formulas[1][0] == "=SIN(B3*A3^2)"

    formulas = excel.get_formula_from_range("Sheet1!C600:C601")
    assert formulas == ((None, ), (None, ))

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

    from_opxl = ExcelOpxWrapperNoData(excel.workbook)
    assert value == from_opxl.get_formula_or_value(address)


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
    expected = {'SINUS': [('$C$1:$C$18', 'Sheet1')]}
    assert expected == excel.defined_names

    assert excel.defined_names == excel.defined_names


def test_get_tables(excel):
    for table_name in ('Table1', 'tAbLe1'):
        table, sheet_name = excel.table(table_name)
        assert 'sref' == sheet_name
        assert 'D1:F4' == table.ref
        assert 'Table1' == table.name

    assert (None, None) == excel.table('JUNK')


@pytest.mark.parametrize(
    'address, table_name',
    [
        ('sref!D1', 'Table1'),
        ('sref!F1', 'Table1'),
        ('sref!D4', 'Table1'),
        ('sref!F4', 'Table1'),
        ('sref!F4', 'Table1'),
        ('sref!C1', None),
        ('sref!G1', None),
        ('sref!D5', None),
        ('sref!F5', None),
    ]
)
def test_table_name_containing(excel, address, table_name):
    table = excel.table_name_containing(address)
    if table_name is None:
        assert table is None
    else:
        assert table.lower() == table_name.lower()


@pytest.mark.parametrize(
    'address, values, formula',
    [
        ('ArrayForm!H1:I2', ((1, 2), (1, 2)),
         (('=INDEX(COLUMN(A1:B1),1,1,1,2)', '=INDEX(COLUMN(A1:B1),1,2,1,2)'),
          ('=INDEX(COLUMN(A1:B1),1,1)', '=INDEX(COLUMN(A1:B1),1,2)')),
         ),
        ('ArrayForm!E1:F3', ((1, 1), (2, 2), (3, 3)),
         (('=INDEX(ROW(A1:A3),1,1,3,1)', '=INDEX(ROW(A1:A3), 1)'),
          ('=INDEX(ROW(A1:A3),2,1,3,1)', '=INDEX(ROW(A1:A3), 2)'),
          ('=INDEX(ROW(A1:A3),3,1,3,1)', '=INDEX(ROW(A1:A3), 3)'))
         ),
        ('ArrayForm!E7:E9', ((11,), (10,), (16,)),
         (('=SUM((A7:A13="a")*(B7:B13="y")*C7:C13)',),
          ('=SUM((A7:A13<>"b")*(B7:B13<>"y")*C7:C13)',),
          ('=SUM((A7:A13>"b")*(B7:B13<"z")*(C7:C13+3.5))',))
         ),
        ('ArrayForm!G16:H17',
         ((1, 6), (6, 16)), '={A16:B17*D16:E17}'),
        ('ArrayForm!E21:F24',
         ((6, 6), (8, 8), (10, 10), (12, 12)), '={A21:A24+C21:C24}'
         ),
        ('ArrayForm!A32:D33',
         ((6, 8, 10, 12), (6, 8, 10, 12)), '={A28:D28+A30:D30}'
         ),
        ('ArrayForm!F28:I31',
         ((5, 6, 7, 8), (10, 12, 14, 16), (15, 18, 21, 24), (20, 24, 28, 32)),
         '={A21:A24*A30:D30}',
         ),
    ]
)
def test_array_formulas(excel, address, values, formula):
    result = excel.get_range(address)
    assert result.address == AddressRange(address)
    assert result.values == values
    if result.formula:
        assert result.formula == formula


def test_get_datetimes(excel):
    result = excel.get_range("datetime!A1:B13").values
    for row in result:
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


@pytest.mark.parametrize(
    'address, expecteds',
    (
        ('Sheet1!B2', ((1, '=B2=2'), (2, '=B2>1'), (4, '=B2>0'), (5, '=B2<0'))),
        ('Sheet1!B5', ((1, '=B5=2'), (2, '=B5>1'), (4, '=B5>0'), (5, '=B5<0'))),
        ('Sheet1!A1', ()),
    )
)
def test_conditional_format(cond_format_ws, address, expecteds):
    excel = cond_format_ws.excel
    results = excel.conditional_format(address)
    for result, expected in zip(results, expecteds):
        assert (result.priority, result.formula) == expected


@pytest.mark.parametrize(
    'value, formula',
    (
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 1, 1, 2, 2), '={xyzzy}'),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 1, 2, 2, 2), None),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 2, 1, 2, 2), None),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 2, 2, 2, 2), None),
    )
)
def test_cell_to_formulax(value, formula, ATestCell):
    cells = ((ATestCell('A', 1, value=value), ), )
    assert _OpxRange(cells, cells, '').formula == formula


@pytest.mark.parametrize(
    'value, formula',
    (
        (None, ""),
        ("xyzzy", ""),
        ("=xyzzy", "=xyzzy"),
        ("={1,2;3,4}", "=index({1,2;3,4},1,1)"),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 1, 1, 2, 2), "=index(s!E3:F4,1,1)"),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 1, 2, 2, 2), "=index(s!D3:E4,1,2)"),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 2, 1, 2, 2), "=index(s!E2:F3,2,1)"),
        (ARRAY_FORMULA_FORMAT % ('xyzzy', 2, 2, 2, 2), "=index(s!D2:E3,2,2)"),
    )
)
def test_cell_to_formula(value, formula):
    """"""
    from unittest import mock
    parent = mock.Mock()
    parent.title = 's'
    cell = mock.Mock()
    cell.value = value
    cell.row = 3
    cell.col_idx = 5
    cell.parent = parent
    assert _OpxRange.cell_to_formula(cell) == formula
