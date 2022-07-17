# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import numpy as np
import pytest

import pycel.lib
from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import (
    AddressCell,
    AddressRange,
    DIV0,
    ExcelCmp,
    is_address,
    NA_ERROR,
    NUM_ERROR,
    REF_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import error_string_wrapper, load_to_test_module
from pycel.lib.lookup import (
    _match,
    address,
    choose,
    column,
    columns,
    hlookup,
    index,
    indirect,
    lookup,
    match,
    offset,
    row,
    rows,
    vlookup,
)


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.lookup, __name__)


def test_lookup_ws(fixture_xls_copy):
    INDIRECT_FORMULA_ADDRESS = AddressCell('Offset!B53')
    compiler = ExcelCompiler(fixture_xls_copy('lookup.xlsx'))

    # do an INDIRECT() before other cells are loaded to verify it can load what it needs
    result = compiler.validate_calcs([INDIRECT_FORMULA_ADDRESS])
    assert result == {}

    # now load and check everything
    result = compiler.validate_serialized()
    assert result == {}

    # use indirect to an existing range
    loaded = ExcelCompiler.from_file(compiler.filename)
    loaded.set_value(INDIRECT_FORMULA_ADDRESS.address_at_offset(1, 0), 'B2:F6')
    indirect = loaded.evaluate(INDIRECT_FORMULA_ADDRESS)
    assert indirect == loaded.evaluate('Offset!B2')

    # use indirect to a non-pre-existing and empty range
    loaded.set_value(INDIRECT_FORMULA_ADDRESS.address_at_offset(1, 0), 'H1:H2')
    indirect = loaded.evaluate(INDIRECT_FORMULA_ADDRESS)
    assert indirect is None

    # use indirect to a non-pre-existing range to existing cells
    loaded.set_value(INDIRECT_FORMULA_ADDRESS.address_at_offset(1, 0), 'D3:E3')
    indirect = loaded.evaluate(INDIRECT_FORMULA_ADDRESS)
    assert indirect == 8


@pytest.mark.parametrize(
    'row_num, col_num, abs_num, style, sheet_text, expected',
    (
        (2, 3, 1, None, '', '$C$2'),
        (2, 3, 3, None, '', '$C2'),
        (2, 3, 2, None, '', 'C$2'),
        (2, 3, 2, False, '', 'R2C[3]'),
        (2, 3, 2, True, '', 'C$2'),
        (5, 4, 4, True, 'Sheet1', '\'Sheet1\'!D5'),
        (5, 4, 1, True, 'Sheet1', '\'Sheet1\'!$D$5'),
        (5, 4, 1, None, 'Sheet1', '\'Sheet1\'!$D$5'),
        (5, 4, 1, False, 'Sheet1', '\'Sheet1\'!R5C4'),
    )
)
def test_address(row_num, col_num, abs_num, style, sheet_text, expected):
    assert address(row_num, col_num, abs_num, style, sheet_text) == expected


@pytest.mark.parametrize(
    'index, data, expected', (
        (-1, 'ABCDEFG', VALUE_ERROR),
        (0, 'ABCDEFG', VALUE_ERROR),
        (1, 'ABCDEFG', 'A'),
        (2, 'ABCDEFG', 'B'),
        (7, 'ABCDEFG', 'G'),
        (8, 'ABCDEFG', VALUE_ERROR),
        (DIV0, 'ABCDEFG', DIV0),
        (NUM_ERROR, 'ABCDEFG', NUM_ERROR),
        (VALUE_ERROR, 'ABCDEFG', VALUE_ERROR),
        (1, (), VALUE_ERROR),
        (False, 'ABCDEFG', VALUE_ERROR),
        (True, 'ABCDEFG', 'A'),
        ('0', 'ABCDEFG', VALUE_ERROR),
        ('1', 'ABCDEFG', 'A'),
        ('1.5', 'ABCDEFG', 'A'),
        (1.5, 'ABCDEFG', 'A'),
    )
)
def test_choose(index, data, expected):
    assert choose(index, *data) == expected


@pytest.mark.parametrize(
    'address, expected', (
        ('L45', 12),
        ('B:E', ((2, 3, 4, 5), )),
        ('4:7', None),
        ('D1:E1', ((4, 5), )),
        ('D1:D2', ((4, ), )),
        ('D1:E2', ((4, 5), )),
        (DIV0, DIV0),
        (NUM_ERROR, NUM_ERROR),
        (VALUE_ERROR, VALUE_ERROR),
    )
)
def test_column(address, expected):
    try:
        address = AddressRange.create(address)
    except ValueError:
        pass

    result = column(address)
    if expected is None:
        assert 1 == next(iter(result))
    else:
        assert expected == result


@pytest.mark.parametrize(
    'values, expected', (
        (((1, None, None), (1, 2, None)), 3),
        (1, 1),
        ("s", 1),
        (((1.2, 3.4), (0.4, 5)), 2),
        (((None, None, None, None,), ), 4)
    )
)
def test_columns(values, expected):
    assert columns(values) == expected


def test_xlws_filter(fixture_xls_copy):
    compiler = ExcelCompiler(fixture_xls_copy('filter.xlsx'))
    result = compiler.validate_serialized()
    assert result == {}


@pytest.mark.parametrize(
    'lkup, row_idx, result, approx', (
        ('A', 0, VALUE_ERROR, True),
        ('A', 1, 'A', True),
        ('A', 2, 1, True),
        ('A', 3, 'Z', True),
        ('A', 4, 5, True),
        ('A', 5, REF_ERROR, True),
        ('B', 1, 'B', True),
        ('C', 1, 'C', True),
        ('B', 2, 2, True),
        ('C', 2, 3, True),
        ('B', 3, 'Y', True),
        ('C', 3, 'X', True),
        ('D', 3, 'X', True),
        ('D', 3, NA_ERROR, False),
        ('D', 3, 'X', -1),
        ((('D', 'A'),), 3, ((NA_ERROR, 'Z'), ), False),
    )
)
def test_hlookup(lkup, row_idx, result, approx):
    table = (
        ('A', 'B', 'C'),
        (1, 2, 3),
        ('Z', 'Y', 'X'),
        (5, 6, 7),
    )
    assert result == hlookup(lkup, table, row_idx, approx)


@pytest.mark.parametrize(
    'values, expected', (
        ((1, 1, 1, 1), NA_ERROR),
        ((1, ((1, 2), (3, 4)), 1, 1), 1),
        ((REF_ERROR, ((1, 2), (3, 4)), 1, 1), REF_ERROR),
        ((1, REF_ERROR, 1, 1), NA_ERROR),
        ((1, ((1, 2), (3, 4)), REF_ERROR, 1), REF_ERROR),
        ((1, ((1, 2), (3, 4)), 1, REF_ERROR), REF_ERROR),
        ((1, ((1, 2), (3, 4)), 0, 1), VALUE_ERROR),
        ((1, ((1, 2), (3, 4)), 3, 1), REF_ERROR),
    )
)
def test_hlookup_vlookup_error(values, expected):
    assert hlookup(*values) == expected
    assert vlookup(*values) == expected


class TestIndex:
    """Value of an element in an array selected by the row and column number indexes."""
    test_data_all = ((0, 1), (DIV0, 3))
    test_data_col = ((0,), (DIV0,))
    test_data_row = ((0, 1),)
    test_data_np = np.asarray(((0, 1), (2, 3)))

    test_addr = AddressCell('sh!A1')
    test_data_ref = {
        test_data_all: ((test_addr, test_addr), (test_addr, test_addr)),
        test_data_col: ((test_addr,), (test_addr,)),
        test_data_row: ((test_addr, test_addr),),
    }

    test_data = (
        'data, row_num, col_num, expected', (
            (test_data_all, -1, 0, VALUE_ERROR),
            (test_data_all, 0, -1, VALUE_ERROR),
            (test_data_all, 0, 0, test_data_all),
            (test_data_all, 0, 1, ((0,), (DIV0,))),
            (test_data_all, 0, 2, ((1,), (3,))),
            (test_data_all, 0, 3, REF_ERROR),
            (test_data_all, 1, -1, VALUE_ERROR),
            (test_data_all, 1, 0, ((0, 1),)),
            (test_data_all, 1, 1, 0),
            (test_data_all, 1, 2, 1),
            (test_data_all, 1, 3, REF_ERROR),
            (test_data_all, 2, -1, VALUE_ERROR),
            (test_data_all, 2, 0, ((DIV0, 3),)),
            (test_data_all, 2, 1, DIV0),
            (test_data_all, 2, 2, 3),
            (test_data_all, 2, 3, REF_ERROR),
            (test_data_all, 3, -1, VALUE_ERROR),
            (test_data_all, 3, 0, REF_ERROR),
            (test_data_all, 3, 1, REF_ERROR),
            (test_data_all, None, None, test_data_all),
            (test_data_row, 2, 2, REF_ERROR),
            (test_data_col, 1, 3, REF_ERROR),

            # no col given
            (test_data_all, -1, None, VALUE_ERROR),
            (test_data_all, 0, None, test_data_all),
            (test_data_all, 1, None, ((0, 1),)),
            (test_data_all, 2, None, ((DIV0, 3),)),
            (test_data_all, 3, None, REF_ERROR),
            (test_data_col, -1, None, VALUE_ERROR),
            (test_data_col, 0, None, test_data_col),
            (test_data_col, 1, None, 0),
            (test_data_col, 2, None, DIV0),
            (test_data_col, 3, None, REF_ERROR),
            (test_data_row, -1, None, VALUE_ERROR),
            (test_data_row, 0, None, test_data_row),
            (test_data_row, 1, None, 0),
            (test_data_row, 2, None, 1),
            (test_data_row, 3, None, REF_ERROR),

            # no row given
            (test_data_all, None, -1, VALUE_ERROR),
            (test_data_all, None, 0, test_data_all),
            (test_data_all, None, 1, ((0,), (DIV0,))),
            (test_data_all, None, 2, ((1,), (3,))),
            (test_data_all, None, 3, REF_ERROR),
            (test_data_col, None, -1, VALUE_ERROR),
            (test_data_col, None, 0, test_data_col),
            (test_data_col, None, 1, 0),
            (test_data_col, None, 2, DIV0),
            (test_data_col, None, 3, REF_ERROR),
            (test_data_row, None, -1, VALUE_ERROR),
            (test_data_row, None, 0, test_data_row),
            (test_data_row, None, 1, 0),
            (test_data_row, None, 2, 1),
            (test_data_row, None, 3, REF_ERROR),

            # numpy
            (test_data_np, 1, 1, 0),
            (test_data_np, 1, 2, 1),
            (test_data_np, 2, 1, 2),
            (test_data_np, 2, 2, 3),

            (test_data_np, 1, None, np.array(((0, 1),))),
            (test_data_np, 2, None, np.array(((2, 3),))),

            (test_data_np, None, 1, np.array(((0,), (2,)))),
            (test_data_np, None, 2, np.array(((1,), (3,)))),
        )
    )

    @staticmethod
    @pytest.mark.parametrize(*test_data)
    def test_index(data, row_num, col_num, expected):
        result = index(data, row_num, col_num)
        if isinstance(expected, np.ndarray):
            assert (result == expected).all()
        else:
            assert result == expected

    @pytest.mark.parametrize(*test_data)
    def test_index_reference(self, data, row_num, col_num, expected):
        if not isinstance(data, np.ndarray):
            def lookup_test_data(address):
                address = AddressCell(address)
                return data[address.row - 1][address.col_idx - 1]

            from unittest import mock
            with mock.patch.dict(index.excel_func_meta['name_space'],
                                 {'_C_': lookup_test_data}):
                result = index(self.test_data_ref[data], row_num, col_num)
            if row_num or col_num:
                assert result == expected
            else:
                assert result == self.test_data_ref[data]

    @staticmethod
    def test_index_error_inputs():
        index_f = error_string_wrapper(index, {1, 2})
        assert NA_ERROR == index_f(NA_ERROR, 1)
        assert NA_ERROR == index_f(TestIndex.test_data_all, NA_ERROR, 1)
        assert NA_ERROR == index_f(TestIndex.test_data_all, 1, NA_ERROR)
        assert VALUE_ERROR == index_f((0, 1), 0, 1)
        assert VALUE_ERROR == index_f(None, 1, 1)


@pytest.mark.parametrize(
    'address, expected', (
        ("A1", AddressCell("A1")),
        ("XFD1", AddressCell("XFD1")),
        ("XFE1", REF_ERROR),
        ("A1048576", AddressCell("A1048576")),
        ("A1048577", REF_ERROR),
        ("XFD1048576", AddressCell("XFD1048576")),
        ("XFD1048577", REF_ERROR),
        ("XFE1048576", REF_ERROR),
        ("R1C1", AddressCell("A1")),
        ("ab", REF_ERROR),
    )
)
def test_indirect(address, expected):
    assert indirect(address) == expected
    if is_address(expected):
        with_sheet = expected.create(expected, sheet='S')
        assert indirect(address, None, 'S') == with_sheet

        address = f'S!{address}'
        assert indirect(address) == with_sheet
        assert indirect(address, None, 'S') == with_sheet


lookup_vector = (('b', 'c', 'd'), )
lookup_result = ((1, 2, 3), )
lookup_rows = lookup_vector[0], lookup_result[0]
lookup_columns = tuple(zip(*lookup_rows))


@pytest.mark.parametrize(
    'lookup_value, result1, result2', (
        ('A', NA_ERROR, NA_ERROR),
        ('B', 'b', 1),
        ('C', 'c', 2),
        ('D', 'd', 3),
        ('E', 'd', 3),
        ('1', NA_ERROR, NA_ERROR),
        (1, NA_ERROR, NA_ERROR),
    )
)
def test_lookup(lookup_value, result1, result2):
    assert result1 == lookup(lookup_value, lookup_vector)
    assert result1 == lookup(lookup_value, tuple(zip(*lookup_vector)))
    assert result2 == lookup(lookup_value, lookup_vector, lookup_result)
    assert result2 == lookup(lookup_value, tuple(zip(*lookup_vector)),
                             tuple(zip(*lookup_result)))
    assert result2 == lookup(lookup_value, lookup_rows)
    assert result2 == lookup(lookup_value, lookup_columns)


def test_lookup_error():
    assert lookup(1, 1) == NA_ERROR
    assert lookup(1, ((1,), (2,)), 0) == NA_ERROR
    assert lookup(1, ((1, 2),), ((),)) == NA_ERROR
    assert lookup(1, ((1,), (2,)), ((1, 2), (3, 4))) == NA_ERROR


@pytest.mark.parametrize(
    'lookup_value, lookup_array, match_type, expected', (
        (DIV0, [1, 2, 3], -1, DIV0),
        (0, [1, 3.3, 5], 1, NA_ERROR),
        (1, [1, 3.3, 5], 1, 1),
        (2, [1, 3.3, 5], 1, 1),
        (4, [1, 3.3, 5], 1, 2),
        (5, [1, 3.3, 5], 1, 3),
        (6, [1, 3.3, 5], 1, 3),

        (6, [5, 3.3, 1], -1, NA_ERROR),
        (5, [5, 3.3, 1], -1, 1),
        (4, [5, 3.3, 1], -1, 1),
        (2, [5, 3.3, 1], -1, 2),
        (1, [5, 3.3, 1], -1, 3),
        (0, [5, 3.3, 1], -1, 3),

        (5, [10, 3.3, 5.0], 0, 3),
        (3, [10, 3.3, 5, 2], 0, NA_ERROR),

        (0, [None, None, 1, 3.3, 5, None, None], 1, NA_ERROR),
        (1, [None, None, 1, 3.3, 5, None, None], 1, 3),
        (2, [None, None, 1, 3.3, 5, None, None], 1, 3),
        (4, [None, None, 1, 3.3, 5, None, None], 1, 4),
        (5, [None, None, 1, 3.3, 5, None, None], 1, 5),
        (6, [None, None, 1, 3.3, 5, None, None], 1, 5),

        ('b', ['c', DIV0, 'a'], 0, NA_ERROR),
        ('b', ['c', DIV0, 'a'], -1, 1),

        (False, [True, True, True], 0, NA_ERROR),
        (False, [True, False, True], -1, 2),

        (NA_ERROR, [True, False, True], -1, NA_ERROR),
        (DIV0, [1, 2, 3], -1, DIV0),

        ('Th*t', ['xyzzy', 1, False, DIV0, 'That', 'TheEnd'], 0, 5),
        ('Th*t', ['xyzzy', 1, False, DIV0, 'Tht', 'TheEnd'], 0, 5),
        ('Th*t', ['xyzzy', 1, False, DIV0, 'Tt', 'TheEnd'], 0, NA_ERROR),
        ('Th?t', ['zyzzy', 1, False, DIV0, 'That', 'TheEnd'], 0, 5),
        ('Th?t', ['xyzzy', 1, False, DIV0, 'Tht', 'TheEnd'], 0, NA_ERROR),
        ('Th*t', ['xyzzy', 1, False, DIV0, 'Tat', 'TheEnd'], 0, NA_ERROR),
    )
)
def test_match(lookup_value, lookup_array, match_type, expected):
    lookup_row = (tuple(lookup_array), )
    lookup_col = tuple((i, ) for i in lookup_array)
    assert match(lookup_value, lookup_row, match_type) == expected
    assert match(lookup_value, lookup_col, match_type) == expected


@pytest.mark.parametrize(
    'lookup_array, lookup_value, result1, result0, resultm1', (
        (('a', 'b', 'c', 'd', 'e'), 'c', 3, 3, '#N/A'),  # 0
        (('a', 'b', 'bb', 'd', 'e'), 'c', 3, '#N/A', '#N/A'),  # 1
        (('a', 'b', True, 'd', 'e'), 'c', 2, '#N/A', '#N/A'),  # 2
        (('a', 'b', 1, 'd', 'e'), 'c', 2, '#N/A', '#N/A'),  # 3
        (('a', 'b', '#DIV/0!', 'd', 'e'), 'c', 2, '#N/A', '#N/A'),  # 4
        (('e', 'd', 'c', 'b', 'a'), 'c', 3, 3, 3),  # 5
        (('e', 'd', 'ca', 'b', 'a'), 'c', '#N/A', '#N/A', 3),  # 6
        (('e', 'd', True, 'b', 'a'), 'c', 5, '#N/A', 2),  # 7
        (('e', 'd', 1, 'b', 'a'), 'c', 5, '#N/A', 2),  # 8
        (('e', 'd', '#DIV/0!', 'b', 'a'), 'c', 5, '#N/A', 2),  # 9
        ((5, 4, 3, 2, 1), 3, 3, 3, 3),  # 10
        ((5, 4, 3.5, 2, 1), 3, '#N/A', '#N/A', 3),  # 11
        ((5, 4, True, 2, 1), 3, 5, '#N/A', 2),  # 12
        ((5, 4, 'A', 2, 1), 3, 5, '#N/A', 2),  # 13
        ((5, 4, '#DIV/0!', 2, 1), 3, 5, '#N/A', 2),  # 14
        ((1, 2, 3, 2, 4), 0.5, '#N/A', '#N/A', 5),  # 15
        ((1, 2, 3, 2, 4), 1, 1, 1, 1),  # 16
        ((1, 2, 3, 2, 4), 1.5, 1, '#N/A', '#N/A'),  # 17
        ((1, 2, 3, 2, 4), 2, 2, 2, '#N/A'),  # 18
        ((1, 2, 3, 2, 4), 2.5, 2, '#N/A', '#N/A'),  # 19
        ((1, 2, 3, 2, 4), 3, 3, 3, '#N/A'),  # 20
        ((1, 2, 3, 2, 4), 3.5, 4, '#N/A', '#N/A'),  # 21
        ((1, 2, 3, 2, 4), 4, 5, 5, '#N/A'),  # 22
        ((1, 2, 3, 2, 4), 4.5, 5, '#N/A', '#N/A'),  # 23
        ((4, 3, 2, 3, 1), 4.5, 5, '#N/A', '#N/A'),  # 24
        ((4, 3, 2, 3, 1), 4, 5, 1, 1),  # 25
        ((4, 3, 2, 3, 1), 3.5, 5, '#N/A', 1),  # 26
        ((4, 3, 2, 3, 1), 3, 4, 2, 2),  # 27
        ((4, 3, 2, 3, 1), 2.5, 3, '#N/A', 2),  # 28
        ((4, 3, 2, 3, 1), 2, 3, 3, 3),  # 29
        ((4, 3, 2, 3, 1), 1.5, '#N/A', '#N/A', 4),  # 30
        ((4, 3, 2, 3, 1), 1, '#N/A', 5, 5),  # 31
        ((4, 3, 2, 3, 1), 0.5, '#N/A', '#N/A', 5),  # 32
        (('a', 'b', 'c', 'b', 'd'), '-', '#N/A', '#N/A', 5),  # 33
        (('a', 'b', 'c', 'b', 'd'), 'a', 1, 1, 1),  # 34
        (('a', 'b', 'c', 'b', 'd'), 'aa', 1, '#N/A', '#N/A'),  # 35
        (('a', 'b', 'c', 'b', 'd'), 'b', 2, 2, '#N/A'),  # 36
        (('a', 'b', 'c', 'b', 'd'), 'bb', 2, '#N/A', '#N/A'),  # 37
        (('a', 'b', 'c', 'b', 'd'), 'c', 3, 3, '#N/A'),  # 38
        (('a', 'b', 'c', 'b', 'd'), 'cc', 4, '#N/A', '#N/A'),  # 39
        (('a', 'b', 'c', 'b', 'd'), 'd', 5, 5, '#N/A'),  # 40
        (('a', 'b', 'c', 'b', 'd'), 'dd', 5, '#N/A', '#N/A'),  # 41
        (('d', 'c', 'b', 'c', 'a'), 'dd', 5, '#N/A', '#N/A'),  # 42
        (('d', 'c', 'b', 'c', 'a'), 'd', 5, 1, 1),  # 43
        (('d', 'c', 'b', 'c', 'a'), 'cc', 5, '#N/A', 1),  # 44
        (('d', 'c', 'b', 'c', 'a'), 'c', 4, 2, 2),  # 45
        (('d', 'c', 'b', 'c', 'a'), 'bb', 3, '#N/A', 2),  # 46
        (('d', 'c', 'b', 'c', 'a'), 'b', 3, 3, 3),  # 47
        (('d', 'c', 'b', 'c', 'a'), 'aa', '#N/A', '#N/A', 4),  # 48
        (('d', 'c', 'b', 'c', 'a'), 'a', '#N/A', 5, 5),  # 49
        (('d', 'c', 'b', 'c', 'a'), '-', '#N/A', '#N/A', 5),  # 50

        ((False, False, True), True, 3, 3, NA_ERROR),  # 51
        ((False, False, True), False, 2, 1, 1),  # 52
        ((False, True, False), True, 2, 2, NA_ERROR),  # 53
        ((False, True, False), False, 1, 1, 1),  # 54
        ((True, False, False), True, 3, 1, 1),  # 55
        ((True, False, False), False, 3, 2, 2),  # 56

        (('a', 'AAB', 'rars'), 'rars', 3, 3, NA_ERROR),  # 57
        (('a', 'AAB', 'rars'), 'AAB', 2, 2, NA_ERROR),  # 58
        (('a', 'AAB', 'rars'), 'a', 1, 1, 1),  # 59

        (('AAB', 'a', 'rars'), 'b', 2, NA_ERROR, NA_ERROR),  # 60
        (('AAB', 'a', 'rars'), 3, NA_ERROR, NA_ERROR, NA_ERROR),  # 61
        (('a', 'rars', 'AAB'), 'b', 1, NA_ERROR, NA_ERROR),  # 62

        ((), 'a', NA_ERROR, NA_ERROR, NA_ERROR),  # 63

        (('c', 'b', 'a'), 'a', NA_ERROR, 3, 3),  # 64
        ((1, 2, 3), None, NA_ERROR, NA_ERROR, 3),  # 65

        ((2,), 1, NA_ERROR, NA_ERROR, 1),  # 66
        ((2,), 2, 1, 1, 1),  # 67
        ((2,), 3, 1, NA_ERROR, NA_ERROR),  # 68

        ((3, 5, 4.5, 3, 1), 4, 1, NA_ERROR, NA_ERROR),  # 69
        ((3, 5, 4, 3, 1), 4, 3, 3, NA_ERROR),  # 70
        ((3, 5, 3.5, 3, 1), 4, 5, NA_ERROR, NA_ERROR),  # 71

        ((4, 5, 4.5, 3, 1), 4, 1, 1, 1),  # 72
        ((4, 5, 4, 3, 1), 4, 3, 1, 1),  # 73
        ((4, 5, 3.5, 3, 1), 4, 5, 1, 1),  # 74

        ((1, 3, 3, 3, 5), 3, 4, 2, NA_ERROR),  # 75
        ((5, 3, 3, 3, 1), 3, 4, 2, 2),  # 76
    )
)
def test_match_crazy_order(
        lookup_array, lookup_value, result1, result0, resultm1):
    assert result0 == _match(lookup_value, lookup_array, 0)
    assert resultm1 == _match(lookup_value, lookup_array, -1)
    if result1 != _match(lookup_value, lookup_array, 1):
        lookup_array = [ExcelCmp(x) for x in lookup_array]
        if sorted(lookup_array) == lookup_array:
            # only complain on failures for mode 0 when array is sorted
            assert result1 == _match(lookup_value, lookup_array, 1)


@pytest.mark.parametrize(
    "crwh, refer, rows, cols, height, width", (
        (REF_ERROR, "A1", -1, 0, 1, 1),
        ((1, 1, 1, 1), "A1", 0, 0, 1, 1),
        ((2, 1, 1, 1), "A1", 0, 1, 1, 1),
        ((1, 2, 1, 1), "A1", 1, 0, 1, 1),
        (REF_ERROR, "A1", -1, 0, 1, 1),
        (REF_ERROR, "A1", 0, -1, 1, 1),
        ((16384, 1048576, 1, 1), "XFD1048576", 0, 0, 1, 1),
        ((16384, 1048575, 1, 1), "XFD1048576", -1, 0, 1, 1),
        ((16383, 1048576, 1, 1), "XFD1048576", 0, -1, 1, 1),
        (REF_ERROR, "XFD1048576", 1, 0, 1, 1),
        (REF_ERROR, "XFD1048576", 0, 1, 1, 1),
        (REF_ERROR, "XFD1048576", 0, 0, 2, 1),
        (REF_ERROR, "XFD1048576", 0, 0, 1, 2),
        ((16384, 1048575, 1, 2), "XFD1048576", -1, 0, 2, 1),
        ((16383, 1048576, 2, 1), "XFD1048576", 0, -1, 1, 2),
    )
)
def test_offset(crwh, refer, rows, cols, height, width):
    expected = crwh
    if isinstance(crwh, tuple):
        start = AddressCell((crwh[0], crwh[1], crwh[0], crwh[1]))
        end = AddressCell((crwh[0] + crwh[2] - 1, crwh[1] + crwh[3] - 1,
                           crwh[0] + crwh[2] - 1, crwh[1] + crwh[3] - 1))

        expected = AddressRange.create(f'{start.coordinate}:{end.coordinate}')

    result = offset(refer, rows, cols, height, width)
    assert result == expected

    refer_addr = AddressRange.create(refer)
    if height == refer_addr.size.height:
        height = None
    if width == refer_addr.size.width:
        width = None
    assert offset(refer_addr, rows, cols, height, width) == expected


@pytest.mark.parametrize(
    'address, expected', (
        ('L45', 45),
        ('B:E', None),
        ('4:7', ((4,), (5,), (6,), (7,))),
        ('D1:E1', ((1,), )),
        ('D1:D2', ((1,), (2,))),
        (DIV0, DIV0),
        (NUM_ERROR, NUM_ERROR),
        (VALUE_ERROR, VALUE_ERROR),
    )
)
def test_row(address, expected):
    try:
        address = AddressRange.create(address)
    except ValueError:
        pass

    result = row(address)
    if expected is None:
        assert 1 == next(iter(result))
    else:
        assert expected == result


@pytest.mark.parametrize(
    'values, expected', (
        (((1, None, None), (1, 2, None)), 2),
        (1, 1),
        ("s", 1),
        (((1.2, 3.4), (0.4, 5)), 2),
        (((None, None,), ), 1),
        (((1,), (2,), (3,)), 3)
    )
)
def test_rows(values, expected):
    assert rows(values) == expected


@pytest.mark.parametrize(
    'lkup, col_idx, result, approx', (
        ('A', 0, VALUE_ERROR, True),
        ('A', 1, 'A', True),
        ('A', 2, 1, True),
        ('A', 3, 'Z', True),
        ('A', 4, 5, True),
        ('A', 5, REF_ERROR, True),
        ('B', 1, 'B', True),
        ('C', 1, 'C', True),
        ('B', 2, 2, True),
        ('C', 2, 3, True),
        ('B', 3, 'Y', True),
        ('C', 3, 'X', True),
        ('D', 3, 'X', True),
        ('D', 3, NA_ERROR, False),
        ('D', 3, 'X', -1),
        ((('D', 'A'),), 3, ((NA_ERROR, 'Z'),), False),
    )
)
def test_vlookup(lkup, col_idx, result, approx):
    table = (
        ('A', 1, 'Z', 5),
        ('B', 2, 'Y', 6),
        ('C', 3, 'X', 7),
    )
    assert result == vlookup(lkup, table, col_idx, approx)
