# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import os
import pickle
import threading
from collections import namedtuple

import pytest
from openpyxl.utils import quote_sheetname

from pycel.excelutil import (
    AddressCell,
    AddressMultiAreaRange,
    AddressRange,
    assert_list_like,
    build_operator_operand_fixup,
    coerce_to_number,
    coerce_to_string,
    criteria_parser,
    EMPTY,
    ExcelCmp,
    find_corresponding_index,
    flatten,
    get_linest_degree,
    handle_ifs,
    in_array_formula_context,
    is_address,
    is_number,
    iterative_eval_tracker,
    list_like,
    MAX_COL,
    MAX_ROW,
    NULL_ERROR,
    NUM_ERROR,
    OPERATORS,
    PyCelException,
    range_boundaries,
    split_sheetname,
    structured_reference_boundaries,
    uniqueify,
    unquote_sheetname,
    VALUE_ERROR,
)
from pycel.excelutil import DIV0


def test_address_range():
    a = AddressRange('a1:b2')
    b = AddressRange('A1:B2')
    c = AddressRange(a)

    assert a == b
    assert b == c

    assert b == AddressRange(b)
    assert b == AddressRange.create(b)

    assert AddressRange('sh!a1:b2') == AddressRange(a, sheet='sh')
    assert AddressCell('C13') == AddressCell('R13C3')

    with pytest.raises(ValueError):
        AddressRange(AddressRange('sh!a1:b2'), sheet='sheet')

    a = AddressRange('A:A')
    assert 'A' == a.start.column
    assert 'A' == a.end.column
    assert 0 == a.start.row
    assert 0 == a.end.row

    b = AddressRange('1:1')
    assert '' == b.start.column
    assert '' == b.end.column
    assert 1 == b.start.row
    assert 1 == b.end.row

    # address cell's start and end is himself
    c = b.start
    assert c.start == c.end == c


def test_address_range_errors():

    with pytest.raises(ValueError):
        AddressRange('B32:B')

    with pytest.raises(ValueError):
        AddressRange('B32:B33:B')


@pytest.mark.parametrize(
    'address, expected', (
        ('s!D2:F4:E3', 's!D2:F4'),
        ('s!D2:E3:F4', 's!D2:F4'),
        ('s!E3:D2:F4', 's!D2:F4'),
        ('s!D2:F4:G3', 's!D2:G4'),
        ('s!D2:G3:F4', 's!D2:G4'),
        ('s!G3:D2:F4', 's!D2:G4'),
        ('s!G3:G3:G3', 's!G3'),
    )
)
def test_address_range_multi_colon(address, expected):
    a_range = AddressRange(address)
    assert a_range == AddressRange(expected)


@pytest.mark.parametrize(
    'left, right, result', (
        ('a1', 'a1', 'a1'),
        ('a1:b2', 'b1:c3', 'b1:b2'),
        ('a1:d5', 'b3', 'b3'),
        ('d4:e5', 'c3', NULL_ERROR),
        ('d4:e5', 'd3', NULL_ERROR),
        ('d4:e5', 'e3', NULL_ERROR),
        ('d4:e5', 'f3', NULL_ERROR),
        ('d4:e5', 'c4', NULL_ERROR),
        ('d4:e5', 'd4', 'd4'),
        ('d4:e5', 'e4', 'e4'),
        ('d4:e5', 'f4', NULL_ERROR),
        ('d4:e5', 'c5', NULL_ERROR),
        ('d4:e5', 'd5', 'd5'),
        ('d4:e5', 'e5', 'e5'),
        ('d4:e5', 'f5', NULL_ERROR),
        ('d4:e5', 'c6', NULL_ERROR),
        ('d4:e5', 'd6', NULL_ERROR),
        ('d4:e5', 'e6', NULL_ERROR),
        ('d4:e5', 'f6', NULL_ERROR),
        ('c4:e5', 'd1', NULL_ERROR),
        ('c4:e6', 'a5', NULL_ERROR),
        ('c4:e6', 's!a5', NULL_ERROR),
        ('s!c4:e6', 'a5', NULL_ERROR),
        ('s!c4:e6', 's!a5', NULL_ERROR),
        ('s!c4:e6', 't!a5', VALUE_ERROR),
        ('d4:e5', 's!e5', 's!e5'),
        ('s!d4:e5', 'e5', 's!e5'),
        ('s!d4:e5', 's!e5', 's!e5'),
        ('s!d4:e5', 't!e5', VALUE_ERROR),
    )
)
def test_address_range_and(left, right, result):
    result = AddressRange(result)
    assert AddressRange(left) & AddressRange(right) == result
    assert AddressRange(left) & right == result
    assert left & AddressRange(right) == result


@pytest.mark.parametrize(
    'left, right, result', (
        ('a1', 'a1', 'a1'),
        ('a1:b2', 'b1:c3', 'a1:c3'),
        ('a1:b2', 'd5', 'a1:d5'),
        ('a1:d5', 'b3', 'a1:d5'),
        ('d4:e5', 'a1', 'a1:e5'),
        ('c4:e5', 'd1', 'c1:e5'),
        ('c4:e6', 'a5', 'a4:e6'),
        ('c4:e5', 'd9', 'c4:e9'),
        ('c4:e6', 'j5', 'c4:j6'),
        ('c4:e6', 's!a5', 's!a4:e6'),
        ('s!c4:e6', 'a5', 's!a4:e6'),
        ('s!c4:e6', 's!a5', 's!a4:e6'),
        ('s!c4:e6', 't!a5', VALUE_ERROR),
    )
)
def test_address_range_or(left, right, result):
    result = AddressRange(result)
    assert AddressRange(left) | AddressRange(right) == result
    assert AddressRange(left) | right == result
    assert left | AddressRange(right) == result


@pytest.mark.parametrize(
    'a_range, address, expected', (
        ('s!D2:F4', 's!D2', True),
        ('s!D2:F4', 's!F2', True),
        ('s!D2:F4', 's!D4', True),
        ('s!D2:F4', 's!F4', True),
        ('s!D2:F4', 's!C2', False),
        ('s!D2:F4', 's!D1', False),
        ('s!D2:F4', 's!G4', False),
        ('s!D2:F4', 's!F5', False),
    )
)
def test_address_range_contains(a_range, address, expected):
    a_range = AddressRange(a_range)
    assert expected == (address in a_range)
    address = AddressCell(address)
    assert expected == (address in a_range)
    assert address in address


def test_is_range():

    assert AddressRange('a1:b2').is_range
    assert not AddressRange('a1').is_range


def test_has_sheet():

    assert AddressRange('Sheet1!a1').has_sheet
    assert not AddressRange('a1').has_sheet
    assert AddressRange('Sheet1!a1:b2').has_sheet
    assert not AddressRange('a1:b2').has_sheet

    assert AddressCell('sh!A2') == AddressRange(AddressRange('A2'), sheet='sh')

    with pytest.raises(ValueError, match='Mismatched sheets'):
        AddressRange(AddressRange('shx!a1'), sheet='sh')


def test_address_range_size():

    assert (1, 1) == AddressRange('B1').size
    assert (1, 2) == AddressRange('B1:C1').size
    assert (2, 1) == AddressRange('B1:B2').size
    assert (2, 2) == AddressRange('B1:C2').size

    assert (MAX_ROW, 2) == AddressRange('B:C').size
    assert (3, MAX_COL) == AddressRange('2:4').size


def test_address_cell_addr_inc():

    cell_addr = AddressCell('sh!C2')

    assert MAX_COL - 1 == cell_addr.inc_col(-4)
    assert MAX_COL == cell_addr.inc_col(-3)
    assert 1 == cell_addr.inc_col(-2)
    assert 5 == cell_addr.inc_col(2)
    assert 6 == cell_addr.inc_col(3)

    assert MAX_ROW - 1 == cell_addr.inc_row(-3)
    assert MAX_ROW == cell_addr.inc_row(-2)
    assert 1 == cell_addr.inc_row(-1)
    assert 5 == cell_addr.inc_row(3)
    assert 6 == cell_addr.inc_row(4)


def test_address_cell_addr_offset():

    cell_addr = AddressCell('sh!C2')

    assert AddressCell('sh!XFC1048575') == cell_addr.address_at_offset(-3, -4)
    assert AddressCell('sh!XFD1048576') == cell_addr.address_at_offset(-2, -3)
    assert AddressCell('sh!A1') == cell_addr.address_at_offset(-1, -2)
    assert AddressCell('sh!E5') == cell_addr.address_at_offset(3, 2)
    assert AddressCell('sh!F6') == cell_addr.address_at_offset(4, 3)


def test_address_sort_keys():

    a1_b2 = AddressRange('sh!A1:B2')
    a1 = AddressRange('sh!A1')
    b2 = AddressRange('sh!B2')

    assert a1.sort_key == a1_b2.sort_key
    assert a1.sort_key < b2.sort_key


def test_address_range_columns():
    columns = list(list(x) for x in AddressRange('sh!A1:C3').cols)
    assert 3 == len(columns)
    assert 3 == len(columns[0])

    assert all('A' == addr.column for addr in columns[0])
    assert all('C' == addr.column for addr in columns[-1])


def test_address_pickle(tmpdir):
    addrs = [
        AddressRange('B1'),
        AddressRange('B1:C1'),
        AddressRange('B1:B2'),
        AddressRange('B1:C2'),
        AddressRange('sh!B1'),
        AddressRange('sh!B1:C1'),
        AddressRange('sh!B1:B2'),
        AddressRange('sh!B1:C2'),
        AddressRange('B:C'),
        AddressRange('2:4'),
        AddressCell('sh!XFC1048575'),
        AddressCell('sh!XFD1048576'),
        AddressCell('sh!A1'),
        AddressCell('sh!E5'),
        AddressCell('sh!F6'),
    ]

    filename = os.path.join(str(tmpdir), 'test_addrs.pkl')
    with open(filename, 'wb') as f:
        pickle.dump(addrs, f)

    with open(filename, 'rb') as f:
        new_addrs = pickle.load(f)

    assert addrs == new_addrs


@pytest.mark.parametrize(
    'sheet_name',
    [
        u'In Dusseldorf',
        u'My-Sheet',
        u"Demande d'autorisation",
        "1sheet",
        ".sheet",
        '"',
    ]
)
def test_unquote_sheetname(sheet_name):
    assert sheet_name == unquote_sheetname(quote_sheetname(sheet_name))


@pytest.mark.parametrize(
    'sheet_name',
    [
        u'In Dusseldorf',
        u'My-Sheet',
        u"Demande d'autorisation",
        "1sheet",
        ".sheet",
        '"',
    ]
)
def test_quoted_address(sheet_name):
    addr = AddressCell('A2', sheet=sheet_name)
    assert addr.quoted_address == '{}!A2'.format(addr.quote_sheet(sheet_name))


@pytest.mark.parametrize(
    'address, expected', (
        ('s!D2', 's!$D$2'),
        ('s!D2:F4', 's!$D$2:$F$4'),
        (AddressRange("D2:F4", sheet='sh 1'), "'sh 1'!$D$2:$F$4"),
    )
)
def test_address_absolute(address, expected):
    assert AddressRange.create(address).abs_address == expected


def test_split_sheetname():

    assert ('', 'B1') == split_sheetname('B1')
    assert ('sheet', 'B1') == split_sheetname('sheet!B1')
    assert ('', 'B1:C2') == split_sheetname('B1:C2')
    assert ('sheet', 'B1:C2') == split_sheetname('sheet!B1:C2')

    assert ("shee't", 'B1:C2') == split_sheetname("'shee''t'!B1:C2")
    assert ("shee t", 'B1:C2') == split_sheetname("'shee t'!B1:C2")

    with pytest.raises(ValueError):
        split_sheetname('sh!B1', sheet='shx')

    with pytest.raises(NotImplementedError):
        split_sheetname('sh!B1:C2:sh2!B1:C2')


def test_address_cell_enum(ATestCell):
    assert ('B1', '', 2, 1, 'B1') == AddressCell('B1')
    assert ('sheet!B1', 'sheet', 2, 1, 'B1') == AddressCell('sheet!B1')

    assert ('A1', '', 1, 1, 'A1') == AddressCell('R1C1')
    assert ('sheet!A1', 'sheet', 1, 1, 'A1') == AddressCell('sheet!R1C1')

    cell = ATestCell('A', 1)
    assert ('B2', '', 2, 2, 'B2') == AddressCell.create(
        'R[1]C[1]', cell=cell)
    assert ('sheet!B2', 'sheet', 2, 2, 'B2') == AddressCell.create(
        'sheet!R[1]C[1]', cell=cell)

    with pytest.raises(ValueError):
        AddressCell('B1:C2')

    with pytest.raises(ValueError):
        AddressCell('sheet!B1:C2')

    with pytest.raises(ValueError):
        AddressCell('xyzzy')


def test_resolve_range():
    a = AddressRange.create

    assert ((a('B1'), ), ) == a('B1').resolve_range
    assert ((a('B1'), a('C1')),) == a('B1:C1').resolve_range
    assert ((a('B1'),), (a('B2'), )) == a('B1:B2').resolve_range
    assert ((a('B1'), a('C1')), (a('B2'), a('C2'))) == a('B1:C2').resolve_range

    assert ((a('sh!B1'),),) == a('sh!B1').resolve_range
    assert ((a('sh!B1'), a('sh!C1')),) == a('sh!B1:C1').resolve_range
    assert ((a('sh!B1'),), (a('sh!B2'),)) == a('sh!B1:B2').resolve_range
    assert ((a('sh!B1'), a('sh!C1')),
            (a('sh!B2'), a('sh!C2'))) == (a('sh!B1:C2')).resolve_range

    assert ((a('sh!B1'),),) == a('sh!B1', sheet='sh').resolve_range
    assert ((a('sh!B1'), a('sh!C1')),) == (
        a('sh!B1:C1', sheet='sh')).resolve_range
    assert ((a('sh!B1'),), (a('sh!B2'),)) == (
        a('sh!B1:B2', sheet='sh')).resolve_range
    assert ((a('sh!B1'), a('sh!C1')), (a('sh!B2'), a('sh!C2'))) == \
        (a('sh!B1:C2', sheet='sh')).resolve_range

    with pytest.raises(AssertionError):
        a('B:C').resolve_range

    with pytest.raises(AssertionError):
        a('1:2').resolve_range


addr_cr = AddressRange.create


@pytest.mark.parametrize(
    'address, string, mar', (
        (((addr_cr('B1'), ), (addr_cr('B1'), addr_cr('C1'),)),
         'B1,B1:C1',
         AddressMultiAreaRange((addr_cr('B1'), addr_cr('B1:C1')))),
        (((addr_cr('B1'),), (addr_cr('B2'),),
          (addr_cr('B1'), addr_cr('C1')), (addr_cr('B2'), addr_cr('C2'))),
         'B1:B2,B1:C2',
         AddressMultiAreaRange((addr_cr('B1:B2'), addr_cr('B1:C2')))),
    )
)
def test_multi_area_range(address, string, mar):
    assert address == tuple(mar.resolve_range)
    assert not mar.is_unbounded_range
    assert address[0][0] in mar
    assert AddressRange('Z99') not in mar
    assert str(mar) == string


@pytest.mark.parametrize(
    'ref, expected', (
        # valid addresses
        ('a_table[[#This Row], [col5]]', 'E5'),
        ('a_table[[#All],[col3]]', 'C1:C8'),
        ('a_table[[#All],[col3]:[col4]]', 'C1:D8'),
        ('a_table[[#Headers],[col4]]', 'D1'),
        ('a_table[[#Headers],[col2]:[col5]]', 'B1:E1'),

        # Not Supported
        ('a_table[[#Headers],[#Data],[col4]]', PyCelException('D1:D7')),

        ('a_table[[#Data],[col4]:[col4]]', 'D2:D7'),
        ('a_table[[#Data],[col4]:[col5]]', 'D2:E7'),
        ('a_table[[#Totals],[col2]]', 'B8'),
        ('a_table[[#Totals],[col3]:[col5]]', 'C8:E8'),
        ('a_table[[#This Row], [col5]]', 'E5'),
        ('a_table[[col4]:[col4]]', 'D2:D7'),
        ('a_table[@col5]', 'E5'),
        ('a_table[@[col2]]', 'B5'),
        ('a_table[#This Row]', 'A5:E5'),
        ('a_table[@]', 'A5:E5'),
        ('a_table[]', 'A2:E7'),

        # bad table / cell
        ('JUNK[]', PyCelException()),
        ('a_table[]', None),

        # unknown rows or columns
        ('a_table[[#JUNK]]', PyCelException()),
        ('a_table[[#Data],[JUNK]]', PyCelException()),
        ('a_table[[#Data],[JUNK]:[col4]]', PyCelException()),

        # misordered columns
        ('a_table[[#Data],[col5]:[col4]]', PyCelException()),

        # malformed
        ('a_table[[]', PyCelException()),
        ('a_table[[[col4]:[col4]]', PyCelException()),
    )
)
def test_structured_table_reference_boundaries(ref, expected):

    Column = namedtuple('Column', 'name')

    class Table:
        def __init__(self, ref, header_rows, totals_rows):
            self.ref = ref
            self.headerRowCount = header_rows
            self.totalsRowCount = totals_rows
            self.tableColumns = tuple(
                Column(name) for name in 'col1 col2 col3 col4 col5'.split())

    class Excel:
        def __init__(self, table):
            self.a_table = table

        def table(self, name):
            if name == 'a_table':
                return self.a_table, None
            else:
                return None, None

    class Cell:
        def __init__(self, table, address):
            self.excel = Excel(table)
            self.address = AddressCell(address)

    cell = Cell(Table('A1:E8', 1, 1), 'E5')

    if isinstance(expected, PyCelException):
        with pytest.raises(PyCelException):
            structured_reference_boundaries(ref, cell=cell)

    elif expected is None:
        with pytest.raises(PyCelException):
            structured_reference_boundaries(ref, cell=None)

    else:
        ref_bound = structured_reference_boundaries(ref, cell=cell)
        expected_bound = range_boundaries(expected, cell=cell)
        assert ref_bound == expected_bound

        expected_ref = range_boundaries(ref, cell=cell)
        assert ref_bound == expected_ref


@pytest.mark.parametrize(
    'expected, address', (

        ((1, 2) * 2, 'A2'),
        ((2, 1) * 2, 'B1'),
        ((1, 2) * 2, 'R2C1'),
        ((2, 1) * 2, 'R1C2'),
        ((2, 3) * 2, 'R[2]C[1]'),
        ((3, 2) * 2, 'R[1]C[2]'),

        ((1, 1, 2, 2), 'A1:B2'),
        ((1, 1, 2, 2), 'R1C1:R2C2'),
        ((2, 1, 2, 3), 'R1C2:R[2]C[1]'),

        ((3, 13) * 2, 'R13C3'),

        ((1, 1, 1, 1), 'RC'),

        ((None, 1, None, 4), 'R:R[3]'),
        ((None, 1, None, 4), 'R1:R[3]'),
        ((None, 2, None, 4), 'R2:R[3]'),

        ((1, None, 4, None), 'C:C[3]'),
        ((1, None, 4, None), 'C1:C[3]'),
        ((2, None, 4, None), 'C2:C[3]'),

        ((4, 2, 6, 4), 's!D2:F4:E3'),
        ((4, 2, 6, 4), 's!D2:E3:F4'),
        ((4, 2, 6, 4), 's!E3:D2:F4'),
        ((4, 2, 7, 4), 's!D2:F4:G3'),
        ((4, 2, 7, 4), 's!D2:G3:F4'),
        ((4, 2, 7, 4), 's!G3:D2:F4'),
        ((7, 3, 7, 3), 's!G3:G3:G3'),
    )
)
def test_extended_range_boundaries(expected, address, ATestCell):
    assert range_boundaries(address, cell=ATestCell('A', 1))[0] == expected


def test_range_boundaries_defined_names(excel, ATestCell):
    cell = ATestCell('A', 1, excel=excel)

    assert ((3, 1, 3, 18), 'Sheet1') == range_boundaries('SINUS', cell)
    assert ((2, 1, 5, 18), 'Sheet1') == range_boundaries('B2:E5:SINUS', cell)


@pytest.mark.parametrize(
    'address_string',
    [
        'R',
        'C',
        ':',
        'R:',
        'C:',
        ':R',
        ':C',
        'RC:',
        ':RC',
        'R:C1',
        'C:R1',
        'C1:RC',
        'R1:RC',
        'RC:R1',
        'RC:C1',
        'sheet!B1',
        'xyzzy',
    ]
)
def test_extended_range_boundaries_errors(address_string, ATestCell):
    cell = ATestCell('A', 1)

    with pytest.raises(ValueError, match='not a valid coordinate or range'):
        range_boundaries(address_string, cell)


def test_multi_area_ranges(excel, ATestCell):
    cell = ATestCell('A', 1, excel=excel)
    from unittest import mock
    with mock.patch.object(excel, '_defined_names', {
            'dname': (('$A$1', 's1'), ('$A$3:$A$4', 's2'))}):

        multi_area_range = AddressMultiAreaRange(
            tuple(AddressRange(addr, sheet=sh))
            for addr, sh in excel._defined_names['dname'])

        assert (multi_area_range, None) == range_boundaries('dname', cell)
        assert multi_area_range == AddressRange.create('dname', cell=cell)


@pytest.mark.parametrize(
    'value, expected, expected_type, convert_all', (
        (1, 1, int, False),
        (1.0, 1.0, int, False),
        (None, None, type(None), False),
        ('1', 1, int, False),
        ('1.', 1.0, float, False),
        ('xyzzy', 'xyzzy', str, False),
        (DIV0, DIV0, str, False),
        ('TRUE', 'TRUE', str, False),
        ('FALSE', 'FALSE', str, False),
        (EMPTY, EMPTY, str, False),
        ((('TRUE',), ), 'TRUE', str, False),
        (1, 1, int, True),
        (1.0, 1.0, int, True),
        (None, 0, int, True),
        ('1', 1, int, True),
        ('1.', 1.0, float, True),
        ('xyzzy', 'xyzzy', str, True),
        (DIV0, DIV0, str, True),
        ('TRUE', 1, int, True),
        ('FALSE', 0, int, True),
        (EMPTY, 0, int, True),
        ((('TRUE',), ), 1, int, True),
    )
)
def test_coerce_to_number(value, expected, expected_type, convert_all):
    result = coerce_to_number(value, convert_all=convert_all)
    assert result == expected
    assert isinstance(result, expected_type)


@pytest.mark.parametrize(
    'value, result', (
        (True, 'TRUE'),
        (False, 'FALSE'),
        (None, ''),
        (1, '1'),
        (1.0, '1'),
        (1.1, '1.1'),
        ('xyzzy', 'xyzzy'),
    )
)
def test_coerce_to_string(value, result):
    assert coerce_to_string(value) == result


def test_get_linest_degree():
    # build a spreadsheet with linest formulas horiz and vert

    class Excel:

        def __init__(self, columns, rows):
            self.columns = columns
            self.rows = rows

        def get_formula_from_range(self, address):
            addr = AddressRange.create(address)
            found = addr.column in self.columns and str(addr.row) in self.rows
            return '=linest()' if found else ''

    class Cell:
        def __init__(self, excel):
            self.excel = excel
            self.address = AddressCell('E5')

        @property
        def sheet(self):
            return 'PhonySheet'

        @property
        def formula(self):
            return '=linest()'

    assert (1, 1) == get_linest_degree(Cell(Excel('E', '5')))

    assert (4, 5) == get_linest_degree(Cell(Excel('E', '12345')))
    assert (4, 4) == get_linest_degree(Cell(Excel('E', '23456')))
    assert (4, 3) == get_linest_degree(Cell(Excel('E', '34567')))
    assert (4, 2) == get_linest_degree(Cell(Excel('E', '45678')))
    assert (4, 1) == get_linest_degree(Cell(Excel('E', '56789')))

    assert (4, 5) == get_linest_degree(Cell(Excel('ABCDE', '5')))
    assert (4, 4) == get_linest_degree(Cell(Excel('BCDEF', '5')))
    assert (4, 3) == get_linest_degree(Cell(Excel('CDEFG', '5')))
    assert (4, 2) == get_linest_degree(Cell(Excel('DEFGH', '5')))
    assert (4, 1) == get_linest_degree(Cell(Excel('EFGHI', '5')))


def test_in_array_formula_context():

    assert not in_array_formula_context
    with in_array_formula_context('A1'):
        assert in_array_formula_context

    def return_in_context():
        return bool(in_array_formula_context)

    assert not return_in_context()
    with in_array_formula_context('A1'):
        assert return_in_context()

    assert not return_in_context()
    try:
        with in_array_formula_context('A1'):
            assert return_in_context()
            raise PyCelException
    except PyCelException:
        pass
    assert not return_in_context()

    class AThread(threading.Thread):
        def run(self):
            try:
                in_ctx_1 = return_in_context()
                with in_array_formula_context('A1'):
                    in_ctx_2 = return_in_context()
                self.result = not in_ctx_1 and in_ctx_2
            except:  # noqa: E722
                self.result = False

    thread = AThread()
    thread.start()
    thread.join()
    assert thread.result


@pytest.mark.parametrize(
    'address, value, result', (
        ('A1:A2', 3, ((3, ), (3, ))),
        (None, 1, 1),
        (None, ((1, 2), (3, 4)), ((1, 2), (3, 4))),
        ('A1', 1, ((1,),)),
        ('A1', ((1, 2), (3, 4)), ((1,),)),

        ('A1:B1', 2, ((2, 2),)),
        ('A1:A2', 3, ((3, ), (3, ))),
        ('A1:B2', 4, ((4, 4), (4, 4),)),

        ('A1:B1', ((1, 2),), ((1, 2),)),
        ('A1:B2', ((1, 2),), ((1, 2), (1, 2),)),

        ('A1:A2', ((1, ), (3, )), ((1, ), (3, ))),
        ('A1:B2', ((1, ), (3, )), ((1, 1), (3, 3))),

        ('A1:B1', ((1, 2), (3, 4)), ((1, 2),)),
        ('A1:A2', ((1, 2), (3, 4)), ((1, ), (3, ),)),

        ('A1:C3', ((1, 2), (3, 4)),
         ((1, 2, '#N/A'), (3, 4, '#N/A'), ('#N/A', '#N/A', '#N/A'))),
    )
)
def test_array_formula_context_fit_to_range(address, value, result):
    if address is not None:
        address = AddressRange(address, sheet='s')
    with in_array_formula_context(address):
        assert in_array_formula_context.fit_to_range(value) == result


def test_flatten():
    assert ['ddd'] == list(flatten(['ddd']))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', 1, (2, 3)]))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', (1, (2, 3))]))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', (1, 2), 3]))

    assert [None] == list(flatten(None))
    assert [True] == list(flatten(True))
    assert [1.0] == list(flatten(1.0))


def test_uniqueify():
    assert (1, 2, 3, 4) == uniqueify((1, 2, 3, 4, 3))
    assert (4, 1, 2, 3) == uniqueify((4, 1, 2, 3, 4, 3))


@pytest.mark.parametrize(
    'data, result', (
        (AddressCell('A1'), True),
        (AddressRange('A1:B2'), True),
        ('A1', False),
        ('A1:B2', False),

        (1, False),
        (0, False),
        (-1, False),
        (1.0, False),
        ('-1.0', False),
        (True, False),
        (False, False),
        (None, False),
        ('x', False),
    )
)
def test_is_address(data, result):
    assert is_address(data) == result


@pytest.mark.parametrize(
    'data, result', (
        (1, True),
        (0, True),
        (-1, True),
        (1.0, True),
        (0.0, True),
        (-1.0, True),
        ('1.0', True),
        ('0.0', True),
        ('-1.0', True),
        (True, True),
        (False, True),

        (None, False),
        ('False', False),
        ('x', False),
        (AddressCell('A1'), False),
    )
)
def test_is_number(data, result):
    assert is_number(data) == result


@pytest.mark.parametrize(
    'data, result', (
        ((12, 12), ((0, 0), )),
        ((12, 12, 12), AssertionError),
        ((((1, 1, 2, 2, 2), ), 2), ((0, 2), (0, 3), (0, 4))),
        ((((1, 2, 3, 4, 5), ), ">=3"), ((0, 2), (0, 3), (0, 4))),
        ((((1, 2, 3, 4, 5), ), ">=3"), ((0, 2), (0, 3), (0, 4))),
        ((((1, 2), (3, 4)), ">=3"), ((1, 0), (1, 1))),
        ((((1, 2, 3, 4, 5), ), ">=3"), ((0, 2), (0, 3), (0, 4))),
        (('JUNK', ((), ), ((), ), ), AssertionError),
        ((((1,),), '', ((1, 2),), ''), VALUE_ERROR),
        ((((1, 2, 3, 4, 5), ), ">=3",
          ((1, 2, 3, 4, 5), ), "<=4"), ((0, 2), (0, 3))),
    )
)
def test_handle_ifs(data, result):
    if isinstance(result, type(Exception)):
        with pytest.raises(result):
            handle_ifs(data)
    elif isinstance(result, str):
        assert handle_ifs(data) == result
    else:
        assert tuple(sorted(handle_ifs(data))) == result


def test_handle_ifs_op_range():
    with pytest.raises(TypeError):
        handle_ifs(((1, ), (1, )), 2)

    assert handle_ifs((((1, 2), (3, 4)), ">=3"), ((1, ), (1, ))) == VALUE_ERROR

    assert handle_ifs((((1,), ), "=1"), 1) == ((0, 0), )


def test_find_corresponding_index():
    assert ((0, 0), ) == find_corresponding_index(((1, 2, 3), ), '<2')
    assert ((0, 2),) == find_corresponding_index(((1, 2, 3), ), '>2')
    assert ((0, 0), (0, 2)) == find_corresponding_index(((1, 2, 3), ), '<>2')
    assert ((0, 0), (0, 1)) == find_corresponding_index(((1, 2, 3), ), '<=2')
    assert ((0, 1), (0, 2)) == find_corresponding_index(((1, 2, 3), ), '>=2')
    assert ((0, 1),) == find_corresponding_index(((1, 2, 3), ), '2')
    assert ((0, 1),) == find_corresponding_index((list('ABC'), ), 'B')
    assert ((0, 1), (0, 2)) == find_corresponding_index((list('ABB'), ), 'B')
    assert ((0, 1), (0, 2)) == find_corresponding_index((list('ABB'), ), '<>A')
    assert () == find_corresponding_index((list('ABB'), ), 'D')

    with pytest.raises(TypeError):
        find_corresponding_index('ABB', '<B')

    with pytest.raises(ValueError):
        find_corresponding_index((list('ABB'), ), None)


@pytest.mark.parametrize(
    'value, expected', (
        ('xyzzy', False),
        (AddressRange('A1:B2'), False),
        (AddressCell('A1'), False),
        ([1, 2], True),
        ((1, 2), True),
        ({1: 2, 3: 4}, True),
        ((a for a in range(2)), True),
    )
)
def test_list_like(value, expected):
    assert list_like(value) == expected
    if expected:
        assert_list_like(value)
    else:
        with pytest.raises(TypeError, match='Must be a list like: '):
            assert_list_like(value)


@pytest.mark.parametrize(
    'value, criteria, expected', (
        (0, 1, False),
        (1, 1, True),
        (2, 1, False),
        ('0', 1, False),
        ('1', 1, True),
        ('2', 1, False),

        (0, '1', False),
        (1, '1', True),
        (2, '1', False),
        ('0', '1', False),
        ('1', '1', True),
        ('2', '1', False),

        (0, '=1', False),
        (1, '=1', True),
        (2, '=1', False),
        ('0', '=1', False),
        ('1', '=1', True),
        ('2', '=1', False),

        (0, '<>1', True),
        (1, '<>1', False),
        (2, '<>1', True),
        ('0', '<>1', True),
        ('1', '<>1', True),
        ('2', '<>1', True),

        (0, '>1', False),
        (1, '>1', False),
        (2, '>1', True),
        ('0', '>1', False),
        ('1', '>1', False),
        ('2', '>1', False),

        (0, '<1', True),
        (1, '<1', False),
        (2, '<1', False),
        ('0', '<1', False),
        ('1', '<1', False),
        ('2', '<1', False),

        (0, '>1x', False),
        (1, '>1x', False),
        (2, '>1x', False),
        ('0', '>1x', False),
        ('1', '>1x', False),
        ('2', '>1x', True),

        ('a', 'b', False),
        ('b', 'b', True),
        ('c', 'b', False),
        ('a', '=b', False),
        ('b', '=b', True),
        ('c', '=b', False),

        ('a', '<>b', True),
        ('b', '<>b', False),
        ('c', '<>b', True),

        ('a', '<b', True),
        ('b', '<b', False),
        ('c', '<b', False),
        ('a', '<=b', True),
        ('b', '<=b', True),
        ('c', '<=b', False),

        ('a', '<0', False),
        ('b', '<1', False),
        ('c', '>=1', False),

        ('a', '<0x', False),
        ('b', '<1x', False),
        ('c', '>=1x', True),

        ('a', '<=B', True),
        ('b', '<=B', True),
        ('c', '<=B', False),
        ('a', 'B', False),
        ('b', 'B', True),
        ('c', 'B', False),

        ('1x', '1x', True),
        ('1x', '=1x', True),
        ('1x', '>1x', False),
        ('1x', '>=1x', True),
        ('1x', '<1x', False),
        ('1x', '<=1x', True),
        ('1x', '<>1x', False),

        ('That', 'Th?t', True),
        ('That', 'T*t', True),
        ('Tt', 'T*t', True),
        ('Tht', 'Th?t', False),
        ('Tat', 'Th*t', False),
        (None, 'Th?t', False),
        (None, 'Th*t', False),

        ('', '', True),
        (None, '', True),
        (1, '', False),
        ('1', '', False),
        ('1x', '', False),
        ('a', '', False),

        (None, '', True),
        (None, 1, False),
        (None, '1', False),
        (None, '=1', False),
        (None, '<>1', True),
        (None, '>1', False),
        (None, '<1', False),
        (None, '>1x', False),
        (None, 'b', False),
        (None, '=b', False),
        (None, '<>b', True),
        (None, '<b', False),
        (None, '<=b', False),
        (None, '=', True),
        (None, '<>', False),
    )
)
def test_criteria_parser(value, criteria, expected):
    assert expected == criteria_parser(criteria)(value)


@pytest.mark.parametrize(
    'lval, op, rval, result', (
        (1, '>', 1, False),
        (1, '>=', 1, True),
        (1, '<', 1, False),
        (1, '<=', 1, True),
        (1, '=', 1, True),
        (1, '<>', 1, False),

        (1, '>', 2, False),
        (1, '>=', 2, False),
        (1, '<', 2, True),
        (1, '<=', 2, True),
        (1, '=', 2, False),
        (1, '<>', 2, True),

        (2, '>', 1, True),
        (2, '>=', 1, True),
        (2, '<', 1, False),
        (2, '<=', 1, False),
        (2, '=', 1, False),
        (2, '<>', 1, True),

        ('a', '>', 'a', False),
        ('a', '>=', 'a', True),
        ('a', '<', 'a', False),
        ('a', '<=', 'a', True),
        ('a', '=', 'A', True),
        ('a', '<>', 'a', False),

        ('a', '>', 'b', False),
        ('a', '>=', 'b', False),
        ('a', '<', 'b', True),
        ('a', '<=', 'b', True),
        ('a', '=', 'B', False),
        ('a', '<>', 'b', True),

        ('b', '>', 'a', True),
        ('b', '>=', 'a', True),
        ('b', '<', 'a', False),
        ('b', '<=', 'a', False),
        ('b', '=', 'A', False),
        ('b', '<>', 'a', True),

        (True, '<', DIV0, True),
        (True, '=', DIV0, False),
        (False, '<', True, True),
        (False, '=', True, False),
        ('z', '<', False, True),
        ('z', '=', False, False),
        ('a', '<', 'z', True),
        ('a', '=', 'z', False),
        (1E10, '<', 'a', True),
        (1E10, '=', 'a', False),
        (0, '<', 1E10, True),
        (0, '=', 1E10, False),
        (-1E10, '<', 0, True),
        (-1E10, '=', 0, False),

        (None, '=', 0, True),
        (None, '<>', 0, False),
        (0, '=', None, True),
        (0, '<>', None, False),

        (None, '=', 0.0, True),
        (None, '<>', 0.0, False),
        (0.0, '=', None, True),
        (0.0, '<>', None, False),

        (False, '=', None, True),
        (False, '<>', None, False),
        ('', '=', None, True),
        ('', '<>', None, False),
    )
)
def test_excel_cmp(lval, op, rval, result):
    assert OPERATORS[op](ExcelCmp(lval), rval) == result


@pytest.mark.parametrize(
    'left_op, op, right_op, expected',
    [
        # left None
        (None, 'Eq', '', True),
        (None, 'Eq', '0', False),
        (None, 'Eq', 0, True),
        (None, 'Eq', 1, False),
        (None, 'Eq', False, True),
        (None, 'Eq', True, False),

        # right None
        ('', 'Eq', None, True),
        ('0', 'Eq', None, False),
        (0, 'Eq', None, True),
        (1, 'Eq', None, False),
        (False, 'Eq', None, True),
        (True, 'Eq', None, False),

        # case in-sensitive
        ('a', 'Eq', 'A', True),
        ('A', 'NotEq', 'a', False),
        ('b', 'NotEq', 'A', True),
        ('A', 'Eq', 'b', False),

        # string concat
        ('0', 'BitAnd', 0, '00'),
        (0, 'BitAnd', '0', '00'),
        ('1', 'BitAnd', 1, '11'),
        (1, 'BitAnd', '1', '11'),
        (0, 'BitAnd', 'X', '0X'),
        ('X', 'BitAnd', 0, 'X0'),
        ('X', 'BitAnd', 5.0, 'X5'),
        ('X', 'BitAnd', 5.0, 'X5'),

        # divsion by zero
        (DIV0, '', '', DIV0),
        ('', '', DIV0, DIV0),

        ('1', 'Div', '0', DIV0),
        ('1', 'Div', 0, DIV0),
        (1, 'Div', '0', DIV0),
        (1, 'Div', 0, DIV0),

        (1, 'Mod', '0', DIV0),
        (1, 'Mod', 0, DIV0),

        # type coercion
        (1, 'Add', 2, 3),
        (1, 'Add', '2', 3),
        ('1', 'Add', 2, 3),
        ('1', 'Add', '2', 3),

        (None, 'Add', 2, 2),
        (2, 'Add', None, 2),
        (None, 'Add', '2', 2),
        ('2', 'Add', None, 2),

        (1, 'Sub', 2, -1),
        (1, 'Sub', '2', -1),
        ('1', 'Sub', 2, -1),
        ('1', 'Sub', '2', -1),

        (1, 'Mult', 2, 2),
        (1, 'Mult', '2', 2),
        ('1', 'Mult', 2, 2),
        ('1', 'Mult', '2', 2),

        (1, 'Div', 2, 0.5),
        (1, 'Div', '2', 0.5),
        ('1', 'Div', 2, 0.5),
        ('1', 'Div', '2', 0.5),

        (5, 'Mod', 2, 1),
        (5, 'Mod', '2', 1),
        ('5', 'Mod', 2, 1),
        ('5', 'Mod', '2', 1),

        (2, 'Pow', 2, 4),
        (2, 'Pow', '2', 4),
        ('2', 'Pow', 2, 4),
        ('2', 'Pow', '2', 4),

        ('', 'USub', 2, -2),
        ('', 'USub', '2', -2),
        ('', 'USub', 'X', VALUE_ERROR),
        (None, 'USub', 'X', VALUE_ERROR),
        ('', 'USub', None, 0),

        (5, 'Eq', 5, True),
        (5, 'Eq', 2, False),
        (5, 'Eq', True, False),
        (5, 'Eq', '5', False),
        ('5', 'Eq', '5', True),
        ('5', 'Eq', '2', False),
        ('5', 'Eq', True, False),
        (True, 'Eq', True, True),
        (True, 'Eq', False, False),
        (False, 'Eq', False, True),

        (5, 'Lt', 5, False),
        (5, 'Lt', 2, False),
        (5, 'Lt', True, True),
        (5, 'Lt', '5', True),
        ('5', 'Lt', '5', False),
        ('5', 'Lt', '2', False),
        (True, 'Lt', True, False),
        (True, 'Lt', False, False),
        (False, 'Lt', False, False),

        (True, 'Add', 5, 6),
        (False, 'Add', 5, 5),
        (True, 'Mult', 5, 5),
        (False, 'Mult', 5, 0),
        (5, 'Add', True, 6),
        (5, 'Add', False, 5),

        (True, 'BitAnd', 'xyzzy', 'TRUExyzzy'),
        (False, 'BitAnd', 'xyzzy', 'FALSExyzzy'),
        ('xyzzy', 'BitAnd', True, 'xyzzyTRUE'),
        (True, 'BitAnd', True, 'TRUETRUE'),

        (True, 'BitAnd', 5, 'TRUE5'),
        (False, 'BitAnd', 5, 'FALSE5'),
        (5, 'BitAnd', True, '5TRUE'),
        (5, 'BitAnd', False, '5FALSE'),

        (None, 'BitAnd', False, 'FALSE'),
        (None, 'BitAnd', 5, '5'),
        (None, 'BitAnd', 'xyzzy', 'xyzzy'),
        (False, 'BitAnd', None, 'FALSE'),
        (5, 'BitAnd', None, '5'),
        ('xyzzy', 'BitAnd', None, 'xyzzy'),

        # value errors
        (VALUE_ERROR, 'Add', 0, VALUE_ERROR),
        (0, 'Add', VALUE_ERROR, VALUE_ERROR),
        ('X', 'Add', 0, VALUE_ERROR),
        (0, 'Add', 'X', VALUE_ERROR),
        ('X', 'Add', 'X', VALUE_ERROR),
        (True, 'Add', 'X', VALUE_ERROR),
        (None, 'Add', 'X', VALUE_ERROR),
        ('X', 'Sub', 0, VALUE_ERROR),
        (0, 'Sub', 'X', VALUE_ERROR),
        ('X', 'Sub', 'X', VALUE_ERROR),
        (True, 'Sub', 'X', VALUE_ERROR),
        (None, 'Sub', 'X', VALUE_ERROR),
        ('X', 'Mult', 0, VALUE_ERROR),
        (0, 'Mult', 'X', VALUE_ERROR),
        ('X', 'Mult', 'X', VALUE_ERROR),
        (True, 'Mult', 'X', VALUE_ERROR),
        (None, 'Mult', 'X', VALUE_ERROR),
        ('X', 'Div', 0, VALUE_ERROR),
        (0, 'Div', 'X', VALUE_ERROR),
        ('X', 'Div', 'X', VALUE_ERROR),
        (True, 'Div', 'X', VALUE_ERROR),
        (None, 'Div', 'X', VALUE_ERROR),
        ('X', 'Mod', 0, VALUE_ERROR),
        (0, 'Mod', 'X', VALUE_ERROR),
        ('X', 'Mod', 'X', VALUE_ERROR),
        (True, 'Mod', 'X', VALUE_ERROR),
        (None, 'Mod', 'X', VALUE_ERROR),
        ('X', 'Pow', 0, VALUE_ERROR),
        (0, 'Pow', 'X', VALUE_ERROR),
        ('X', 'Pow', 'X', VALUE_ERROR),
        (True, 'Pow', 'X', VALUE_ERROR),
        (None, 'Pow', 'X', VALUE_ERROR),

        # mixed errors
        (VALUE_ERROR, 'Add', DIV0, VALUE_ERROR),
        (DIV0, 'Add', VALUE_ERROR, DIV0),
        (NUM_ERROR, 'Add', DIV0, NUM_ERROR),
        (DIV0, 'Add', NUM_ERROR, DIV0),
        (NUM_ERROR, 'Add', VALUE_ERROR, NUM_ERROR),
        (VALUE_ERROR, 'Add', NUM_ERROR, VALUE_ERROR),

        # right op errors
        (0, 'Add', DIV0, DIV0),
        (0, 'Sub', VALUE_ERROR, VALUE_ERROR),
        (0, 'Div', NUM_ERROR, NUM_ERROR),

        ('', 'BadOp', '', VALUE_ERROR),

        # arrays
        (((0, 1),), 'Add', ((2, 3),), ((2, 4),)),
        (((0, 1),), 'Sub', ((2, 3),), ((-2, -2), )),
        (((0,), (1,)), 'Mult', ((2,), (3,)), ((0,), (3,))),
        (((0, 2), (1, 3)), 'Div', ((2, 1), (3, 2)), ((0, 2), (1 / 3, 3 / 2))),

        # ::TODO:: need error processing for arrays
    ]
)
def test_excel_operator_operand_fixup(left_op, op, right_op, expected):
    error_messages = []

    def capture_error_state(is_exception, msg):
        error_messages.append((is_exception, msg))

    assert expected == build_operator_operand_fixup(
        capture_error_state)(left_op, op, right_op)

    if expected == VALUE_ERROR:
        if expected == VALUE_ERROR and VALUE_ERROR not in (left_op, right_op):
            assert [(True, 'Values: {} {} {}'.format(
                coerce_to_number(left_op, convert_all=True), op, right_op))
            ] == error_messages

    elif expected == DIV0 and DIV0 not in (left_op, right_op):
        assert [(True, 'Values: {} {} {}'.format(left_op, op, right_op))
                ] == error_messages


def test_iterative_eval_tracker():
    assert isinstance(iterative_eval_tracker.ns.todo, set)

    def init_tracker():
        # init the tracker
        iterative_eval_tracker(iterations=100, tolerance=0.001)
        assert iterative_eval_tracker.ns.iteration_number == 0
        assert iterative_eval_tracker.ns.iterations == 100
        assert iterative_eval_tracker.ns.tolerance == 0.001
        assert iterative_eval_tracker.tolerance == 0.001
        assert iterative_eval_tracker.done

    def do_test_tracker():
        # test done if no WIP
        iterative_eval_tracker.wip(1)
        assert iterative_eval_tracker.ns.todo == {1}
        assert not iterative_eval_tracker.done
        iterative_eval_tracker.inc_iteration_number()
        assert iterative_eval_tracker.done

    init_tracker()
    do_test_tracker()

    # init the tracker
    iterative_eval_tracker(iterations=2, tolerance=5)
    assert iterative_eval_tracker.ns.iteration_number == 0
    assert iterative_eval_tracker.ns.iterations == 2
    assert iterative_eval_tracker.ns.tolerance == 5

    # test done if max iters exceeded
    iterative_eval_tracker.inc_iteration_number()
    assert iterative_eval_tracker.ns.iteration_number == 1
    iterative_eval_tracker.wip(1)
    assert not iterative_eval_tracker.done

    iterative_eval_tracker.inc_iteration_number()
    assert iterative_eval_tracker.ns.iteration_number == 2
    iterative_eval_tracker.wip(1)
    assert iterative_eval_tracker.done

    # check calced / iscalced
    assert not iterative_eval_tracker.is_calced(1)
    iterative_eval_tracker.calced(1)
    assert iterative_eval_tracker.is_calced(1)
    iterative_eval_tracker.inc_iteration_number()
    assert not iterative_eval_tracker.is_calced(1)

    class AThread(threading.Thread):
        def run(self):
            try:
                init_tracker()
                import time
                time.sleep(0.1)
                do_test_tracker()
                self.result = True
            except:  # noqa: E722
                self.result = False

    thread = AThread()
    thread.start()
    init_tracker()
    do_test_tracker()
    thread.join()
    assert thread.result
