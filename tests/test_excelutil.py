import os
import pickle
from collections import namedtuple

import pytest
from openpyxl.utils import column_index_from_string, quote_sheetname
from pycel.excelutil import (
    AddressCell,
    AddressRange,
    assert_list_like,
    build_operator_operand_fixup,
    coerce_to_number,
    coerce_to_string,
    criteria_parser,
    date_from_int,
    ExcelCmp,
    find_corresponding_index,
    flatten,
    get_linest_degree,
    get_max_days_in_month,
    in_array_formula_context,
    is_leap_year,
    is_number,
    list_like,
    MAX_COL,
    MAX_ROW,
    NUM_ERROR,
    normalize_year,
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


class ATestCell:

    def __init__(self, col, row, sheet='', excel=None):
        self.row = row
        self.col = col
        self.col_idx = column_index_from_string(col)
        self.sheet = sheet
        self.excel = excel
        self.address = AddressCell(
            '{}{}'.format(col, row), sheet=sheet)


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


def test_address_range_errors():

    with pytest.raises(ValueError):
        AddressRange('B32:B')


@pytest.mark.parametrize(
    'left, right, result', (
        ('a1:b2', 'b1:c3', 'a1:c3'),
        ('a1:b2', 'd5', 'a1:d5'),
        ('a1:d5', 'b3', 'a1:d5'),
        ('d4:e5', 'a1', 'a1:e5'),
        ('c4:e5', 'd1', 'c1:e5'),
        ('c4:e6', 'a5', 'a4:e6'),
        ('c4:e5', 'd9', 'c4:e9'),
        ('c4:e6', 'j5', 'c4:j6'),
    )
)
def test_address_range_add(left, right, result):
    assert AddressRange(left) + AddressRange(right) == AddressRange(result)


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


def test_is_range():

    assert AddressRange('a1:b2').is_range
    assert not AddressRange('a1').is_range


def test_has_sheet():

    assert AddressRange('Sheet1!a1').has_sheet
    assert not AddressRange('a1').has_sheet
    assert AddressRange('Sheet1!a1:b2').has_sheet
    assert not AddressRange('a1:b2').has_sheet

    assert AddressCell('sh!a1') == AddressRange(AddressRange('a1'), sheet='sh')

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
    addr = AddressCell('A1', sheet=sheet_name)
    assert addr.quoted_address == '{}!A1'.format(quote_sheetname(sheet_name))


def test_split_sheetname():

    assert ('', 'B1') == split_sheetname('B1')
    assert ('sheet', 'B1') == split_sheetname('sheet!B1')
    assert ('', 'B1:C2') == split_sheetname('B1:C2')
    assert ('sheet', 'B1:C2') == split_sheetname('sheet!B1:C2')

    assert ("shee't", 'B1:C2') == split_sheetname("'shee''t'!B1:C2")
    assert ("shee t", 'B1:C2') == split_sheetname("'shee t'!B1:C2")

    with pytest.raises(ValueError):
        split_sheetname('sh!B1', sheet='shx')


def test_address_cell_enum():
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


def test_extended_range_boundaries():
    cell = ATestCell('A', 1)

    assert (1, 2) * 2 == range_boundaries('A2')[0]
    assert (2, 1) * 2 == range_boundaries('B1')[0]
    assert (1, 2) * 2 == range_boundaries('R2C1')[0]
    assert (2, 1) * 2 == range_boundaries('R1C2')[0]
    assert (2, 3) * 2 == range_boundaries('R[2]C[1]', cell)[0]
    assert (3, 2) * 2 == range_boundaries('R[1]C[2]', cell)[0]

    assert (1, 1, 2, 2) == range_boundaries('A1:B2')[0]
    assert (1, 1, 2, 2) == range_boundaries('R1C1:R2C2')[0]
    assert (2, 1, 2, 3) == range_boundaries('R1C2:R[2]C[1]', cell)[0]

    assert (3, 13) * 2 == range_boundaries('R13C3')[0]

    assert (1, 1, 1, 1) == range_boundaries('RC', cell)[0]

    assert (None, 1, None, 4) == range_boundaries('R:R[3]', cell)[0]
    assert (None, 1, None, 4) == range_boundaries('R1:R[3]', cell)[0]
    assert (None, 2, None, 4) == range_boundaries('R2:R[3]', cell)[0]

    assert (1, None, 4, None) == range_boundaries('C:C[3]', cell)[0]
    assert (1, None, 4, None) == range_boundaries('C1:C[3]', cell)[0]
    assert (2, None, 4, None) == range_boundaries('C2:C[3]', cell)[0]

    with pytest.raises(NotImplementedError, match='Multiple Colon Ranges'):
        range_boundaries('A1:B2:C3')


def test_range_boundaries_defined_names(excel):
    cell = ATestCell('A', 1, excel=excel)

    assert ((3, 1, 3, 18), 'Sheet1') == range_boundaries('SINUS', cell)


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
def test_extended_range_boundaries_errors(address_string):
    cell = ATestCell('A', 1)

    with pytest.raises(ValueError, match='not a valid coordinate or range'):
        range_boundaries(address_string, cell)


def test_coerce_to_number():
    assert 1 == coerce_to_number(1)
    assert 1.0 == coerce_to_number(1.0)

    assert coerce_to_number(None) is None

    assert 1 == coerce_to_number('1')
    assert isinstance(coerce_to_number('1'), int)

    assert 1 == coerce_to_number('1.')
    assert isinstance(coerce_to_number('1.'), float)

    assert 'xyzzy' == coerce_to_number('xyzzy')

    with pytest.raises(ZeroDivisionError):
        coerce_to_number(DIV0)


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
        return in_array_formula_context

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


def test_is_number():
    assert is_number(1)
    assert is_number(0)
    assert is_number(-1)
    assert is_number(1.0)
    assert is_number(0.0)
    assert is_number(-1.0)
    assert is_number('1.0')
    assert is_number('0.0')
    assert is_number('-1.0')
    assert is_number(True)
    assert is_number(False)

    assert not is_number(None)
    assert not is_number('x')


def test_is_leap_year():

    assert is_leap_year(1900)
    assert is_leap_year(1904)
    assert is_leap_year(2000)
    assert is_leap_year(2104)

    assert not is_leap_year(1)
    assert not is_leap_year(2100)
    assert not is_leap_year(2101)
    assert not is_leap_year(2103)
    assert not is_leap_year(2102)

    with pytest.raises(TypeError):
        is_leap_year('x')

    with pytest.raises(TypeError):
        is_leap_year(-1)

    with pytest.raises(TypeError):
        is_leap_year(0)


def test_get_max_days_in_month():
    assert 31 == get_max_days_in_month(1, 2000)
    assert 29 == get_max_days_in_month(2, 2000)
    assert 28 == get_max_days_in_month(2, 2001)
    assert 31 == get_max_days_in_month(3, 2000)
    assert 30 == get_max_days_in_month(4, 2000)
    assert 31 == get_max_days_in_month(5, 2000)
    assert 30 == get_max_days_in_month(6, 2000)
    assert 31 == get_max_days_in_month(7, 2000)
    assert 31 == get_max_days_in_month(8, 2000)
    assert 30 == get_max_days_in_month(9, 2000)
    assert 31 == get_max_days_in_month(10, 2000)
    assert 30 == get_max_days_in_month(11, 2000)
    assert 31 == get_max_days_in_month(12, 2000)

    # excel thinks 1900 is a leap year
    assert 29 == get_max_days_in_month(2, 1900)


def test_normalize_year():
    assert (1900, 1, 1) == normalize_year(1900, 1, 1)
    assert (1900, 2, 1) == normalize_year(1900, 1, 32)
    assert (1900, 3, 1) == normalize_year(1900, 1, 61)
    assert (1900, 4, 1) == normalize_year(1900, 1, 92)
    assert (1900, 5, 1) == normalize_year(1900, 1, 122)
    assert (1900, 4, 1) == normalize_year(1900, 0, 123)
    assert (1900, 3, 1) == normalize_year(1900, -1, 122)

    assert (1899, 12, 1) == normalize_year(1900, 1, -31)
    assert (1899, 12, 1) == normalize_year(1900, 0, 1)
    assert (1899, 11, 1) == normalize_year(1900, -1, 1)


def test_date_from_int():
    assert (1900, 1, 1) == date_from_int(1)
    assert (1900, 1, 31) == date_from_int(31)
    assert (1900, 2, 29) == date_from_int(60)
    assert (1900, 3, 1) == date_from_int(61)

    assert (2009, 7, 6) == date_from_int(40000)


def test_find_corresponding_index():
    assert (0,) == find_corresponding_index([1, 2, 3], '<2')
    assert (2,) == find_corresponding_index([1, 2, 3], '>2')
    assert (0, 2) == find_corresponding_index([1, 2, 3], '<>2')
    assert (0, 1) == find_corresponding_index([1, 2, 3], '<=2')
    assert (1, 2) == find_corresponding_index([1, 2, 3], '>=2')
    assert (1,) == find_corresponding_index([1, 2, 3], '2')
    assert (1,) == find_corresponding_index(list('ABC'), 'B')
    assert (1, 2) == find_corresponding_index(list('ABB'), 'B')
    assert (1, 2) == find_corresponding_index(list('ABB'), '<>A')
    assert () == find_corresponding_index(list('ABB'), 'D')

    with pytest.raises(TypeError):
        find_corresponding_index('ABB', '<B')

    with pytest.raises(ValueError):
        find_corresponding_index(list('ABB'), None)


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
            assert [(True, 'Values: {} {} {}'.format(left_op, op, right_op))
                    ] == error_messages

    elif expected == DIV0 and DIV0 not in (left_op, right_op):
        assert [(True, 'Values: {} {} {}'.format(left_op, op, right_op))
                ] == error_messages
