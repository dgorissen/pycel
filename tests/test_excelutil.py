import pytest
from pycel.excelutil import *


def test_is_range():

    assert is_range('a1:b2')
    assert not is_range('a1')


def test_has_sheet():

    assert has_sheet('Sheet1!a1')
    assert not has_sheet('a1')
    assert has_sheet('Sheet1!a1:b2')
    assert not has_sheet('a1:b2')


def test_get_sheet():

    assert ('', 'B1') == get_sheet('B1')
    assert ('sheet', 'B1') == get_sheet('sheet!B1')
    assert ('', 'B1:C2') == get_sheet('B1:C2')
    assert ('sheet', 'B1:C2') == get_sheet('sheet!B1:C2')

    assert ("shee't", 'B1:C2') == get_sheet("shee''t!B1:C2")
    assert ("shee t", 'B1:C2') == get_sheet("'shee t'!B1:C2")

    with pytest.raises(Exception):
        get_sheet('sh!B1', sheet='shx')
    

def test_split_range():

    assert ('', 'B1', None) == split_range('B1')
    assert ('sheet', 'B1', None) == split_range('sheet!B1')
    assert ('', 'B1', 'C2') == split_range('B1:C2')
    assert ('sheet', 'B1', 'C2') == split_range('sheet!B1:C2')


def test_split_address():
    assert ('', 'B', '1') == split_address('B1')
    assert ('sheet', 'B', '1') == split_address('sheet!B1')

    assert ('', 'A', '1') == split_address('R1C1')
    assert ('sheet', 'A', '1') == split_address('sheet!R1C1')

    assert ('', 'A', '1') == split_address('R[1]C[1]')
    assert ('sheet', 'A', '1') == split_address('sheet!R[1]C[1]')

    with pytest.raises(Exception):
        split_address('B1:C2')
        
    with pytest.raises(Exception):
        split_address('sheet!B1:C2')

    with pytest.raises(Exception):
        split_address('xyzzy')


def test_resolve_range():

    assert (['B1'], 1, 1) == resolve_range('B1')
    assert (['B1', 'C1'], 1, 2) == resolve_range('B1:C1')
    assert (['B1', 'B2'], 2, 1) == resolve_range('B1:B2')
    assert ([['B1', 'C1'], ['B2', 'C2']], 2, 2) == resolve_range('B1:C2')

    assert (['sh!B1'], 1, 1) == resolve_range('sh!B1')
    assert (['sh!B1', 'sh!C1'], 1, 2) == resolve_range('sh!B1:C1')
    assert (['sh!B1', 'sh!B2'], 2, 1) == resolve_range('sh!B1:B2')
    assert ([['sh!B1', 'sh!C1'], ['sh!B2', 'sh!C2']], 2, 2
            ) == resolve_range('sh!B1:C2')

    assert (['sh!B1'], 1, 1) == resolve_range('sh!B1', sheet='sh')
    assert (['sh!B1', 'sh!C1'], 1, 2) == resolve_range('sh!B1:C1', sheet='sh')
    assert (['sh!B1', 'sh!B2'], 2, 1) == resolve_range('sh!B1:B2', sheet='sh')
    assert ([['sh!B1', 'sh!C1'], ['sh!B2', 'sh!C2']], 2, 2
            ) == resolve_range('sh!B1:C2', sheet='sh')

    assert (['sh!B1'], 1, 1) == resolve_range('sh!B1')
    assert (['sh!B1', 'sh!C1'], 1, 2) == resolve_range('sh!B1:C1', sheet='sh')
    assert (['sh!B1', 'sh!B2'], 2, 1) == resolve_range('sh!B1:B2', sheet='sh')
    assert ([['sh!B1', 'sh!C1'], ['sh!B2', 'sh!C2']], 2, 2
            ) == resolve_range('sh!B1:C2', sheet='sh')

    with pytest.raises(Exception):
        resolve_range('sh!B1', sheet='shx')


def test_col2num():
    assert 1 == col2num('A')
    assert 1 == col2num('a')
    assert 53 == col2num('BA')
    assert 53 == col2num('Ba')

    with pytest.raises(ValueError):
        col2num('')

    with pytest.raises(AttributeError):
        col2num(2)


def test_num2col():
    assert 'A' == num2col(1)
    assert 'BA' == num2col(53)

    with pytest.raises(ValueError):
        num2col('')

    with pytest.raises(ValueError):
        num2col(0)


def test_address2index():
    assert (1, 2) == address2index('A2')

    assert (2, 1) == address2index('B1')
    assert (2, 1) == address2index('sheet!B1')

    assert (1, 2) == address2index('R2C1')
    assert (2, 1) == address2index('sheet!R1C2')

    assert (1, 2) == address2index('R[2]C[1]')
    assert (2, 1) == address2index('sheet!R[1]C[2]')

    with pytest.raises(Exception):
        address2index('B1:C2')

    with pytest.raises(Exception):
        address2index('sheet!B1:C2')

    with pytest.raises(Exception):
        address2index('xyzzy')


def test_index2address():

    assert 'B1' == index2address(2, 1)
    assert 'sh!B1' == index2address(2, 1, sheet='sh')

    assert 'A2' == index2address(1, 2)
    assert 'sh!A2' == index2address(1, 2, sheet='sh')

    assert 'A2' == index2address('A', 2)
    assert 'sh!A2' == index2address('A', 2, sheet='sh')

    assert 'A2' == index2address('A', '2')
    assert 'sh!A2' == index2address('A', '2', sheet='sh')

    assert "'shee t'!A2" == index2address('A', '2', sheet="shee t")
    assert "'shee''t'!A2" == index2address('A', '2', sheet="shee't")


def test_get_linest_degree():
    # build a spreadsheet with linest formulas horiz and vert

    class Excel:

        def __init__(self, columns, rows):
            self.columns = columns
            self.rows = rows

        def get_formula_from_range(self, address):
            sheet, col, row = split_address(address)
            found = col in self.columns and row in self.rows
            return '=linest()' if found else ''

    class Cell:
        def __init__(self, excel):
            self.excel = excel

        @property
        def sheet(self):
            return 'PhonySheet'

        @property
        def formula(self):
            return '=linest()'

        def address_parts(self):
            return self.sheet, 'E', 5, 5

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


def test_flatten():
    assert ['ddd'] == list(flatten(['ddd']))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', 1, (2, 3)]))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', (1, (2, 3))]))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', (1, 2), 3]))


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

    assert [0] == find_corresponding_index([1, 2, 3], '<2')
    assert [2] == find_corresponding_index([1, 2, 3], '>2')
    assert [0, 2] == find_corresponding_index([1, 2, 3], '<>2')
    assert [0, 1] == find_corresponding_index([1, 2, 3], '<=2')
    assert [1, 2] == find_corresponding_index([1, 2, 3], '>=2')
    assert [1] == find_corresponding_index([1, 2, 3], '2')
    assert [1] == find_corresponding_index('ABC', 'B')
    assert [1, 2] == find_corresponding_index('ABB', 'B')
    assert [] == find_corresponding_index('ABB', 'D')

    with pytest.raises(TypeError):
        find_corresponding_index('ABB', '<2')

    with pytest.raises(ValueError):
        find_corresponding_index('ABB', None)
