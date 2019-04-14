import datetime as dt
import math

import numpy as np
import pytest
import pycel.excellib
from pycel.excellib import (
    _numerics,
    average,
    column,
    concat,
    concatenate,
    count,
    countif,
    countifs,
    date,
    floor,
    hlookup,
    index,
    isNa,
    istext,
    # ::TODO:: finish test cases for remainder of functions
    # linest,
    ln,
    log,
    lookup,
    match,
    _match,
    mid,
    mod,
    npv,
    power,
    right,
    roundup,
    row,
    sumif,
    sumifs,
    trunc,
    vlookup,
    x_abs,
    xatan2,
    x_int,
    x_len,
    xmax,
    xmin,
    x_round,
    xsum,
    yearfrac,
)
from pycel.excelutil import (
    AddressRange,
    DIV0,
    ExcelCmp,
    NA_ERROR,
    NAME_ERROR,
    NUM_ERROR,
    PyCelException,
    REF_ERROR,
    VALUE_ERROR,
)

from pycel.lib.function_helpers import error_string_wrapper, load_to_test_module


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.excellib, __name__)


def test_numerics():
    assert (1, 2, 3.1) == _numerics(1, '3', 2.0, pytest, 3.1, 'x')
    assert (1, 2, 3.1) == _numerics((1, '3', (2.0, pytest, 3.1), 'x'))


def test_average():
    assert 2 == average(1, '3', 2.0, pytest, 3, 'x')
    assert 2 == average((1, '3', (2.0, pytest, 3), 'x'))

    assert -0.1 == average((-0.1, None, 'x', True))

    assert DIV0 == average(['x'])

    assert VALUE_ERROR == average(VALUE_ERROR)
    assert VALUE_ERROR == average((2, VALUE_ERROR))

    assert DIV0 == average(DIV0)
    assert DIV0 == average((2, DIV0))


@pytest.mark.parametrize(
    'value, expected', (
        (1, 1),
        (-2, 2),
        (((2, -3, 4, -5), ), ((2, 3, 4, 5), )),
        (DIV0, DIV0),
        (NUM_ERROR, NUM_ERROR),
        (VALUE_ERROR, VALUE_ERROR),
    )
)
def test_x_abs(value, expected):
    assert x_abs(value) == expected


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
    'args, result', (
        ('a 1 abc'.split(), 'a1abc'),
        ('a Jan-00 abc'.split(), 'aJan-00abc'),
        ('a	#DIV/0! abc'.split(), DIV0),
        ('a	1 #DIV/0!'.split(), DIV0),
        ('a #NAME? abc'.split(), NAME_ERROR),
        (('a', True, 'abc'), 'aTRUEabc'),
        (('a', False, 'abc'), 'aFALSEabc'),
        (('a', 2, 'abc'), 'a2abc'),
    )
)
def test_concatenate(args, result):
    assert concat(*args) == result
    assert concatenate(*args) == result
    assert concat(args) == result
    assert concatenate(args) == VALUE_ERROR


class TestCount:

    def test_without_nested_booleans(self):
        assert 3 == count([1, 2, 'e'], True, 'r')

    def test_with_nested_booleans(self):
        assert 2 == count([1, True, 'e'], True, 'r')

    def test_with_text_representations(self):
        assert 4 == count([1, '2.2', 'e'], True, '20')

    def test_with_date_representations(self):
        assert 4 == count([1, '2.2', dt.datetime.now()], True, '20')


class TestCountIf:

    def test_countif_strictly_superior(self):
        assert 3 == countif([7, 25, 13, 25], '>10')

    def test_countif_strictly_inferior(self):
        assert 1 == countif([7, 25, 13, 25], '<10')

    def test_countif_superior(self):
        assert 3 == countif([7, 10, 13, 25], '>=10')

    def test_countif_inferior(self):
        assert 2 == countif([7, 10, 13, 25], '<=10')

    def test_countif_different(self):
        assert 3 == countif([7, 10, 13, 25], '<>10')

    def test_countif_with_string_equality(self):
        assert 2 == countif([7, 'e', 13, 'e'], 'e')

    def test_countif_with_string_inequality(self):
        assert 1 == countif([7, 'e', 13, 'f'], '>e')

    def test_countif_regular(self):
        assert 2 == countif([7, 25, 13, 25], 25)


class TestCountIfs:
    # more tests might be welcomed

    def test_countifs_regular(self):
        assert 1 == countifs([7, 25, 13, 25], 25, [100, 102, 201, 20], ">100")

    def test_countifs_odd_args_len(self):
        with pytest.raises(PyCelException):
            countifs([7, 25, 13, 25], 25, [100, 102, 201, 20])


class TestDate:

    def test_values_can_str(self):
        assert date('2016', 1, 1) == date(2016, '1', 1) == date(2016, 1, '1')

    def test_year_must_be_positive(self):
        assert NUM_ERROR == date(-1, 1, 1)

    def test_year_must_have_less_than_10000(self):
        assert NUM_ERROR == date(10000, 1, 1)

    def test_result_must_be_positive(self):
        assert NUM_ERROR == date(1900, 1, -1)

    def test_not_stricly_positive_month_substracts(self):
        assert date(2009, -1, 1) == date(2008, 11, 1)

    def test_not_stricly_positive_day_substracts(self):
        assert date(2009, 1, -1) == date(2008, 12, 30)

    def test_month_superior_to_12_change_year(self):
        assert date(2009, 14, 1) == date(2010, 2, 1)

    def test_day_superior_to_365_change_year(self):
        assert date(2009, 1, 400) == date(2010, 2, 4)

    def test_year_for_29_feb(self):
        assert 39507 == date(2008, 2, 29)

    def test_year_regular(self):
        assert 39755 == date(2008, 11, 3)

    def test_year_offset(self):
        zero = dt.datetime(1900, 1, 1) - dt.timedelta(2)
        assert (dt.datetime(1900, 1, 1) - zero).days == date(0, 1, 1)
        assert (dt.datetime(1900 + 1899, 1, 1) - zero).days == date(1899, 1, 1)
        assert (dt.datetime(1900 + 1899, 1, 1) - zero).days == date(1899, 1, 1)


@pytest.mark.parametrize(
    'number, significance, result', (
        (2.5, 1, 2),
        (2.5, 2, 2),
        (2.5, 3, 0),
        (-2.5, -1, -2),
        (-2.5, -2, -2),
        (-2.5, -3, 0),
        (0, 0, 0),
        (-2.5, 0, DIV0),
        (-1, 1, NUM_ERROR),
        (1, -1, NUM_ERROR),
    )
)
def test_floor(number, significance, result):
    assert floor(number, significance) == result


@pytest.mark.parametrize(
    'lkup, col_idx, result, approx', (
        ('A', 0, VALUE_ERROR, True),
        ('A', 1, 'A', True),
        ('A', 2, 1, True),
        ('A', 3, 'Z', True),
        ('A', 4, REF_ERROR, True),
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
def test_hlookup(lkup, col_idx, result, approx):
    table = (
        ('A', 'B', 'C'),
        (1, 2, 3),
        ('Z', 'Y', 'X'),
    )
    assert result == hlookup(lkup, table, col_idx, approx)


@pytest.mark.parametrize(
    'values, expected', (
        ((1, 1, 1, 1), NA_ERROR),
        ((1, ((1, 2), (3, 4)), 1, 1), 1),
        ((REF_ERROR, ((1, 2), (3, 4)), 1, 1), REF_ERROR),
        ((1, REF_ERROR, 1, 1), REF_ERROR),
        ((1, ((1, 2), (3, 4)), REF_ERROR, 1), REF_ERROR),
        ((1, ((1, 2), (3, 4)), 1, REF_ERROR), REF_ERROR),
        ((1, ((1, 2), (3, 4)), 0, 1), VALUE_ERROR),
        ((1, ((1, 2), (3, 4)), 3, 1), REF_ERROR),
    )
)
def test_hlookup_vlookup_error(values, expected):
    assert hlookup(*values) == expected
    assert vlookup(*values) == expected


def test_is_text():
    assert istext('a')
    assert not istext(1)
    assert not istext(None)


class TestIndex:
    """
    Description
    Returns the value of an element in a table or an array, selected by the row
    and column number indexes.

    Use the array form if the first argument to INDEX is an array constant.

    Syntax
    INDEX(array, row_num, [column_num])

    The INDEX function syntax has the following arguments.

    Array    Required. A range of cells or an array constant.

    If array contains only one row or column, the corresponding Row_num or
    Column_num argument is optional.

    If array has more than one row and more than one column, and only Row_num
    or Column_num is used, INDEX returns an array of the entire row or column
    in array.

    Row_num    Required. Selects the row in array from which to return a value.
               If Row_num is omitted, Column_num is required.

    Column_num    Optional. Selects the column in array from which to return a
                value. If Column_num is omitted, Row_num is required.

    Remarks
    If both the Row_num and Column_num arguments are used, INDEX returns the
    value in the cell at the intersection of Row_num and Column_num.

    If you set Row_num or Column_num to 0 (zero), INDEX returns the array of
    values for the entire column or row, respectively. To use values returned
    as an array, enter the INDEX function as an array formula in a horizontal
    range of cells for a row, and in a vertical range of cells for a column.
    To enter an array formula, press CTRL+SHIFT+ENTER.

    Note: In Excel Web App, you cannot create array formulas.

    Row_num and Column_num must point to a cell within array; otherwise,
    INDEX returns the #REF! error value.

    """
    test_data = ((0, 1), (2, 3))
    test_data_col = ((0,), (2,))
    test_data_row = ((0, 1),)
    test_data_np = np.asarray(test_data)

    @staticmethod
    @pytest.mark.parametrize(
        'data, row_num, col_num, expected', (
            (test_data, 1, 1, 0),
            (test_data, 1, 2, 1),
            (test_data, 2, 1, 2),
            (test_data, 2, 2, 3),

            # no col given
            (test_data, 1, None, ((0, 1),)),
            (test_data, 2, None, ((2, 3),)),
            (test_data_col, 1, None, 0),
            (test_data_col, 2, None, 2),
            (test_data_row, 1, None, 0),
            (test_data_row, 2, None, 1),

            # no row given
            (test_data, None, 1, ((0,), (2,))),
            (test_data, None, 2, ((1,), (3,))),
            (test_data_col, None, 1, 0),
            (test_data_col, None, 2, 2),
            (test_data_row, None, 1, 0),
            (test_data_row, None, 2, 1),

            # OOR
            (test_data_row, 2, 2, NA_ERROR),
            (test_data_col, 1, 3, NA_ERROR),
            (test_data, None, None, NA_ERROR),

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
    def test_index(data, row_num, col_num, expected):
        result = index(data, row_num, col_num)
        if isinstance(expected, np.ndarray):
            assert (result == expected).all()
        else:
            assert result == expected

    @staticmethod
    def test_index_error_inputs():
        index_f = error_string_wrapper(index)
        assert NA_ERROR == index_f(NA_ERROR, 1)
        assert NA_ERROR == index_f(TestIndex.test_data, NA_ERROR, 1)
        assert NA_ERROR == index_f(TestIndex.test_data, 1, NA_ERROR)
        assert VALUE_ERROR == index_f(None, 1, 1)


class TestIsNa:
    # This function might need more solid testing

    def test_isNa_false(self):
        assert not isNa('2 + 1')

    def test_isNa_true(self):
        assert isNa('x + 1')


@pytest.mark.parametrize(
    'expected, value',
    (
        (math.log(5), 5),
        (math.log(2), 2),
        (NUM_ERROR, None),
        (VALUE_ERROR, VALUE_ERROR),
        (DIV0, DIV0),
        (math.log(5), 5),
        (((math.log(5), math.log(6)), ), ((5, 6), )),
        (((math.log(5), math.log(6)), ), ((5, 6), )),
        # (((math.log(5), math.log(6)), ), np.array(((5, 6), ))),
        (NUM_ERROR, None),
        (VALUE_ERROR, VALUE_ERROR),
        (((math.log(2), VALUE_ERROR), ), ((2, VALUE_ERROR), )),
        (DIV0, DIV0),
        (((math.log(2), DIV0), ), ((2, DIV0), )),
    )
)
def test_ln(expected, value):
    assert ln(value) == expected


@pytest.mark.parametrize(
    'expected, value',
    (
        (math.log(5, 10), 5),
        (math.log(2, 10), 2),
        (NUM_ERROR, None),
        (VALUE_ERROR, VALUE_ERROR),
        (DIV0, DIV0),
        (math.log(5, 10), 5),
        (((math.log(5, 10), math.log(6, 10)), ), ((5, 6), )),
        (((math.log(5, 10), math.log(6, 10)), ), ((5, 6), )),
        # (((math.log(5), math.log(6)), ), np.array(((5, 6), ))),
        (NUM_ERROR, None),
        (VALUE_ERROR, VALUE_ERROR),
        (((math.log(2, 10), VALUE_ERROR), ), ((2, VALUE_ERROR), )),
        (DIV0, DIV0),
        (((math.log(2, 10), DIV0), ), ((2, DIV0), )),
    )
)
def test_log(expected, value):
    assert log(value) == expected


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
    assert NA_ERROR == lookup(1, 1)


@pytest.mark.parametrize(
    'lookup_value, lookup_array, match_type, result', (
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
def test_match(lookup_value, lookup_array, match_type, result):
    lookup_row = (tuple(lookup_array), )
    lookup_col = tuple((i, ) for i in lookup_array)
    assert result == match(lookup_value, lookup_row, match_type)
    assert result == match(lookup_value, lookup_col, match_type)


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


class TestMid:

    def test_invalid_parameters(self):
        assert mid(VALUE_ERROR, 2, 2) == VALUE_ERROR
        assert mid('Romain', VALUE_ERROR, 2) == VALUE_ERROR
        assert mid('Romain', 2, VALUE_ERROR) == VALUE_ERROR
        assert mid(DIV0, 2, 2) == DIV0
        assert mid('Romain', DIV0, 2) == DIV0
        assert mid('Romain', 2, DIV0) == DIV0

        assert mid('Romain', 'x', 2) == VALUE_ERROR
        assert mid('Romain', 2, 'x') == VALUE_ERROR

    def test_num_chars_must_be_integer(self):
        assert 'Ro' == mid('Romain', 1, 2.1)

    def test_start_num_must_be_superior_or_equal_to_1(self):
        assert VALUE_ERROR == mid('Romain', 0, 3)

    def test_num_chars_must_be_positive(self):
        assert VALUE_ERROR == mid('Romain', 1, -1)

    def test_from_not_str(self):
        assert '23' == mid(1234, 2, 2)

    def test_mid(self):
        assert 'omain' == mid('Romain', 2, 9)
        assert 'om' == mid('Romain', 2.1, 2)
        assert 'om' == mid('Romain', 2, 2.1)


class TestMod:

    def test_first_argument_validity(self):
        assert mod(VALUE_ERROR, 1) == VALUE_ERROR
        assert mod('x', 1) == VALUE_ERROR

    def test_second_argument_validity(self):
        assert mod(2, VALUE_ERROR) == VALUE_ERROR
        assert mod(2, 'x') == VALUE_ERROR
        assert mod(2, 0) == DIV0
        assert mod(2, None) == DIV0

    def test_output_value(self):
        assert 2 == mod(10, 4)
        assert mod(2.2, 1) == pytest.approx(0.2)
        assert mod(2, 1.1) == pytest.approx(0.9)


@pytest.mark.parametrize(
    'data, expected', (
        ((0.1, -10000, 3000, 4200, 6800), 1188.44),
        ((0.08, 8000, 9200, 10000, 12000, 14500), 41922.06),
        ((0.08, 8000, 9200, 10000, 12000, 14500, -9000), 40000 - 3749.47),
        ((NA_ERROR, 8000, 9200, 10000, 12000, 14500, -9000), NA_ERROR),
        ((0.08, 8000, DIV0, 10000, 12000, 14500, -9000), DIV0),
    )
)
def test_npv(data, expected):
    result = npv(*data)

    if isinstance(result, str):
        assert result == expected
    else:
        assert result == pytest.approx(expected, rel=1e-3)


@pytest.mark.parametrize(
    'data, expected', (
        ((0, 0), NA_ERROR),
        ((0, 1), 0),
        ((1, 0), 1),
        ((1, 2), 1),
        ((2, 1), 2),
        ((2, -1), 0.5),
        ((-2, 1), -2),
        ((0.1, 0.1), 0.1 ** 0.1),
        ((True, 1), 1),
        (('x', 1), VALUE_ERROR),
        ((1, 'x'), VALUE_ERROR),
        ((NA_ERROR, 1), NA_ERROR),
        ((1, NA_ERROR), NA_ERROR),
        ((0, -1), DIV0),
        ((1, DIV0), DIV0),
        ((DIV0, 1), DIV0),
        ((NA_ERROR, DIV0), NA_ERROR),
        ((DIV0, NA_ERROR), DIV0),
    )
)
def test_power(data, expected):
    result = power(*data)
    if isinstance(result, str):
        assert result == expected
    else:
        assert result == pytest.approx(expected, rel=1e-3)


@pytest.mark.parametrize(
    'text, num_chars, expected', (
        ('abcd', 5, 'abcd'),
        ('abcd', 4, 'abcd'),
        ('abcd', 3, 'bcd'),
        ('abcd', 2, 'cd'),
        ('abcd', 1, 'd'),
        ('abcd', 0, ''),

        (1234.1, 2, '.1'),

        ('abcd', -1, VALUE_ERROR),
        ('abcd', 'x', VALUE_ERROR),
        (VALUE_ERROR, 1, VALUE_ERROR),
        ('abcd', VALUE_ERROR, VALUE_ERROR),
    )
)
def test_right(text, num_chars, expected):
    assert right(text, num_chars) == expected


@pytest.mark.parametrize(
    'number, digits, result', (
        (3.2, 0, 4),
        (76.9, 0, 77),
        (3.14159, 3, 3.142),
        (-3.14159, 1, -3.2),
        (31415.92654, -2, 31500),
        (None, -2, 0),
        (True, -2, 100),
        (3.2, 'X', VALUE_ERROR),
        ('X', 0, VALUE_ERROR),
        (3.2, VALUE_ERROR, VALUE_ERROR),
        (VALUE_ERROR, 0, VALUE_ERROR),
    )
)
def test_roundup(number, digits, result):
    assert result == roundup(number, digits)


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


class TestSumIf:

    def test_range_is_a_list(self):
        with pytest.raises(TypeError):
            sumif(12, 12)

    def test_sum_range_is_a_list(self):
        with pytest.raises(TypeError):
            sumif(12, 12, 12)

    def test_regular_with_number_criteria(self):
        assert 6 == sumif([1, 1, 2, 2, 2], 2)

    def test_regular_with_string_criteria(self):
        assert 12 == sumif([1, 2, 3, 4, 5], ">=3")

    def test_sum_range(self):
        assert 668 == sumif([1, 2, 3, 4, 5], ">=3", [100, 123, 12, 23, 633])

    def test_sum_range_with_more_indexes(self):
        assert 668 == sumif([1, 2, 3, 4, 5], ">=3", [100, 123, 12, 23, 633, 1])

    def test_sum_range_with_less_indexes(self):
        assert 35 == sumif([1, 2, 3, 4, 5], ">=3", [100, 123, 12, 23])

    def test_sum_range_not_list(self):
        with pytest.raises(TypeError):
            sumif([], [], 'JUNK')


class TestSumIfs:

    def test_range_is_a_list(self):
        with pytest.raises(TypeError):
            sumifs(12, 12)

    def test_sum_range_is_a_list(self):
        with pytest.raises(TypeError):
            sumifs(12, 12, 12)

    def test_regular_with_number_criteria(self):
        assert 6 == sumifs([1, 1, 2, 2, 2], [1, 1, 2, 2, 2], 2)

    def test_regular_with_string_criteria(self):
        assert 12 == sumifs([1, 2, 3, 4, 5], [1, 2, 3, 4, 5], ">=3")

    def test_sum_range(self):
        assert 668 == sumifs([100, 123, 12, 23, 633], [1, 2, 3, 4, 5], ">=3")

    def test_sum_range_with_more_indexes(self):
        assert 668 == sumifs([100, 123, 12, 23, 633, 1], [1, 2, 3, 4, 5], ">=3")

    def test_sum_range_with_less_indexes(self):
        assert 35 == sumifs([100, 123, 12, 23], [1, 2, 3, 4, 5], ">=3")

    def test_sum_range_with_empty(self):
        assert 35 == sumifs([100, 123, 12, 23, None], [1, 2, 3, 4, 5], ">=3")

    def test_sum_range_not_list(self):
        with pytest.raises(TypeError):
            sumifs('JUNK', [], [], )

    def test_multiple_criteria(self):
        assert 7 == sumifs([1, 2, 3, 4, 5],
                           [1, 2, 3, 4, 5], ">=3",
                           [1, 2, 3, 4, 5], "<=4")


@pytest.mark.parametrize(
    'number, num_digits, result', (
        (2.5, -1, 0),
        (2.5, 0, 2),
        (2.5, 1, 2.5),
        (-2.5, -1, 0),
        (-2.5, 0, -2),
        (-2.5, 1, -2.5),
        (NUM_ERROR, 1, NUM_ERROR),
        (1, NUM_ERROR, NUM_ERROR),
    )
)
def test_trunc(number, num_digits, result):
    assert trunc(number, num_digits) == result


@pytest.mark.parametrize(
    'lkup, col_idx, result, approx', (
        ('A', 0, VALUE_ERROR, True),
        ('A', 1, 'A', True),
        ('A', 2, 1, True),
        ('A', 3, 'Z', True),
        ('A', 4, REF_ERROR, True),
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
        ('A', 1, 'Z'),
        ('B', 2, 'Y'),
        ('C', 3, 'X'),
    )
    assert result == vlookup(lkup, table, col_idx, approx)


@pytest.mark.parametrize(
    'param1, param2, result', (
        (1, 1, math.pi / 4),
        (1, 0, 0),
        (0, 1, math.pi / 2),
        (NA_ERROR, 1, NA_ERROR),
        (1, NA_ERROR, NA_ERROR),
        (DIV0, 1, DIV0),
        (1, DIV0, DIV0),
    )
)
def test_xatan2(param1, param2, result):
    assert xatan2(param1, param2) == result


@pytest.mark.parametrize(
    'value, expected', (
        (1, 1),
        (1.2, 1),
        (-1.2, -2),
        (((2.1, -3.9, 4.6, -5.3),), ((2, -4, 4, -6),)),
        (DIV0, DIV0),
        (NUM_ERROR, NUM_ERROR),
        (VALUE_ERROR, VALUE_ERROR),
    )
)
def test_x_int(value, expected):
    assert x_int(value) == expected


@pytest.mark.parametrize(
    'param, result', (
        ('A', 1),
        ('BB', 2),
        (3.0, 3),
        (True, 4),
        (False, 5),
        (None, 0),
        (NA_ERROR, NA_ERROR),
        (DIV0, DIV0),
    )
)
def test_x_len(param, result):
    assert x_len(param) == result


def test_xmax():
    assert 0 == xmax('abcd')
    assert 3 == xmax((2, None, 'x', 3))

    assert -0.1 == xmax((-0.1, None, 'x', True))

    assert VALUE_ERROR == xmax(VALUE_ERROR)
    assert VALUE_ERROR == xmax((2, VALUE_ERROR))

    assert DIV0 == xmax(DIV0)
    assert DIV0 == xmax((2, DIV0))


def test_xmin():
    assert 0 == xmin('abcd')
    assert 2 == xmin((2, None, 'x', 3))

    assert -0.1 == xmin((-0.1, None, 'x', True))

    assert VALUE_ERROR == xmin(VALUE_ERROR)
    assert VALUE_ERROR == xmin((2, VALUE_ERROR))

    assert DIV0 == xmin(DIV0)
    assert DIV0 == xmin((2, DIV0))


@pytest.mark.parametrize(
    'result, digits', (
        (0, -5),
        (10000, -4),
        (12000, -3),
        (12300, -2),
        (12350, -1),
        (12346, 0),
        (12345.7, 1),
        (12345.68, 2),
        (12345.679, 3),
        (12345.6789, 4),
    )
)
def test_x_round(result, digits):
    assert result == x_round(12345.6789, digits)
    assert result == x_round(12345.6789, digits + (-0.9 if digits < 0 else 0.9))


@pytest.mark.parametrize(
    'number, digits, result', (
        (2.15, 1, 2.2),
        (2.149, 1, 2.1),
        (-1.475, 2, -1.48),
        (21.5, -1, 20),
        (626.3, -3, 1000),
        (1.98, -1, 0),
        (-50.55, -2, -100),
        (DIV0, 1, DIV0),
        (1, DIV0, DIV0),
        ('er', 1, VALUE_ERROR),
        (2.323, 'ze', VALUE_ERROR),
        (2.675, 2, 2.68),
        (2352.67, -2, 2400),
        ("2352.67", "-2", 2400),
    )
)
def test_x_round2(number, digits, result):
    assert result == x_round(number, digits)


def test_xsum():
    assert 0 == xsum('abcd')
    assert 5 == xsum((2, None, 'x', 3))

    assert -0.1 == xsum((-0.1, None, 'x', True))

    assert VALUE_ERROR == xsum(VALUE_ERROR)
    assert VALUE_ERROR == xsum((2, VALUE_ERROR))

    assert DIV0 == xsum(DIV0)
    assert DIV0 == xsum((2, DIV0))


class TestYearfrac:

    def test_start_date_must_be_number(self):
        assert VALUE_ERROR == yearfrac('not a number', 1)

    def test_end_date_must_be_number(self):
        assert VALUE_ERROR == yearfrac(1, 'not a number')

    def test_start_date_must_be_positive(self):
        assert NUM_ERROR == yearfrac(-1, 0)

    def test_end_date_must_be_positive(self):
        assert NUM_ERROR == yearfrac(0, -1)

    def test_basis_must_be_between_0_and_4(self):
        assert NUM_ERROR == yearfrac(1, 2, 5)

    def test_yearfrac_basis_0(self):
        assert 7.30277777777778 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20)))

    def test_yearfrac_basis_1(self):
        assert 7.299110198 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 1))

    def test_yearfrac_basis_2(self):
        assert 7.405555556 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 2))

    def test_yearfrac_basis_3(self):
        assert 7.304109589 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 3))

    def test_yearfrac_basis_4(self):
        assert 7.302777778 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 4))

    def test_yearfrac_inverted(self):
        assert yearfrac(date(2008, 1, 1), date(2015, 4, 20)) == pytest.approx(
            yearfrac(date(2015, 4, 20), date(2008, 1, 1)))

    def test_yearfrac_basis_1_sub_year(self):
        assert 11 / 365 == pytest.approx(
            yearfrac(date(2015, 4, 20), date(2015, 5, 1), 1))

        assert 11 / 366 == pytest.approx(
            yearfrac(date(2016, 4, 20), date(2016, 5, 1), 1))

        assert 316 / 366 == pytest.approx(
            yearfrac(date(2016, 2, 20), date(2017, 1, 1), 1))

        assert 61 / 366 == pytest.approx(
            yearfrac(date(2015, 12, 31), date(2016, 3, 1), 1))
