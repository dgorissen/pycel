import datetime as dt
import math

import numpy as np
import pytest
from pycel.excellib import (
    # ::TODO:: finish test cases for remainder of functions
    _numerics,
    average,
    column,
    count,
    countif,
    countifs,
    date,
    iferror,
    index,
    isNa,
    istext,
    # linest,
    # lookup,
    match,
    mid,
    mod,
    npv,
    right,
    roundup,
    row,
    sumif,
    sumifs,
    value,
    vlookup,
    xlog,
    xmax,
    xmin,
    xround,
    xsum,
    yearfrac,
)
from pycel.excelutil import (
    AddressRange,
    DIV0,
    ERROR_CODES,
    ExcelCmp,
    NA_ERROR,
    PyCelException,
    VALUE_ERROR,
)


def test_numerics():
    assert (1, 3, 2, 3.1) == _numerics(1, '3', 2.0, pytest, 3.1, 'x')
    assert (1, 2, 3.1) == _numerics((1, '3', 2.0, pytest, 3.1, 'x'))


def test_average():
    assert 2.25 == average(1, '3', 2.0, pytest, 3, 'x')
    assert 2 == average((1, '3', 2.0, pytest, 3, 'x'))

    assert DIV0 == average(['x'])

    assert VALUE_ERROR == average(VALUE_ERROR)
    assert VALUE_ERROR == average((2, VALUE_ERROR))

    assert DIV0 == average(DIV0)
    assert DIV0 == average((2, DIV0))


@pytest.mark.parametrize(
    'address, result', (
        ('L45', 12),
        ('B:E', 2),
        ('4:7', 1),
        ('D1:E1', 4),
        ('D1:D2', 4),
        ('D1:E2', 4),
    )
)
def test_column(address, result):
    assert result == column(AddressRange.create(address))


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

    def test_year_must_be_integer(self):
        with pytest.raises(TypeError):
            date('2016', 1, 1)

    def test_month_must_be_integer(self):
        with pytest.raises(TypeError):
            date(2016, '1', 1)

    def test_day_must_be_integer(self):
        with pytest.raises(TypeError):
            date(2016, 1, '1')

    def test_year_must_be_positive(self):
        with pytest.raises(ValueError):
            date(-1, 1, 1)

    def test_year_must_have_less_than_10000(self):
        with pytest.raises(ValueError):
            date(10000, 1, 1)

    def test_result_must_be_positive(self):
        with pytest.raises(ArithmeticError):
            date(1900, 1, -1)

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


def test_iferror():
    assert 'A' == iferror('A', 2)

    for error in ERROR_CODES:
        assert 2 == iferror(error, 2)


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
    test_data = [[0, 1], [2, 3]]

    def test_array(self):

        assert 0 == index(TestIndex.test_data, 1, 1)
        assert 1 == index(TestIndex.test_data, 1, 2)
        assert 2 == index(TestIndex.test_data, 2, 1)
        assert 3 == index(TestIndex.test_data, 2, 2)

    def test_no_column_on_matrix(self):
        assert [0, 1] == index(TestIndex.test_data, 1)
        assert [2, 3] == index(TestIndex.test_data, 2)

    def test_column_on_matrix(self):
        assert [0, 2] == index(TestIndex.test_data, None, 1)
        assert [1, 3] == index(TestIndex.test_data, None, 2)

        assert (0, 2) == index(tuple(TestIndex.test_data), None, 1)
        assert (1, 3) == index(tuple(TestIndex.test_data), None, 2)

    def test_no_column_on_vector(self):
        assert 2 == index(TestIndex.test_data[1], 1)
        assert 3 == index(TestIndex.test_data[1], 2)

    def test_column_on_vector(self):
        assert 2 == index(TestIndex.test_data[1], 1, 1)
        assert 3 == index(TestIndex.test_data[1], 1, 2)

    def test_out_of_range(self):
        with pytest.raises(IndexError):
            index(TestIndex.test_data[1], 2, 2)

        with pytest.raises(IndexError):
            index(TestIndex.test_data[1], 1, 3)

        with pytest.raises(IndexError):
            index(TestIndex.test_data, None)

    def test_np_ndarray(self):
        test_data = np.asarray(self.test_data)

        assert 0 == index(test_data, 1, 1)
        assert 1 == index(test_data, 1, 2)
        assert 2 == index(test_data, 2, 1)
        assert 3 == index(test_data, 2, 2)

        assert [0, 1] == list(index(test_data, 1))
        assert [2, 3] == list(index(test_data, 2))

        assert [0, 2] == list(index(test_data, None, 1))
        assert [1, 3] == list(index(test_data, None, 2))


class TestIsNa:
    # This function might need more solid testing

    def test_isNa_false(self):
        assert not isNa('2 + 1')

    def test_isNa_true(self):
        assert isNa('x + 1')


@pytest.mark.parametrize(
    'lookup_value, lookup_array, match_type, result', (
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
    assert result == match(lookup_value, lookup_array, match_type)


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
    assert result0 == match(lookup_value, lookup_array, 0)
    assert resultm1 == match(lookup_value, lookup_array, -1)
    if result1 != match(lookup_value, lookup_array, 1):
        lookup_array = [ExcelCmp(x) for x in lookup_array]
        if sorted(lookup_array) == lookup_array:
            # only complain on failures for mode 0 when array is sorted
            assert result1 == match(lookup_value, lookup_array, 1)


class TestMid:

    def test_start_num_must_be_integer(self):
        with pytest.raises(TypeError):
            mid('Romain', 1.1, 2)

    def test_num_chars_must_be_integer(self):
        with pytest.raises(TypeError):
            mid('Romain', 1, 2.1)

    def test_start_num_must_be_superior_or_equal_to_1(self):
        with pytest.raises(ValueError):
            mid('Romain', 0, 3)

    def test_num_chars_must_be_positive(self):
        with pytest.raises(ValueError):
            mid('Romain', 1, -1)

    def test_mid(self):
        assert 'main' == mid('Romain', 2, 9)


class TestMod:

    def test_first_argument_validity(self):
        with pytest.raises(TypeError):
            mod(2.2, 1)

    def test_second_argument_validity(self):
        with pytest.raises(TypeError):
            mod(2, 1.1)

    def test_output_value(self):
        assert 2 == mod(10, 4)


def test_npv():
    assert 1188.44 == round(
        npv(0.1, -10000, 3000, 4200, 6800), 2)
    assert 1922.06 == round(
        npv(0.08, 8000, 9200, 10000, 12000, 14500) - 40000, 2)
    assert -3749.47 == round(
        npv(0.08, 8000, 9200, 10000, 12000, 14500, -9000) - 40000, 2)


def test_right():
    assert 'abcd' == right('abcd', 5)
    assert 'abcd' == right('abcd', 4)
    assert 'bcd' == right('abcd', 3)
    assert 'cd' == right('abcd', 2)
    assert 'd' == right('abcd', 1)
    assert '' == right('abcd', 0)

    assert '34' == right(1234.1, 2)

    with pytest.raises(ValueError):
        right('abcd', -1)


@pytest.mark.parametrize(
    'number, digits, result', (
            (3.2, 0, 4),
            (76.9, 0, 77),
            (3.14159, 3, 3.142),
            (-3.14159, 1, -3.2),
            (31415.92654, -2, 31500),
    )
)
def test_roundup(number, digits, result):
    assert result == roundup(number, digits)


@pytest.mark.parametrize(
    'number, digits', (
            (3.2, 'X'),
            ('X', 0),
    )
)
def test_roundup_errors(number, digits):
    with pytest.raises(TypeError):
        roundup(number, digits)


@pytest.mark.parametrize(
    'address, result', (
        ('L45', 45),
        ('B:E', 1),
        ('4:7', 4),
        ('D1:E1', 1),
        ('D1:D2', 1),
        ('D1:E2', 1),
    )
)
def test_row(address, result):
    assert result == row(AddressRange.create(address))


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


def test_value():
    assert 0.123 == value('.123')
    assert 123 == value('123')
    assert isinstance(value('123'), int)


@pytest.mark.parametrize(
    'lookup, col_idx, result', (
        ('A', 0, '#VALUE!'),
        ('A', 1, 'A'),
        ('A', 2, 1),
        ('A', 3, 'Z'),
        ('A', 4, '#REF!'),
        ('B', 1, 'B'),
        ('C', 1, 'C'),
        ('B', 2, 2),
        ('C', 2, 3),
        ('B', 3, 'Y'),
        ('C', 3, 'X'),
        ('D', 3, '#N/A'),
    )
)
def test_vlookup(lookup, col_idx, result):
    table = (
        ('A', 1, 'Z'),
        ('B', 2, 'Y'),
        ('C', 3, 'X'),
    )
    assert result == vlookup(lookup, table, col_idx)


def test_xlog():
    assert math.log(5) == xlog(5)
    assert [math.log(5), math.log(6)] == xlog([5, 6])
    assert [math.log(5), math.log(6)] == xlog((5, 6))
    assert [math.log(5), math.log(6)] == xlog(np.array([5, 6]))


def test_xmax():
    assert 0 == xmax('abcd')
    assert 3 == xmax((2, None, 'x', 3))

    assert VALUE_ERROR == xmax(VALUE_ERROR)
    assert VALUE_ERROR == xmax((2, VALUE_ERROR))

    assert DIV0 == xmax(DIV0)
    assert DIV0 == xmax((2, DIV0))


def test_xmin():
    assert 0 == xmin('abcd')
    assert 2 == xmin((2, None, 'x', 3))

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
def test_xround(result, digits):
    assert result == xround(12345.6789, digits)


@pytest.mark.parametrize(
    'number, digits, result', (
            (2.15, 1, 2.2),
            (2.149, 1, 2.1),
            (-1.475, 2, -1.48),
            (21.5, -1, 20),
            (626.3, -3, 1000),
            (1.98, -1, 0),
            (-50.55, -2, -100),
    )
)
def test_xround2(number, digits, result):
    assert result == xround(number, digits)


class TestXRound:

    def test_nb_must_be_number(self):
        with pytest.raises(TypeError):
            xround('er', 1)

    def test_nb_digits_must_be_number(self):
        with pytest.raises(TypeError):
            xround(2.323, 'ze')

    def test_positive_number_of_digits(self):
        assert 2.68 == xround(2.675, 2)

    def test_negative_number_of_digits(self):
        assert 2400 == xround(2352.67, -2)


def test_xsum():
    assert 0 == xsum('abcd')
    assert 5 == xsum((2, None, 'x', 3))

    assert VALUE_ERROR == xsum(VALUE_ERROR)
    assert VALUE_ERROR == xsum((2, VALUE_ERROR))

    assert DIV0 == xsum(DIV0)
    assert DIV0 == xsum((2, DIV0))


class TestYearfrac:

    def test_start_date_must_be_number(self):
        with pytest.raises(TypeError):
            yearfrac('not a number', 1)

    def test_end_date_must_be_number(self):
        with pytest.raises(TypeError):
            yearfrac(1, 'not a number')

    def test_start_date_must_be_positive(self):
        with pytest.raises(ValueError):
            yearfrac(-1, 0)

    def test_end_date_must_be_positive(self):
        with pytest.raises(ValueError):
            yearfrac(0, -1)

    def test_basis_must_be_between_0_and_4(self):
        with pytest.raises(ValueError):
            yearfrac(1, 2, 5)

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
        assert 11/365 == pytest.approx(
            yearfrac(date(2015, 4, 20), date(2015, 5, 1), basis=1))

        assert 11/366 == pytest.approx(
            yearfrac(date(2016, 4, 20), date(2016, 5, 1), basis=1))

        assert 316/366 == pytest.approx(
            yearfrac(date(2016, 2, 20), date(2017, 1, 1), basis=1))

        assert 61/366 == pytest.approx(
            yearfrac(date(2015, 12, 31), date(2016, 3, 1), basis=1))
