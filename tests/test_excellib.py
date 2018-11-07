import datetime as dt
import math
import numpy as np
import pytest

from pycel.excellib import (
    # ::TODO:: finish test cases for remainder of functions
    _numerics,
    average,
    count,
    countif,
    countifs,
    date,
    iferror,
    index,
    istext,
    isNa,
    linest,
    lookup,
    match,
    mid,
    mod,
    npv,
    right,
    roundup,
    sumif,
    value,
    vlookup,
    xcmp,
    xlog,
    xmax,
    xmin,
    xround,
    xsum,
    yearfrac,
)


def test_numerics():
    assert [1, 2, 3] == _numerics((1, '3', 2.0, pytest, 3))


def test_average():
    assert 2.0 == average((1, '3', 2.0, pytest, 3))

    with pytest.raises(ZeroDivisionError):
        average('3')


class TestCount:

    def test_without_nested_booleans(self):
        assert 3 == count([1, 2, 'e'], True, 'r')

    def test_with_nested_booleans(self):
        assert 2 == count([1, True, 'e'], True, 'r')

    def test_with_text_representations(self):
        assert 4 == count([1, '2.2', 'e'], True, '20')

    #def test_with_date_representations(self):
    #    assert 5 == count([1, '2.2', dt.datetime.now()], True, '20')


class TestCountIf:

    def test_argument_validity(self):
        with pytest.raises(TypeError):
            countif(['e', 1], '>=d')

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

    def test_countif_regular(self):
        assert 2 == countif([7, 25, 13, 25], 25)


class TestCountIfs:
    # more tests might be welcomed

    def test_countifs_regular(self):
        assert 1 == countifs([7, 25, 13, 25], 25, [100, 102, 201, 20], ">100")

    def test_countifs_odd_args_len(self):
        with pytest.raises(Exception):
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

    for error in ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?",
                  "#NUM!", "#N/A", "#GETTING_DATA"):
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


class TestMatch:

    def test_numeric_in_ascending_mode(self):
        # Closest inferior value is found
        assert 3 == match(5, [1, 3.3, 5])

    def test_numeric_in_ascending_mode_with_descending_array(self):
        # Not ascending arrays raise exception
        with pytest.raises(Exception):
            match(3, [10, 9.1, 6.23, 1])

    def test_numeric_in_ascending_mode_with_any_array(self):
        # Not ascending arrays raise exception
        with pytest.raises(Exception):
            match(3, [10, 3.3, 5, 2])

    def test_numeric_in_exact_mode(self):
        # Value is found
        assert 3 == match(5, [10, 3.3, 5.0], 0)

    def test_numeric_in_exact_mode_not_found(self):
        # Value not found raises Exception
        with pytest.raises(ValueError):
            match(3, [10, 3.3, 5, 2], 0)

    def test_numeric_in_descending_mode(self):
        # Closest superior value is found
        assert 2 == match(8, [10, 9.1, 6.2], -1)

    def test_numeric_in_descending_mode_with_ascending_array(self):
        # Non descending arrays raise exception
        with pytest.raises(Exception):
            match(3, [1, 3.3, 5, 6], -1)

    def test_numeric_in_descending_mode_with_any_array(self):
        # Non descending arrays raise exception
        with pytest.raises(Exception):
            match(3, [10, 3.3, 5, 2], -1)

    def test_string_in_ascending_mode(self):
        # Closest inferior value is found
        assert 3 == match('rars', ['a', 'AAB', 'rars'])

    def test_string_in_ascending_mode_with_descending_array(self):
        # Not ascending arrays raise exception
        with pytest.raises(Exception):
            match(3, ['rars', 'aab', 'a'])

    def test_string_in_ascending_mode_with_any_array(self):
        with pytest.raises(Exception):
            match(3, ['aab', 'a', 'rars'])

    def test_string_in_exact_mode(self):
        # Value is found
        assert 2 == match('a', ['aab', 'a', 'rars'], 0)

    def test_string_in_exact_mode_not_found(self):
        # Value not found raises Exception
        with pytest.raises(ValueError):
            match('b', ['aab', 'a', 'rars'], 0)

    def test_string_in_descending_mode(self):
        # Closest superior value is found
        assert 3 == match('a', ['c', 'b', 'a'], -1)

    def test_string_in_descending_mode_with_ascending_array(self):
        # Non descending arrays raise exception
        with pytest.raises(Exception):
            match('a', ['a', 'aab', 'rars'], -1)

    def test_string_in_descending_mode_with_any_array(self):
        # Non descending arrays raise exception
        with pytest.raises(Exception):
            match('a', ['aab', 'a', 'rars'], -1)

    def test_boolean_in_ascending_mode(self):
        # Closest inferior value is found
        assert 3 == match(True, [False, False, True])

    def test_boolean_in_ascending_mode_with_descending_array(self):
        # Not ascending arrays raise exception
        with pytest.raises(Exception):
            match(False, [True, False, False])

    def test_boolean_in_ascending_mode_with_any_array(self):
        # Not ascending arrays raise exception
        with pytest.raises(Exception):
            match(True, [False, True, False])

    def test_boolean_in_exact_mode(self):
        # Value is found
        assert 2 == match(False, [True, False, False], 0)

    def test_boolean_in_exact_mode_not_found(self):
        # Value not found raises Exception
        with pytest.raises(ValueError):
            match(False, [True, True, True], 0)

    def test_boolean_in_descending_mode(self):
        # Closest superior value is found
        assert 3 == match(False, [True, False, False], -1)

    def test_boolean_in_descending_mode_with_ascending_array(self):
        # Non descending arrays raise exception
        with pytest.raises(Exception):
            match(False, [False, False, True], -1)

    def test_boolean_in_descending_mode_with_any_array(self):
        with pytest.raises(Exception):
            match(True, [False, True, False], -1)


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
    pass


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


class TestSumIf:

    def test_range_is_a_list(self):
        with pytest.raises(TypeError):
            sumif(12, 12)

    def test_sum_range_is_a_list(self):
        with pytest.raises(TypeError):
            sumif(12, 12, 12)

    def test_criteria_is_number_string_boolean(self):
        assert 0 == sumif([1, 2, 3], [1, 2])

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


def test_xcmp():
    assert not xcmp(1, 'a')
    assert not xcmp(1, 2)
    assert not xcmp('a', 'b')

    assert xcmp(1, 1)
    assert xcmp('A', 'A')
    assert xcmp('A', 'a')


def test_xlog():
    assert math.log(5) == xlog(5)
    assert [math.log(5), math.log(6)] == xlog([5, 6])
    assert [math.log(5), math.log(6)] == xlog((5, 6))
    assert [math.log(5), math.log(6)] == xlog(np.array([5, 6]))


def test_xmax():
    assert 0 == xmax('abcd')
    assert 3 == xmax((2, None, 'x', 3))


def test_xmin():
    assert 0 == xmin('abcd')
    assert 2 == xmin((2, None, 'x', 3))


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
