"""
Python equivalents of various excel functions
"""
from bisect import bisect_right
import itertools as it
from collections import Counter
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP
from math import log

import numpy as np
from pycel.excelutil import (
    assert_list_like,
    build_wildcard_re,
    coerce_to_number,
    date_from_int,
    DIV0,
    ERROR_CODES,
    ExcelCmp,
    find_corresponding_index,
    flatten,
    is_leap_year,
    is_number,
    list_like,
    NA_ERROR,
    normalize_year,
    PyCelException,
)


def _numerics(*args):
    # ignore non numeric cells
    args = tuple(flatten(args, lambda x: coerce_to_number(x, raise_div0=False)))
    error = next((x for x in args if x in ERROR_CODES), None)
    if error is not None:
        # return the first error in the list
        return error
    else:
        return tuple(x for x in args if isinstance(x, (int, float)))


def average(*args):
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data
    elif len(data) == 0:
        return DIV0
    else:
        return sum(data) / len(data)


def column(ref):
    if ref.is_range:
        ref = ref.start
    return max(ref.col_idx, 1)


def count(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

    total = 0

    for arg in args:
        if isinstance(arg, list):
            # count inside a list
            total += len(
                [x for x in arg if is_number(x) and not isinstance(x, bool)])
        else:
            total += int(is_number(arg))

    return total


def countif(range, criteria):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34

    # WARNING:
    # - wildcards not supported
    # - support of strings with >, <, <=, =>, <> not provided

    valid = find_corresponding_index(range, criteria)

    return len(valid)


def countifs(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842

    if len(args) % 2 != 0:
        raise PyCelException('excellib.countifs() must have a '
                             'pair number of arguments, here %d' % len(args))

    if len(args):
        # find indexes that match first layer of countif
        indexes = find_corresponding_index(args[0], args[1])

        # get only ranges
        remaining_ranges = [elem for i, elem in enumerate(args[2:])
                            if i % 2 == 0]

        # get only criteria
        remaining_criteria = [elem for i, elem in enumerate(args[2:])
                              if i % 2 == 1]

        filtered_remaining_ranges = []

        # filter items in remaining_ranges that match valid indexes
        # from first countif layer
        for rng in remaining_ranges:
            filtered_remaining_range = []

            for index, item in enumerate(rng):
                if index in indexes:
                    filtered_remaining_range.append(item)

            filtered_remaining_ranges.append(filtered_remaining_range)

        new_tuple = ()

        # rebuild the tuple that will be the argument of next layer
        for index, rng in enumerate(filtered_remaining_ranges):
            new_tuple += (rng, remaining_criteria[index])

        # only consider the minimum number across all layer responses
        return min(countifs(*new_tuple), len(indexes))

    else:
        return float('inf')


def date(year, month, day):
    # Excel reference: https://support.office.com/en-us/article/
    #   DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

    if not isinstance(year, int):
        raise TypeError("%s is not an integer" % year)

    if not isinstance(month, int):
        raise TypeError("%s is not an integer" % month)

    if not isinstance(day, int):
        raise TypeError("%s is not an integer" % day)

    if not (0 <= year <= 9999):
        raise ValueError("Year '%s' must be between 1 and 9999" % year)

    if year < 1900:
        year += 1900

    # taking into account negative month and day values
    year, month, day = normalize_year(year, month, day)

    date_0 = datetime(1900, 1, 1)
    result = (datetime(year, month, day) - date_0).days + 2

    if result <= 0:
        raise ArithmeticError("Date result is negative")
    return result


def iferror(arg, value_if_error):
    # Excel reference: https://support.office.com/en-us/article/
    #   IFERROR-function-C526FD07-CAEB-47B8-8BB6-63F3E417F611

    return value_if_error if arg in ERROR_CODES else arg


def index(array, row_num, col_num=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd

    if isinstance(array[0], (list, tuple, np.ndarray)):
        # rectangular array
        if None not in (row_num, col_num):
            return array[row_num - 1][col_num - 1]

        if row_num is not None:
            return array[row_num - 1]

        if col_num is not None:
            if isinstance(array, np.ndarray):
                return array[:, col_num - 1]
            else:
                return type(array)(row[col_num - 1] for row in array)

        raise IndexError("For Index either row_num or col_num must be given")

    elif col_num in (1, None):
        return array[row_num - 1]

    elif row_num == 1:
        return array[col_num - 1]

    raise IndexError("Index (%s,%s) out of range for %s" % (
        row_num, col_num, array))


def istext(arg):
    return isinstance(arg, str)


def isNa(arg):
    # This function might need more solid testing
    try:
        eval(arg)
        return False
    except Exception:
        return True


def linest(Y, X, const=True, degree=1):  # pragma: no cover  ::TODO::
    if isinstance(const, str):
        const = (const.lower() == "true")

    # build the vandermonde matrix
    A = np.vander(X, degree + 1)

    if not const:
        # force the intercept to zero
        A[:, -1] = np.zeros((1, len(X)))

    # perform the fit
    coefs, residuals, rank, sing_vals = np.linalg.lstsq(A, Y, rcond=None)

    return coefs


def lookup(lookup_value, lookup_array, result_range=None):
    """
    There are two ways to use LOOKUP: Vector form and Array form

    Vector form: lookup_array is list like (ie: n x 1)

    Array form: lookup_array is rectangular (ie: n x m)

        First row or column is the lookup vector.
        Last row or column is the result vector
        The longer dimension is the search dimension

    :param lookup_value: value to match (value or cell reference)
    :param lookup_array: range of cells being searched.
    :param result_range: (optional vector form) values are returned from here
    :return: #N/A if not found else value
    """
    if not list_like(lookup_array):
        return NA_ERROR

    height = len(lookup_array)

    if list_like(lookup_array[0]):
        # rectangular array
        assert result_range is None
        width = len(lookup_array[0])

        # match across the largest dimension
        if width <= height:
            match_idx = match(lookup_value, tuple(i[0] for i in lookup_array))
            result_range = tuple(i[-1] for i in lookup_array)
        else:
            match_idx = match(lookup_value, lookup_array[0])
            result_range = lookup_array[-1]
    else:
        match_idx = match(lookup_value, lookup_array)
        result_range = result_range or lookup_array

    if isinstance(match_idx, int):
        return result_range[match_idx - 1]
    else:
        # error string
        return match_idx


def match(lookup_value, lookup_array, match_type=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   MATCH-function-E8DFFD45-C762-47D6-BF89-533F4A37673A

    """ The relative position of a specified item in a range of cells.

    Match_type Behavior

    1: return the largest value that is less than or equal to
    `lookup_value`. `lookup_array` must be in ascending order.

    0: return the first value that is exactly equal to lookup_value.
    `lookup_array` can be in any order.

    -1: return the smallest value that is greater than or equal to
    `lookup_value`. `lookup_array` must be in descending order.

    If `match_type` is 0 and lookup_value is a text string, you can use the
    wildcard characters — the question mark (?) and asterisk (*).

    :param lookup_value: value to match (value or cell reference)
    :param lookup_array: range of cells being searched.
    :param match_type: The number -1, 0, or 1.
    :return: #N/A if not found, or relative position in `lookup_array`
    """
    if lookup_value in ERROR_CODES:
        return lookup_value

    lookup_value = ExcelCmp(lookup_value)

    if match_type == 1:
        # Use a binary search to speed it up.  Excel seems to do this as it
        # would explain the results seen when doing out of order searches.
        lookup_value = ExcelCmp(lookup_value)

        result = bisect_right(lookup_array, lookup_value)
        while result and lookup_value.cmp_type != ExcelCmp(
                lookup_array[result - 1]).cmp_type:
            result -= 1
        if result == 0:
            result = NA_ERROR
        return result

    result = [NA_ERROR]

    if match_type == 0:
        def compare(idx, val):
            if val == lookup_value:
                result[0] = idx
                return True

        if lookup_value.cmp_type == 1:
            # string matches might be wildcards
            re_compare = build_wildcard_re(lookup_value.value)
            if re_compare is not None:
                def compare(idx, val):  # noqa: F811
                    if re_compare(val.value):
                        result[0] = idx
                        return True
    else:
        def compare(idx, val):
            if val < lookup_value:
                return True
            result[0] = idx
            return val == lookup_value

    for i, value in enumerate(lookup_array, 1):
        if value not in ERROR_CODES:
            value = ExcelCmp(value)
            if value.cmp_type == lookup_value.cmp_type and compare(i, value):
                break

    return result[0]


def mid(text, start_num, num_chars):
    # Excel reference: https://support.office.com/en-us/article/
    #   MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028

    text = str(text)

    if type(start_num) != int:
        raise TypeError("%s is not an integer" % str(start_num))
    if type(num_chars) != int:
        raise TypeError("%s is not an integer" % str(num_chars))

    if start_num < 1:
        raise ValueError("%s is < 1" % str(start_num))
    if num_chars < 0:
        raise ValueError("%s is < 0" % str(num_chars))

    return text[start_num:num_chars]


def mod(nb, q):
    # Excel reference: https://support.office.com/en-us/article/
    #   MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    if not isinstance(nb, int):
        raise TypeError("%s is not an integer" % str(nb))
    elif not isinstance(q, int):
        raise TypeError("%s is not an integer" % str(q))
    else:
        return nb % q


def npv(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   NPV-function-8672CB67-2576-4D07-B67B-AC28ACF2A568

    rate = args[0] + 1
    cashflow = args[1:]
    return sum([float(x) * rate ** -i for i, x in enumerate(cashflow, start=1)])


def right(text, n):
    if n < 0:
        raise ValueError("Offset for right must not be negative")
    if n == 0:
        return ''
    if not isinstance(text, str):
        # TODO: hack to deal with naca section numbers
        text = str(int(text))

    return text[-n:]


def roundup(number, num_digits):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUNDUP-function-F8BC9B23-E795-47DB-8703-DB171D0C42A7
    if not is_number(number):
        raise TypeError("%s is not a number" % str(number))
    if not is_number(num_digits):
        raise TypeError("%s is not a number" % str(num_digits))

    quant = Decimal('1E{}{}'.format('+-'[num_digits >= 0], abs(num_digits)))
    return float(Decimal(repr(number)).quantize(quant, rounding=ROUND_UP))


def row(ref):
    if ref.is_range:
        ref = ref.start
    return max(ref.row, 1)


def sumif(rng, criteria, sum_range=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b

    if sum_range is None:
        sum_range = rng
    return sumifs(sum_range, rng, criteria)


def sumifs(sum_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIFS-function-C9E748F5-7EA7-455D-9406-611CEBCE642B

    # WARNING:
    # - The following is not currently implemented:
    #  The sum_range argument does not have to be the same size and shape as
    #  the range argument. The actual cells that are added are determined by
    #  using the upper leftmost cell in the sum_range argument as the
    #  beginning cell, and then including cells that correspond in size and
    #  shape to the range argument.

    assert_list_like(sum_range)

    assert len(args) and len(args) % 2 == 0, \
        'Must have paired criteria and ranges'

    # count the number of times a particular cell matches the criteria
    index_counts = Counter(it.chain.from_iterable(
        find_corresponding_index(rng, criteria)
        for rng, criteria in zip(args[0::2], args[1::2])))

    ifs_count = len(args) // 2
    max_idx = len(sum_range)
    indices = tuple(idx for idx, cnt in index_counts.items()
                    if cnt == ifs_count and idx < max_idx)
    return sum(_numerics(sum_range[idx] for idx in indices))


def value(text):
    # make the distinction for naca numbers
    if '.' in text:
        return float(text)
    else:
        return int(text)


def vlookup(lookup_value, table_array, col_index_num, range_lookup=False):
    # Excel reference: https://support.office.com/en-us/article/
    #   VLOOKUP-function-0BBC8083-26FE-4963-8AB8-93A18AD188A1

    assert not range_lookup, "range_lookup not implemented in vlookup"

    if col_index_num <= 0:
        return '#VALUE!'

    if col_index_num > len(table_array[0]):
        return '#REF!'

    result_idx = match(
        lookup_value,
        [row[0] for row in table_array],
        match_type=0
    )

    if result_idx == NA_ERROR:
        return NA_ERROR
    else:
        return table_array[result_idx - 1][col_index_num - 1]


def xlog(a):
    if list_like(a):
        return [log(x) for x in flatten(a)]
    else:
        return log(a)


def xmax(*args):
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data

    # however, if no non numeric cells, return zero (is what excel does)
    elif len(data) < 1:
        return 0
    else:
        return max(data)


def xmin(*args):
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data

    # however, if no non numeric cells, return zero (is what excel does)
    elif len(data) < 1:
        return 0
    else:
        return min(data)


def xround(number, num_digits=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

    if not is_number(number):
        raise TypeError("%s is not a number" % str(number))
    if not is_number(num_digits):
        raise TypeError("%s is not a number" % str(num_digits))

    if num_digits >= 0:  # round to the right side of the point
        return float(Decimal(repr(number)).quantize(
            Decimal(repr(pow(10, -num_digits))),
            rounding=ROUND_HALF_UP
        ))
        # see https://docs.python.org/2/library/functions.html#round
        # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5

    else:
        return round(number, num_digits)


def xsum(*args):
    data = _numerics(*args)
    if isinstance(data, str):
        return data

    # if no non numeric cells, return zero (is what excel does)
    return sum(data)


def yearfrac(start_date, end_date, basis=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8

    def actual_nb_days_afb_alter(beg, end):
        # http://svn.finmath.net/finmath%20lib/trunk/src/main/java/net/
        #   finmath/time/daycount/DayCountConvention_ACT_ACT_YEARFRAC.java
        delta = date(*end) - date(*beg)

        if delta <= 365:
            if (is_leap_year(beg[0]) and date(*beg) <= date(beg[0], 2, 29) or
                is_leap_year(end[0]) and date(*end) >= date(end[0], 2, 29) or
                    is_leap_year(beg[0]) and is_leap_year(end[0])):
                denom = 366
            else:
                denom = 365
        else:
            year_range = range(beg[0], end[0] + 1)
            nb = 0

            for y in year_range:
                nb += 366 if is_leap_year(y) else 365

            denom = nb / len(year_range)

        return delta / denom

    if not is_number(start_date):
        raise TypeError("start_date %s must be a number" % str(start_date))
    if not is_number(end_date):
        raise TypeError("end_date %s must be number" % str(end_date))
    if start_date < 0:
        raise ValueError("start_date %s must be positive" % str(start_date))
    if end_date < 0:
        raise ValueError("end_date %s must be positive" % str(end_date))

    if start_date > end_date:  # switch dates if start_date > end_date
        start_date, end_date = end_date, start_date

    y1, m1, d1 = date_from_int(start_date)
    y2, m2, d2 = date_from_int(end_date)

    if basis == 0:  # US 30/360
        d1 = min(d1, 30)
        d2 = max(d2, 30) if d1 == 30 else d2

        day_count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = day_count / 360

    elif basis == 1:  # Actual/actual
        result = actual_nb_days_afb_alter((y1, m1, d1), (y2, m2, d2))

    elif basis == 2:  # Actual/360
        result = (end_date - start_date) / 360

    elif basis == 3:  # Actual/365
        result = (end_date - start_date) / 365

    elif basis == 4:  # Eurobond 30/360
        d2 = min(d2, 30)
        d1 = min(d1, 30)

        day_count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = day_count / 360

    else:
        raise ValueError("basis: %d must be 0, 1, 2, 3 or 4" % basis)

    return result
