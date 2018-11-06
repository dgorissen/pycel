"""
Python equivalents of various excel functions
"""

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP
from math import log

import numpy as np

from pycel.excelutil import (
    flatten,
    is_number,
    date_from_int,
    normalize_year,
    is_leap_year,
    find_corresponding_index
)


def _numerics(args):
    # ignore non numeric cells
    return [x for x in flatten(args) if isinstance(x, (int, float))]


def average(*args):
    args = _numerics(args)
    return sum(args) / len(args)


def count(*args):
    # Excel reference: https://support.office.com/en-us/article/
    # COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

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
        raise Exception('excellib.countifs() must have a '
                        'pair number of arguments, here %d' % len(args))

    if len(args):
        indexes = find_corresponding_index(args[0], args[
            1])  # find indexes that match first layer of countif

        remaining_ranges = [elem for i, elem in enumerate(args[2:]) if
                            i % 2 == 0]  # get only ranges
        remaining_criteria = [elem for i, elem in enumerate(args[2:]) if
                              i % 2 == 1]  # get only criteria

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


def isNa(arg):
    # This function might need more solid testing
    try:
        eval(arg)
        return False
    except:
        return True


def linest(Y, X, const=True, degree=1):
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


def lookup(arg, lookup_range, result_range):
    # TODO
    if not isinstance(arg, (int, float)):
        raise Exception("Non numeric lookups (%s) not supported" % arg)

    # TODO: note, may return the last equal value

    # index of the last numeric value
    lastnum = -1
    for i, v in enumerate(lookup_range):
        if isinstance(v, (int, float)):
            if v > arg:
                break
            else:
                lastnum = i

    if lastnum < 0:
        raise Exception("No numeric data found in the lookup range")
    else:
        if i == 0:
            raise Exception(
                "All values in the lookup range are bigger than %s" % arg)
        else:
            if i >= len(lookup_range) - 1:
                # return the biggest number smaller than value
                return result_range[lastnum]
            else:
                return result_range[i - 1]


def match(lookup_value, lookup_array, match_type=1):
    def type_convert(arg):
        if type(arg) == str:
            arg = arg.lower()
        elif type(arg) == int:
            arg = float(arg)

        return arg

    lookup_value = type_convert(lookup_value)

    if match_type == 1:
        # Verify ascending sort
        pos_max = -1
        for i in range((len(lookup_array))):
            current = type_convert(lookup_array[i])

            if i is not len(lookup_array) - 1 and current > type_convert(
                    lookup_array[i + 1]):
                raise Exception(
                    'for match_type 0, lookup_array must be sorted ascending')
            if current <= lookup_value:
                pos_max = i
        if pos_max == -1:
            raise Exception('no result in lookup_array for match_type 0')
        return pos_max + 1  # Excel starts at 1

    elif match_type == 0:
        # No string wildcard
        return [type_convert(x) for x in lookup_array].index(lookup_value) + 1

    elif match_type == -1:
        # Verify descending sort
        pos_min = -1
        for i in range((len(lookup_array))):
            current = type_convert(lookup_array[i])

            if i is not len(lookup_array) - 1 and current < type_convert(
                    lookup_array[i + 1]):
                raise (
                    'for match_type 0, lookup_array must be sorted descending')
            if current >= lookup_value:
                pos_min = i
        if pos_min == -1:
            raise Exception('no result in lookup_array for match_type 0')
        return pos_min + 1  # Excel starts at 1


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
    # Excel refeence: https://support.office.com/en-us/article/
    #   ROUNDUP-function-F8BC9B23-E795-47DB-8703-DB171D0C42A7
    if not is_number(number):
        raise TypeError("%s is not a number" % str(number))
    if not is_number(num_digits):
        raise TypeError("%s is not a number" % str(num_digits))

    quant = Decimal('1E{}{}'.format('+-'[num_digits >= 0], abs(num_digits)))
    return float(Decimal(repr(number)).quantize(quant, rounding=ROUND_UP))


def sumif(range, criteria, sum_range=[]):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b

    # WARNING:
    # - wildcards not supported
    # - doesn't really follow 2nd remark about sum_range length

    if type(range) != list:
        raise TypeError('%s must be a list' % str(range))

    if type(sum_range) != list:
        raise TypeError('%s must be a list' % str(sum_range))

    if isinstance(criteria, list) and not isinstance(criteria, (str, bool)):
        # ugly...
        return 0

    indexes = find_corresponding_index(range, criteria)

    def f(x):
        return sum_range[x] if x < len(sum_range) else 0

    if len(sum_range) == 0:
        return sum(range[x] for x in indexes)
    else:
        return sum(map(f, indexes))


def value(text):
    # make the distinction for naca numbers
    if '.' in text:
        return float(text)
    else:
        return int(text)


def xcmp(a, b):
    if isinstance(a, str) and isinstance(b, str):
        return a.lower() == b.lower()
    return a == b


def xlog(a):
    if isinstance(a, (list, tuple, np.ndarray)):
        return [log(x) for x in flatten(a)]
    else:
        return log(a)


def xmax(*args):
    data = _numerics(args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return max(data)


def xmin(*args):
    data = _numerics(args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
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
    data = _numerics(args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
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
