# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of various excel functions
"""
import math
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP
from heapq import nlargest, nsmallest

import numpy as np

from pycel.excelutil import (
    coerce_to_number,
    DIV0,
    ERROR_CODES,
    find_corresponding_index,
    flatten,
    handle_ifs,
    list_like,
    NA_ERROR,
    NUM_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import (
    excel_helper,
    excel_math_func,
)


def _numerics(*args, keep_bools=False, to_number=lambda x: x):
    # ignore non numeric cells
    args = tuple(flatten(args))
    error = next((x for x in args if x in ERROR_CODES), None)
    if error is not None:
        # return the first error in the list
        return error
    else:
        args = (
            to_number(a) for a in args if keep_bools or not isinstance(a, bool)
        )
        return tuple(x for x in args if isinstance(x, (int, float)))


def average(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   average-function-047bac88-d466-426c-a32b-8f33eb960cf6
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data
    elif len(data) == 0:
        return DIV0
    else:
        return sum(data) / len(data)


def averageif(rng, criteria, average_range=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642

    # WARNING:
    # - The following is not currently implemented:
    #  The average_range argument does not have to be the same size and shape
    #  as the range argument. The actual cells that are added are determined by
    #  using the upper leftmost cell in the average_range argument as the
    #  beginning cell, and then including cells that correspond in size and
    #  shape to the range argument.
    if average_range is None:
        average_range = rng
    return averageifs(average_range, rng, criteria)


def averageifs(average_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   AVERAGEIFS-function-48910C45-1FC0-4389-A028-F7C5C3001690
    if not list_like(average_range):
        average_range = ((average_range, ), )

    coords = handle_ifs(args, average_range)

    # A returned string is an error code
    if isinstance(coords, str):
        return coords

    data = _numerics((average_range[r][c] for r, c in coords), keep_bools=True)
    if len(data) == 0:
        return DIV0
    return sum(data) / len(data)


@excel_math_func
def ceiling(number, significance):
    # Excel reference: https://support.office.com/en-us/article/
    #   CEILING-function-0A5CD7C8-0720-4F0A-BD2C-C943E510899F
    if significance < 0 < number:
        return NUM_ERROR

    if number == 0 or significance == 0:
        return 0

    if number < 0 < significance:
        return significance * int(number / significance)
    else:
        return significance * math.ceil(number / significance)


@excel_math_func
def ceiling_math(number, significance=1, mode=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   ceiling-math-function-80f95d2f-b499-4eee-9f16-f795a8e306c8
    if significance == 0:
        return 0

    significance = abs(significance)
    if mode and number < 0:
        significance = -significance
    return significance * math.ceil(number / significance)


@excel_math_func
def ceiling_precise(number, significance=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   ceiling-precise-function-f366a774-527a-4c92-ba49-af0a196e66cb
    if significance == 0:
        return 0

    significance = abs(significance)
    return significance * math.ceil(number / significance)


def count(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

    return sum(1 for x in flatten(args)
               if isinstance(x, (int, float)) and not isinstance(x, bool))


def countif(rng, criteria):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34
    if not list_like(rng):
        rng = ((rng, ), )
    valid = find_corresponding_index(rng, criteria)
    return len(valid)


def conditional_format_ids(*args):
    """helper function for getting conditional format ids"""
    # Excel reference: https://support.office.com/en-us/article/
    #   E09711A3-48DF-4BCB-B82C-9D8B8B22463D

    results = []
    for condition, dxf_id, stop_if_true in args:
        if condition:
            results.append(dxf_id)
            if stop_if_true:
                break
    return results


def countifs(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842
    coords = handle_ifs(args)

    # A returned string is an error code
    if isinstance(coords, str):
        return coords

    return len(coords)


@excel_math_func
def even(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   even-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9
    return math.copysign(math.ceil(abs(value) / 2) * 2, value)


@excel_math_func
def fact(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   fact-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3
    return math.factorial(int(value)) if value >= 0 else NUM_ERROR


@excel_helper(cse_params=-1)
def factdouble(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   fact-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3
    if isinstance(value, bool):
        return VALUE_ERROR
    value = coerce_to_number(value, convert_all=True)
    if isinstance(value, str):
        return VALUE_ERROR
    if value < 0:
        return NUM_ERROR

    return np.sum(np.prod(range(int(value), 0, -2), axis=0))


@excel_math_func
def floor(number, significance):
    # Excel reference: https://support.office.com/en-us/article/
    #   FLOOR-function-14BB497C-24F2-4E04-B327-B0B4DE5A8886
    if significance < 0 < number:
        return NUM_ERROR

    if number == 0:
        return 0

    if significance == 0:
        return DIV0

    return significance * math.floor(number / significance)


@excel_math_func
def floor_math(number, significance=1, mode=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   floor-math-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5
    if significance == 0:
        return 0

    significance = abs(significance)
    if mode and number < 0:
        significance = -significance
    return significance * math.floor(number / significance)


@excel_math_func
def floor_precise(number, significance=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   floor-precise-function-f769b468-1452-4617-8dc3-02f842a0702e
    if significance == 0:
        return 0

    significance = abs(significance)
    return significance * math.floor(number / significance)


@excel_helper(cse_params=0, err_str_params=None)
def iserr(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    # Value refers to any error value except #N/A.
    return isinstance(value, str) and value in ERROR_CODES and value != NA_ERROR


@excel_helper(cse_params=0, err_str_params=None)
def iserror(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    # Value refers to any error value:
    #   (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
    return isinstance(value, str) and value in ERROR_CODES or (
        isinstance(value, tuple))


@excel_helper(cse_params=0)
def iseven(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356
    result = isodd(value)
    return not result if isinstance(result, bool) else result


@excel_helper(cse_params=0, err_str_params=None)
def istext(arg):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return isinstance(arg, str) and arg not in ERROR_CODES


@excel_helper(cse_params=0, err_str_params=None)
def isna(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return value == NA_ERROR or isinstance(value, tuple)


@excel_helper(cse_params=0)
def isodd(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    if isinstance(value, bool):
        return VALUE_ERROR
    value = coerce_to_number(value)
    if isinstance(value, str):
        return VALUE_ERROR
    return bool(math.floor(abs(value)) % 2)


@excel_helper(cse_params=0, err_str_params=None)
def isnumber(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return isinstance(value, (int, float))


@excel_helper()
def large(array, k):
    # Excel reference: https://support.office.com/en-us/article/
    #   large-function-3af0af19-1190-42bb-bb8b-01672ec00a64
    data = _numerics(array, to_number=coerce_to_number)
    if isinstance(data, str):
        return data

    k = coerce_to_number(k)
    if isinstance(k, str):
        return VALUE_ERROR

    if not data or k is None or k < 1 or k > len(data):
        return NUM_ERROR

    k = math.ceil(k)
    return nlargest(k, data)[-1]


def linest(Y, X, const=True, degree=1):  # pragma: no cover  ::TODO::
    if isinstance(const, str):
        const = (const.lower() == "true")

    def assert_vector(data):
        vector = np.array(data)
        assert 1 in vector.shape
        return vector.ravel()

    X = assert_vector(X)
    Y = assert_vector(Y)

    # build the vandermonde matrix
    A = np.vander(X, degree + 1)

    if not const:
        # force the intercept to zero
        A[:, -1] = np.zeros((1, len(X)))

    # perform the fit
    coefs, residuals, rank, sing_vals = np.linalg.lstsq(A, Y, rcond=None)

    return coefs


@excel_math_func
def ln(arg):
    # Excel reference: https://support.office.com/en-us/article/
    #   LN-function-81FE1ED7-DAC9-4ACD-BA1D-07A142C6118F
    return math.log(arg)


@excel_math_func
def log(number, base=10):
    # Excel reference: https://support.office.com/en-us/article/
    #   LOG-function-4E82F196-1CA9-4747-8FB0-6C4A3ABB3280
    return math.log(number, base)


def maxifs(max_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883
    if not list_like(max_range):
        max_range = ((max_range, ), )

    try:
        coords = handle_ifs(args, max_range)

        # A returned string is an error code
        if isinstance(coords, str):
            return coords

        return max(_numerics(
            (max_range[r][c] for r, c in coords),
            keep_bools=True
        ))
    except ValueError:
        return 0


def minifs(min_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599
    if not list_like(min_range):
        min_range = ((min_range, ), )

    try:
        coords = handle_ifs(args, min_range)

        # A returned string is an error code
        if isinstance(coords, str):
            return coords

        return min(_numerics(
            (min_range[r][c] for r, c in coords),
            keep_bools=True
        ))
    except ValueError:
        return 0


@excel_math_func
def mod(number, divisor):
    # Excel reference: https://support.office.com/en-us/article/
    #   MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    if divisor == 0:
        return DIV0

    return number % divisor

def get_numeric(value):
    """Return True if the argument is a valid number, return False otherwise."""

    # Test if boolean type
    if np.dtype(type(value)) == bool:
        return False
    else:
        try:
            # Ignore empty values
            if np.isnan(value):
                return False

            # Check if value is a float
            float(value)
            return True

        # If you can't convert to float its not a number
        except:
            return False

@excel_math_func
def npv(rate, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   NPV-function-8672CB67-2576-4D07-B67B-AC28ACF2A568

    if rate in ERROR_CODES:
        return rate

    # Check if rate is a valid number
    try:
        float(rate)
    except:
        return VALUE_ERROR

    _rate = rate + 1

    cashflows = [x for x in flatten(args[0])]

    # Return the correct error code if one of the cash flows is invalid
    for cashflow in cashflows:
        if cashflow in ERROR_CODES:
            return cashflow

    # For entries that are both non-numeric and non-error, Excel removes them
    # and does not treat as zero or raise an error
    fil = [get_numeric(c) for c in cashflows]
    cashflows = np.array([i for (i, v) in zip(cashflows, fil) if v])

    return (cashflows/np.power(_rate, np.arange(1, len(cashflows) + 1))).sum()


@excel_math_func
def odd(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   odd-function-deae64eb-e08a-4c88-8b40-6d0b42575c98
    return math.copysign(math.ceil((abs(value) - 1) / 2) * 2 + 1, value)


@excel_math_func
def power(number, power):
    # Excel reference: https://support.office.com/en-us/article/
    #   POWER-function-D3F2908B-56F4-4C3F-895A-07FB519C362A
    if number == power == 0:
        # Really excel?  What were you thinking?
        return NA_ERROR

    try:
        return number ** power
    except ZeroDivisionError:
        return DIV0


def _round(number, num_digits, rounding):
    num_digits = int(num_digits)
    quant = Decimal('1E{}{}'.format('+-'[num_digits >= 0], abs(num_digits)))
    return float(Decimal(repr(number)).quantize(quant, rounding=rounding))


@excel_math_func
def rounddown(number, num_digits):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUNDDOWN-function-2EC94C73-241F-4B01-8C6F-17E6D7968F53
    return _round(number, num_digits, rounding=ROUND_DOWN)


@excel_math_func
def roundup(number, num_digits):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUNDUP-function-F8BC9B23-E795-47DB-8703-DB171D0C42A7
    return _round(number, num_digits, rounding=ROUND_UP)


@excel_math_func
def sign(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   sign-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8
    return -1 if value < 0 else int(bool(value))


@excel_helper()
def small(array, k):
    # Excel reference: https://support.office.com/en-us/article/
    #   small-function-17da8222-7c82-42b2-961b-14c45384df07
    data = _numerics(array, to_number=coerce_to_number)
    if isinstance(data, str):
        return data

    k = coerce_to_number(k)
    if isinstance(k, str):
        return VALUE_ERROR

    if not data or k is None or k < 1 or k > len(data):
        return NUM_ERROR

    k = math.ceil(k)
    return nsmallest(k, data)[-1]


def sumif(rng, criteria, sum_range=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b

    # WARNING:
    # - The following is not currently implemented:
    #  The sum_range argument does not have to be the same size and shape as
    #  the range argument. The actual cells that are added are determined by
    #  using the upper leftmost cell in the sum_range argument as the
    #  beginning cell, and then including cells that correspond in size and
    #  shape to the range argument.

    if sum_range is None:
        sum_range = rng
    return sumifs(sum_range, rng, criteria)


def sumifs(sum_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIFS-function-C9E748F5-7EA7-455D-9406-611CEBCE642B
    if not list_like(sum_range):
        sum_range = ((sum_range, ), )

    coords = handle_ifs(args, sum_range)

    # A returned string is an error code
    if isinstance(coords, str):
        return coords

    return sum(_numerics(
        (sum_range[r][c] for r, c in coords),
        keep_bools=True
    ))


def sumproduct(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMPRODUCT-function-16753E75-9F68-4874-94AC-4D2145A2FD2E

    # find any errors
    error = next((i for i in flatten(args) if i in ERROR_CODES), None)
    if error:
        return error

    # verify array sizes match
    sizes = set()
    for arg in args:
        assert isinstance(arg, tuple), isinstance(arg[0], tuple)
        sizes.add((len(arg), len(arg[0])))
    if len(sizes) != 1:
        return VALUE_ERROR

    # put the values into numpy vectors
    values = np.array(tuple(tuple(
        x if isinstance(x, (float, int)) and not isinstance(x, bool) else 0
        for x in flatten(arg)) for arg in args))

    # return the sum product
    return np.sum(np.prod(values, axis=0))


@excel_math_func
def trunc(number, num_digits=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   TRUNC-function-8B86A64C-3127-43DB-BA14-AA5CEB292721
    factor = 10 ** int(num_digits)
    return int(number * factor) / factor


@excel_math_func
def x_abs(value1):
    # Excel reference: https://support.office.com/en-us/article/
    #   ABS-function-3420200F-5628-4E8C-99DA-C99D7C87713C
    return abs(value1)


@excel_math_func
def xatan2(x_num, y_num):
    # Excel reference: https://support.office.com/en-us/article/
    #   ATAN2-function-C04592AB-B9E3-4908-B428-C96B3A565033

    # swap arguments
    return math.atan2(y_num, x_num)


@excel_math_func
def x_int(value1):
    # Excel reference: https://support.office.com/en-us/article/
    #   INT-function-A6C4AF9E-356D-4369-AB6A-CB1FD9D343EF
    return math.floor(value1)


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


@excel_math_func
def x_round(number, num_digits=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

    num_digits = int(num_digits)
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
