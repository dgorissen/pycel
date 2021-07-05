# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of excel logical functions (bools)
"""
from numbers import Number

import numpy as np

from pycel.excelutil import (
    ERROR_CODES,
    flatten,
    has_array_arg,
    in_array_formula_context,
    is_array_arg,
    NA_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import cse_array_wrapper, excel_helper


def _clean_logical(test):
    """For logicals that take one argument, clean via excel rules"""
    if test in ERROR_CODES:
        return test

    if isinstance(test, str):
        if test.lower() in ('true', 'false'):
            test = len(test) == 4
        else:
            return VALUE_ERROR

    if test is None:
        return False
    elif isinstance(test, (Number, np.bool_)):
        return bool(test)
    else:
        return VALUE_ERROR


def _clean_logicals(*args):
    """For logicals that take more than one argument, clean via excel rules"""
    values = tuple(flatten(args))

    error = next((x for x in values if x in ERROR_CODES), None)

    if error is not None:
        # return the first error in the list
        return error
    else:
        values = tuple(x for x in values
                       if not (x is None or isinstance(x, str)))
        return VALUE_ERROR if len(values) == 0 else values


def and_(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   and-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9
    values = _clean_logicals(*args)
    if isinstance(values, str):
        # return error code
        return values
    else:
        return all(values)


# def false(value):
    # A "compatibility function", needed only for use with other spreadsheet programs
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   false-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904


@excel_helper(cse_params=(0, 1, 2), err_str_params=0)
def if_(test, true_value, false_value=0):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   IF-function-69AED7C9-4E8A-4755-A9BC-AA8BBFF73BE2
    cleaned = _clean_logical(test)

    if isinstance(cleaned, str):
        # return error code
        return cleaned
    else:
        return true_value if cleaned else false_value


def iferror(arg, value_if_error):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   IFERROR-function-C526FD07-CAEB-47B8-8BB6-63F3E417F611
    if in_array_formula_context and has_array_arg(arg, value_if_error):
        return cse_array_wrapper(iferror, (0, 1))(arg, value_if_error)
    elif arg in ERROR_CODES or is_array_arg(arg):
        return 0 if value_if_error is None else value_if_error
    else:
        return arg


def ifna(arg, value_if_na):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   ifna-function-6626c961-a569-42fc-a49d-79b4951fd461
    if in_array_formula_context and has_array_arg(arg, value_if_na):
        return cse_array_wrapper(ifna, (0, 1))(arg, value_if_na)
    elif arg == NA_ERROR or is_array_arg(arg):
        return 0 if value_if_na is None else value_if_na
    else:
        return arg


def ifs(*args):
    # IFS function
    # Excel 2016
    # Checks whether one or more conditions are met and returns a value that
    # corresponds to the first TRUE condition.
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   ifs-function-36329a26-37b2-467c-972b-4a39bd951d45
    if not len(args) % 2:
        if in_array_formula_context and any(isinstance(a, tuple) for a in args):
            return cse_array_wrapper(ifs, tuple(range(len(args))))(*args)

        for test, value in zip(args[::2], args[1::2]):

            if test in ERROR_CODES:
                return test

            if isinstance(test, str):
                if test.lower() in ('true', 'false'):
                    test = len(test) == 4
                else:
                    return VALUE_ERROR

            if test:
                return value

    return NA_ERROR


def not_(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   not-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77
    cleaned = _clean_logical(value)

    if isinstance(cleaned, str):
        # return error code
        return cleaned
    else:
        return not cleaned


def or_(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   or-function-7d17ad14-8700-4281-b308-00b131e22af0
    values = _clean_logicals(*args)
    if isinstance(values, str):
        # return error code
        return values
    else:
        return any(values)


# SWITCH function
# Excel 2016
# Evaluates an expression against a list of values and returns the result
# corresponding to the first matching value. If there is no match, an optional
# default value may be returned.
# Excel reference: https://support.microsoft.com/en-us/office/
#   switch-function-47ab33c0-28ce-4530-8a45-d532ec4aa25e


# def true(value):
    # A "compatibility function", needed only for use with other spreadsheet programs
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   true-function-7652c6e3-8987-48d0-97cd-ef223246b3fb


def xor_(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   xor-function-1548d4c2-5e47-4f77-9a92-0533bba14f37
    values = _clean_logicals(*args)
    if isinstance(values, str):
        # return error code
        return values
    else:
        return sum(bool(v) for v in values) % 2


# Older mappings for excel functions that match Python built-in and keywords
x_and = and_
x_if = if_
x_not = not_
x_or = or_
x_xor = xor_
