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

from pycel.excelutil import (
    ERROR_CODES,
    flatten,
    in_array_formula_context,
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
    elif isinstance(test, (bool, int, float)):
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


def x_and(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   and-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9

    values = _clean_logicals(*args)
    if isinstance(values, str):
        # return error code
        return values
    else:
        return all(values)


@excel_helper(cse_params=(0, 1, 2), err_str_params=0)
def x_if(test, true_value, false_value=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   IF-function-69AED7C9-4E8A-4755-A9BC-AA8BBFF73BE2

    test = _clean_logical(test)

    if isinstance(test, str):
        # return error code
        return test
    else:
        return true_value if test else false_value


def iferror(arg, value_if_error):
    # Excel reference: https://support.office.com/en-us/article/
    #   IFERROR-function-C526FD07-CAEB-47B8-8BB6-63F3E417F611

    if in_array_formula_context and (
            isinstance(arg, tuple) or isinstance(value_if_error, tuple)):
        return cse_array_wrapper(iferror, (0, 1))(arg, value_if_error)
    elif arg in ERROR_CODES or isinstance(arg, tuple):
        return 0 if value_if_error is None else value_if_error
    else:
        return arg


# IFNA function
# Excel 2013
# Returns the value you specify if the expression resolves to #N/A,
# otherwise returns the result of the expression
# Excel reference: https://support.office.com/en-us/article/
#   ifna-function-6626c961-a569-42fc-a49d-79b4951fd461


def ifs(*args):
    # IFS function
    # Excel 2016
    # Checks whether one or more conditions are met and returns a value that
    # corresponds to the first TRUE condition.
    # Excel reference: https://support.office.com/en-us/article/
    #   ifs-function-36329a26-37b2-467c-972b-4a39bd951d45
    if not len(args) % 2:
        for test, value in zip(args[::2], args[1::2]):

            if test in ERROR_CODES:
                return test

            if isinstance(test, str):
                if test.lower() in ('true', 'false'):
                    test = len(test) == 4
                else:
                    return VALUE_ERROR

            elif not isinstance(test, (bool, int, float, type(None))):
                return VALUE_ERROR

            if test:
                return value

    return NA_ERROR


def x_not(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   not-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77

    value = _clean_logical(value)

    if isinstance(value, str):
        # return error code
        return value
    else:
        return not value


def x_or(*args):
    # Excel reference: https://support.office.com/en-us/article/
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
# Excel reference: https://support.office.com/en-us/article/
#   switch-function-47ab33c0-28ce-4530-8a45-d532ec4aa25e


def x_xor(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   xor-function-1548d4c2-5e47-4f77-9a92-0533bba14f37
    values = _clean_logicals(*args)
    if isinstance(values, str):
        # return error code
        return values
    else:
        return sum(bool(v) for v in values) % 2
