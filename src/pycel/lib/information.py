# -*- coding: UTF-8 -*-
#
# Copyright 2011-2021 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of Information library functions
"""
import math

from pycel.excelutil import (
    coerce_to_number,
    ERROR_CODES,
    NA_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import excel_helper


# def cell(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf


# def error.type(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   error-type-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa


# def info(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   info-function-725f259a-0e4b-49b3-8b52-58815c69acae


@excel_helper(cse_params=0, err_str_params=None)
def isblank(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return value is None


@excel_helper(cse_params=0, err_str_params=None)
def iserr(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    # Value refers to any error value except #N/A.
    return isinstance(value, str) and value in ERROR_CODES and value != NA_ERROR


@excel_helper(cse_params=0, err_str_params=None)
def iserror(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    # Value refers to any error value:
    #   (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
    return isinstance(value, str) and value in ERROR_CODES or (
        isinstance(value, tuple))


@excel_helper(cse_params=0)
def iseven(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356
    result = isodd(value)
    return not result if isinstance(result, bool) else result


# def isformula(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   isformula-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5


@excel_helper(cse_params=0, err_str_params=None)
def islogical(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return isinstance(value, bool)


@excel_helper(cse_params=0, err_str_params=None)
def isna(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return value == NA_ERROR or isinstance(value, tuple)


@excel_helper(cse_params=0, err_str_params=None)
def isnontext(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return not isinstance(value, str) or value in ERROR_CODES


@excel_helper(cse_params=0, err_str_params=None)
def isnumber(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return not isinstance(value, bool) and isinstance(value, (int, float))


@excel_helper(cse_params=0)
def isodd(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    if isinstance(value, bool):
        return VALUE_ERROR
    value = coerce_to_number(value)
    if isinstance(value, str):
        return VALUE_ERROR
    if value is None:
        value = 0
    return bool(math.floor(abs(value)) % 2)


# def isref(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665


@excel_helper(cse_params=0, err_str_params=None)
def istext(arg):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return isinstance(arg, str) and arg not in ERROR_CODES


@excel_helper(cse_params=0, err_str_params=0)
def n(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   n-function-a624cad1-3635-4208-b54a-29733d1278c9
    if isinstance(value, str):
        return 0
    if isinstance(value, bool):
        return int(value)
    return value


def na():
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   na-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c
    return NA_ERROR


# def sheet(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   sheet-function-44718b6f-8b87-47a1-a9d6-b701c06cff24


# def sheets(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   sheets-function-770515eb-e1e8-45ce-8066-b557e5e4b80b


# def type(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   type-function-45b4e688-4bc3-48b3-a105-ffa892995899
