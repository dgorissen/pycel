# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

import pycel.lib.logical
from pycel.excelutil import (
    DIV0,
    in_array_formula_context,
    NA_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import load_to_test_module
from pycel.lib.logical import (
    _clean_logicals,
    iferror,
    ifs,
    x_and,
    x_if,
    x_not,
    x_or,
    x_xor,
)


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.logical, __name__)


@pytest.mark.parametrize(
    'test_value, result', (
        ((1, '3', 2.0, 3.1, ('x', True, None)),
         (1, 2, 3.1, True)),
        ((1, '3', 2.0, 3.1, ('x', VALUE_ERROR)),
         VALUE_ERROR),
        ((1, NA_ERROR, 2.0, 3.1, ('x', VALUE_ERROR)),
         NA_ERROR),
        (('1', ('x', 'y')),
         VALUE_ERROR),
        ((),
         VALUE_ERROR),
    )
)
def test_clean_logicals(test_value, result):
    assert _clean_logicals(*test_value) == result


@pytest.mark.parametrize(
    'result, test_value', (
        (False, (False,)),
        (True, (True,)),
        (True, (1, '3', 2.0, 3.1, ('x', True))),
        (True, (1, '3', 2.0, 3.1, ('x', True, None))),
        (False, (1, '3', 2.0, 3.1, ('x', False))),
        (VALUE_ERROR, (1, '3', 2.0, 3.1, ('x', VALUE_ERROR))),
        (NA_ERROR, (1, NA_ERROR, 2.0, 3.1, ('x', VALUE_ERROR))),
        (VALUE_ERROR, ('1', ('x', 'y'))),
        (VALUE_ERROR, (),),
        (NA_ERROR, (NA_ERROR, 1),),
    )
)
def test_x_and(result, test_value):
    assert x_and(*test_value) == result


@pytest.mark.parametrize(
    'test_value, error_value, result', (
        ('A', 2, 'A'),
        ('#NULL!', 2, 2),
        ('#DIV/0!', 2, 2),
        ('#VALUE!', 2, 2),
        ('#REF!', 2, 2),
        ('#NAME?', 2, 2),
        ('#NUM!', 2, 2),
        ('#N/A', 2, 2),
        ('#N/A', None, 0),
        (((1, VALUE_ERROR), (VALUE_ERROR, 1)), 2, ((1, 2), (2, 1))),
        (((1, VALUE_ERROR), (VALUE_ERROR, 1)), None, ((1, 0), (0, 1))),
    )
)
def test_iferror(test_value, error_value, result):
    if isinstance(test_value, tuple):
        with in_array_formula_context('A1'):
            assert iferror(test_value, error_value) == result
        result = 0 if error_value is None else error_value

    assert iferror(test_value, error_value) == result


@pytest.mark.parametrize(
    'test_value, true_value, false_value, result', (
        ('xyzzy', 3, 2, VALUE_ERROR),
        ('0', 2, 1, VALUE_ERROR),
        (True, 2, 1, 2),
        (False, 2, 1, 1),
        ('True', 2, 1, 2),
        ('False', 2, 1, 1),
        (None, 2, 1, 1),
        (NA_ERROR, 0, 0, NA_ERROR),
        (DIV0, 0, 0, DIV0),
        (1, VALUE_ERROR, 1, VALUE_ERROR),
        (0, VALUE_ERROR, 1, 1),
        (0, 1, VALUE_ERROR, VALUE_ERROR),
        (1, 1, VALUE_ERROR, 1),
        (((1, 0), (0, 1)), 1, VALUE_ERROR,
         ((1, VALUE_ERROR), (VALUE_ERROR, 1))),
        (((1, 0), (0, 1)), 0, ((1, 2), (3, 4)), ((0, 2), (3, 0))),
        (((1, 0), (0, 1)), ((1, 2), (3, 4)), 0, ((1, 0), (0, 4))),
        (((1, 0), (0, 1)), ((1, 2), (3, 4)), ((5, 6), (7, 8)),
         ((1, 6), (7, 4))),
        (1, ((1, 2), (3, 4)), ((5, 6), (7, 8)), ((1, 2), (3, 4))),
    )
)
def test_x_if(test_value, true_value, false_value, result):
    assert x_if(test_value, true_value, false_value) == result


@pytest.mark.parametrize(
    'result, value', (
        (10, (True, 10, True, 20, False, 30)),
        (20, (False, 10, True, 20, True, 30)),
        (30, (False, 10, False, 20, True, 30)),
        (10, ("true", 10, True, 20)),
        (20, ("false", 10, True, 20)),
        (10, (2, 10, True, 20)),
        (20, (0, 10, True, 20)),
        (10, (2.1, 10, True, 20)),
        (20, (0.0, 10, True, 20)),
        (20, (None, 10, True, 20)),
        (10, (True, 10, "xyzzy", 20)),
        (VALUE_ERROR, ("xyzzy", 10, True, 20)),
        (VALUE_ERROR, (tuple(), 10, True, 20)),
        (DIV0, (DIV0, 10, True, 20)),
        (NA_ERROR, (False, 10, 0, 20, 'false', 30)),
        (NA_ERROR, (False, 10, True)),
    )
)
def test_ifs(result, value):
    assert ifs(*value) == result


@pytest.mark.parametrize(
    'result, test_value', (
        (False, True),
        (False, 1),
        (False, 2.1),
        (False, 'true'),
        (False, 'True'),
        (False, True),

        (True, False),
        (True, None),
        (True, 0),
        (True, 0.0),
        (True, 'false'),
        (True, 'faLSe'),

        (VALUE_ERROR, VALUE_ERROR),
        (NA_ERROR, NA_ERROR),
        (VALUE_ERROR, '3'),
        (VALUE_ERROR, ('1', ('x', 'y'))),
        (VALUE_ERROR, (),),
    )
)
def test_x_not(result, test_value):
    assert x_not(test_value) == result


@pytest.mark.parametrize(
    'result, test_value', (
        (False, (False,)),
        (True, (True,)),
        (True, (1, '3', 2.0, 3.1, ('x', True))),
        (True, (1, '3', 2.0, 3.1, ('x', False))),
        (False, (0, '3', 0.0, '3.1', ('x', False))),
        (VALUE_ERROR, (1, '3', 2.0, 3.1, ('x', VALUE_ERROR))),
        (NA_ERROR, (1, NA_ERROR, 2.0, 3.1, ('x', VALUE_ERROR))),
        (VALUE_ERROR, ('1', ('x', 'y'))),
        (VALUE_ERROR, (),),
    )
)
def test_x_or(result, test_value):
    assert x_or(*test_value) == result


@pytest.mark.parametrize(
    'result, test_value', (
        (False, (False,)),
        (True, (True,)),
        (False, (False, False)),
        (True, (False, True)),
        (True, (True, False)),
        (False, (False, False)),
        (False, (1, '3', 2.0, 3.1, ('x', True))),
        (True, (1, '3', 2.0, 3.1, ('x', False))),
        (False, (0, '3', 0.0, '3.1', ('x', False))),
        (VALUE_ERROR, (1, '3', 2.0, 3.1, ('x', VALUE_ERROR))),
        (NA_ERROR, (1, NA_ERROR, 2.0, 3.1, ('x', VALUE_ERROR))),
        (VALUE_ERROR, ('1', ('x', 'y'))),
        (VALUE_ERROR, (),),
    )
)
def test_x_xor(result, test_value):
    assert x_xor(*test_value) == result
