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
from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import (
    DIV0,
    in_array_formula_context,
    NA_ERROR,
    NAME_ERROR,
    NULL_ERROR,
    NUM_ERROR,
    REF_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import load_to_test_module
from pycel.lib.logical import (
    _clean_logicals,
    and_,
    if_,
    iferror,
    ifna,
    ifs,
    not_,
    or_,
    xor_,
)


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.logical, __name__)


def test_logical_ws(fixture_xls_copy):
    compiler = ExcelCompiler(fixture_xls_copy('logical.xlsx'))
    result = compiler.validate_serialized()
    assert result == {}


@pytest.mark.parametrize(
    'test_value, expected', (
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
def test_clean_logicals(test_value, expected):
    assert _clean_logicals(*test_value) == expected


@pytest.mark.parametrize(
    'expected, test_value', (
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
def test_and_(expected, test_value):
    assert and_(*test_value) == expected


@pytest.mark.parametrize(
    'test_value, error_value, expected', (
        ('A', 2, 'A'),
        (NULL_ERROR, 2, 2),
        (DIV0, 2, 2),
        (VALUE_ERROR, 2, 2),
        (REF_ERROR, 2, 2),
        (NAME_ERROR, 2, 2),
        (NUM_ERROR, 2, 2),
        (NA_ERROR, 2, 2),
        (NA_ERROR, None, 0),
        (((1, VALUE_ERROR), (VALUE_ERROR, 1)), 2, ((1, 2), (2, 1))),
        (((1, VALUE_ERROR), (VALUE_ERROR, 1)), None, ((1, 0), (0, 1))),
    )
)
def test_iferror(test_value, error_value, expected):
    if isinstance(test_value, tuple):
        with in_array_formula_context('A1'):
            assert iferror(test_value, error_value) == expected
        expected = 0 if error_value is None else error_value

    assert iferror(test_value, error_value) == expected


@pytest.mark.parametrize(
    'test_value, na_value, expected', (
        ('A', 2, 'A'),
        (NULL_ERROR, 2, NULL_ERROR),
        (DIV0, 2, DIV0),
        (VALUE_ERROR, 2, VALUE_ERROR),
        (REF_ERROR, 2, REF_ERROR),
        (NAME_ERROR, 2, NAME_ERROR),
        (NUM_ERROR, 2, NUM_ERROR),
        (NA_ERROR, 2, 2),
        (NA_ERROR, None, 0),
        (((1, NA_ERROR), (NA_ERROR, 1)), 2, ((1, 2), (2, 1))),
        (((1, NA_ERROR), (NA_ERROR, 1)), None, ((1, 0), (0, 1))),
    )
)
def test_ifna(test_value, na_value, expected):
    if isinstance(test_value, tuple):
        with in_array_formula_context('A1'):
            assert ifna(test_value, na_value) == expected
        expected = 0 if na_value is None else na_value

    assert ifna(test_value, na_value) == expected


@pytest.mark.parametrize(
    'test_value, true_value, false_value, expected', (
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
def test_if_(test_value, true_value, false_value, expected):
    assert if_(test_value, true_value, false_value) == expected


@pytest.mark.parametrize(
    'expected, value', (
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
        (DIV0, (DIV0, 10, True, 20)),
        (NA_ERROR, (False, 10, 0, 20, 'false', 30)),
        (NA_ERROR, (False, 10, True)),
        ((('A', DIV0), (NA_ERROR, 4)),
         (((1, DIV0), (0, 0)), 'A', ((0, 0), (0, 1)), ((1, 2), (3, 4)))
         ),
    )
)
def test_ifs(expected, value):
    if any(isinstance(v, tuple) for v in value):
        with in_array_formula_context('A1'):
            assert ifs(*value) == expected
    else:
        assert ifs(*value) == expected


@pytest.mark.parametrize(
    'expected, test_value', (
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
def test_not_(expected, test_value):
    assert not_(test_value) == expected


@pytest.mark.parametrize(
    'expected, test_value', (
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
def test_or_(expected, test_value):
    assert or_(*test_value) == expected


@pytest.mark.parametrize(
    'expected, test_value', (
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
def test_xor_(expected, test_value):
    assert xor_(*test_value) == expected
