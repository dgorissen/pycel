# -*- coding: UTF-8 -*-
#
# Copyright 2011-2021 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

import pycel.excellib
from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import (
    DIV0,
    NA_ERROR,
    NUM_ERROR,
    REF_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import load_to_test_module
from pycel.lib.information import (
    isblank,
    iserr,
    iserror,
    iseven,
    islogical,
    isna,
    isnontext,
    isnumber,
    isodd,
    istext,
    n,
    na,
)


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.information, __name__)


def test_information_ws(fixture_xls_copy):
    compiler = ExcelCompiler(fixture_xls_copy('information.xlsx'))
    result = compiler.validate_serialized()
    assert result == {}


@pytest.mark.parametrize(
    'value, expected', (
        (None, True),
        (0, False),
        (1, False),
        (1.0, False),
        (-1, False),
        ('a', False),
        (True, False),
        (False, False),
    )
)
def test_isblank(value, expected):
    assert isblank(value) == expected


@pytest.mark.parametrize(
    'value, expected', (
        (0, False),
        (1, False),
        (1.0, False),
        (-1, False),
        ('a', False),
        (((1, NUM_ERROR), ('2', DIV0)), ((False, True), (False, True))),
        (NA_ERROR, False),
        (NUM_ERROR, True),
        (REF_ERROR, True),
    )
)
def test_iserr(value, expected):
    assert iserr(value) == expected


@pytest.mark.parametrize(
    '_iseven, _isodd, value', (
        (True, False, -100.1),
        (True, False, '-100.1'),
        (True, False, -100),
        (False, True, -99.9),
        (True, False, 0),
        (False, True, 1),
        (True, False, 0.1),
        (True, False, '0.1'),
        (True, False, '2'),
        (True, False, 2.9),
        (False, True, 3),
        (False, True, 3.1),
        (True, False, None),
        (VALUE_ERROR, VALUE_ERROR, True),
        (VALUE_ERROR, VALUE_ERROR, False),
        (VALUE_ERROR, ) * 2 + ('xyzzy', ),
        (VALUE_ERROR, ) * 3,
        (DIV0, ) * 3,
    )
)
def test_is_even_odd(_iseven, _isodd, value):
    assert iseven(value) == _iseven
    assert isodd(value) == _isodd


@pytest.mark.parametrize(
    'value, expected', (
        (0, False),
        (1, False),
        (1.0, False),
        (-1, False),
        ('a', False),
        (((1, NA_ERROR), ('2', DIV0)), ((False, True), (False, True))),
        (NUM_ERROR, True),
        (REF_ERROR, True),
    )
)
def test_iserror(value, expected):
    assert iserror(value) == expected


@pytest.mark.parametrize(
    'value, expected', (
        (False, True),
        (True, True),
        (0, False),
        (1, False),
        (1.0, False),
        (-1, False),
        ('a', False),
        (((1, NA_ERROR), ('2', True)), ((False, False), (False, True))),
        (NA_ERROR, False),
        (VALUE_ERROR, False),
    )
)
def test_islogical(value, expected):
    assert islogical(value) == expected


@pytest.mark.parametrize(
    'value, expected', (
        (0, False),
        (1, False),
        (1.0, False),
        (-1, False),
        ('a', False),
        (((1, NA_ERROR), ('2', 3)), ((False, True), (False, False))),
        (NA_ERROR, True),
        (VALUE_ERROR, False),
    )
)
def test_isna(value, expected):
    assert isna(value) == expected


@pytest.mark.parametrize(
    'value, expected', (
        (0, True),
        (1, True),
        (1.0, True),
        (-1, True),
        ('a', False),
        (False, False),
        (True, False),
        (((1, NA_ERROR), ('2', 3)), ((True, False), (False, True))),
        (NA_ERROR, False),
        (VALUE_ERROR, False),
    )
)
def test_isnumber(value, expected):
    assert isnumber(value) == expected


@pytest.mark.parametrize(
    'value, expected', (
        ('a', True),
        (1, False),
        (1.0, False),
        (None, False),
        (DIV0, False),
        (((1, NA_ERROR), ('2', 3)), ((False, False), (True, False))),
        (NA_ERROR, False),
        (VALUE_ERROR, False),
    )
)
def test_istext(value, expected):
    assert istext(value) == expected
    assert isnontext(value) != expected


@pytest.mark.parametrize(
    'value, expected', (
        (False, 0),
        (True, 1),
        ('a', 0),
        (1, 1),
        (1.0, 1.0),
        (-1.0, -1.0),
        (None, None),
        (DIV0, DIV0),
        (((1, NA_ERROR), ('2', 3)), ((1, NA_ERROR), (0, 3))),
        (NA_ERROR, NA_ERROR),
        (VALUE_ERROR, VALUE_ERROR),
    )
)
def test_n(value, expected):
    assert n(value) == expected


def test_na():
    assert na() == NA_ERROR
