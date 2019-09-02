import pytest

import pycel.excellib
from pycel.excelutil import (
    DIV0,
    NA_ERROR,
    NAME_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import load_to_test_module
from pycel.lib.text import (
    concat,
    concatenate,
    find,
    left,
    mid,
    right,
    value,
    x_len,
)

# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.text, __name__)


@pytest.mark.parametrize(
    'args, result', (
        ('a 1 abc'.split(), 'a1abc'),
        ('a Jan-00 abc'.split(), 'aJan-00abc'),
        ('a	#DIV/0! abc'.split(), DIV0),
        ('a	1 #DIV/0!'.split(), DIV0),
        ('a #NAME? abc'.split(), NAME_ERROR),
        (('a', True, 'abc'), 'aTRUEabc'),
        (('a', False, 'abc'), 'aFALSEabc'),
        (('a', 2, 'abc'), 'a2abc'),
    )
)
def test_concatenate(args, result):
    assert concat(*args) == result
    assert concatenate(*args) == result
    assert concat(args) == result
    assert concatenate(args) == VALUE_ERROR


@pytest.mark.parametrize(
    'to_find, find_in, result', (
        (2, 2.5, 1),
        ('.', 2.5, 2),
        (5, 2.5, 3),
        ('2', 2.5, 1),
        ('.', 2.5, 2),
        ('5', 2.5, 3),
        ('2', '2.5', 1),
        ('.', '2.5', 2),
        ('T', True, 1),
        ('U', True, 3),
        ('u', True, VALUE_ERROR),
        (DIV0, 'x' + DIV0, DIV0),
        ('V', DIV0, DIV0),
    )
)
def test_find(to_find, find_in, result):
    assert find(to_find, find_in) == result


@pytest.mark.parametrize(
    'text, num_chars, expected', (
        ('abcd', 5, 'abcd'),
        ('abcd', 4, 'abcd'),
        ('abcd', 3, 'abc'),
        ('abcd', 2, 'ab'),
        ('abcd', 1, 'a'),
        ('abcd', 0, ''),

        (1.234, 3, '1.2'),

        ('abcd', -1, VALUE_ERROR),
        ('abcd', 'x', VALUE_ERROR),
        (DIV0, 1, DIV0),
        ('abcd', NAME_ERROR, NAME_ERROR),
    )
)
def test_left(text, num_chars, expected):
    assert left(text, num_chars) == expected


class TestMid:

    def test_invalid_parameters(self):
        assert mid(VALUE_ERROR, 2, 2) == VALUE_ERROR
        assert mid('Romain', VALUE_ERROR, 2) == VALUE_ERROR
        assert mid('Romain', 2, VALUE_ERROR) == VALUE_ERROR
        assert mid(DIV0, 2, 2) == DIV0
        assert mid('Romain', DIV0, 2) == DIV0
        assert mid('Romain', 2, DIV0) == DIV0

        assert mid('Romain', 'x', 2) == VALUE_ERROR
        assert mid('Romain', 2, 'x') == VALUE_ERROR

    def test_num_chars_must_be_integer(self):
        assert 'Ro' == mid('Romain', 1, 2.1)

    def test_start_num_must_be_superior_or_equal_to_1(self):
        assert VALUE_ERROR == mid('Romain', 0, 3)

    def test_num_chars_must_be_positive(self):
        assert VALUE_ERROR == mid('Romain', 1, -1)

    def test_from_not_str(self):
        assert '23' == mid(1234, 2, 2)

    def test_mid(self):
        assert 'omain' == mid('Romain', 2, 9)
        assert 'om' == mid('Romain', 2.1, 2)
        assert 'om' == mid('Romain', 2, 2.1)


@pytest.mark.parametrize(
    'text, num_chars, expected', (
        ('abcd', 5, 'abcd'),
        ('abcd', 4, 'abcd'),
        ('abcd', 3, 'bcd'),
        ('abcd', 2, 'cd'),
        ('abcd', 1, 'd'),
        ('abcd', 0, ''),

        (1234.1, 2, '.1'),

        ('abcd', -1, VALUE_ERROR),
        ('abcd', 'x', VALUE_ERROR),
        (VALUE_ERROR, 1, VALUE_ERROR),
        ('abcd', VALUE_ERROR, VALUE_ERROR),
    )
)
def test_right(text, num_chars, expected):
    assert right(text, num_chars) == expected


@pytest.mark.parametrize(
    'param, result', (
        (0, 0),
        (2, 2),
        (2.1, 2.1),
        (-2.1, -2.1),
        ('-2.1', -2.1),
        ('3', 3),
        ('3.', 3),
        ('3.0', 3),
        ('.01', 0.01),
        ('1E5', 100000),
        ('X', VALUE_ERROR),
        ('`1', VALUE_ERROR),
        (False, VALUE_ERROR),
        (True, VALUE_ERROR),
        (NA_ERROR, NA_ERROR),
        (DIV0, DIV0),
    )
)
def test_value(param, result):
    assert value(param) == result


@pytest.mark.parametrize(
    'param, result', (
        ('A', 1),
        ('BB', 2),
        (3.0, 3),
        (True, 4),
        (False, 5),
        (None, 0),
        (NA_ERROR, NA_ERROR),
        (DIV0, DIV0),
    )
)
def test_x_len(param, result):
    assert x_len(param) == result
