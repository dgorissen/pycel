# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

from pycel.excelutil import coerce_to_number, ERROR_CODES
from pycel.lib import binary

MAX_BASE_2 = binary._SIZE_MASK[2]
MAX_BASE_8 = binary._SIZE_MASK[8]
MAX_BASE_16 = binary._SIZE_MASK[16]


def compare_result(expected, result):
    expected = coerce_to_number(expected)
    result = coerce_to_number(result)
    if isinstance(expected, (int, float)) and isinstance(result, (int, float)):
        return pytest.approx(expected) == result
    else:
        return expected == result


@pytest.mark.parametrize(
    'value, base, expected', (
        ('0', 2, '0.0'),
        (0, 8, '0.0'),
        (0, 16, '0.0'),
        ('111111111', 2, 511.0),
        ('7777777777', 8, '-1.0'),
        ('9999999999', 16, -439804651111.0),
        (3.5, 2, '#NUM!'),
        (3.5, 8, '#NUM!'),
        ('3.5', 16, '#NUM!'),
        ('1000000000', 2, '-512.0'),
        ('11111111111', 8, '#NUM!'),
        ('11111111111', 16, '#NUM!'),
        (-1, 2, '#NUM!'),
        (-1, 8, '#NUM!'),
        (-1, 16, '#NUM!'),
        (None, 2, '0.0'),
        (None, 8, '0.0'),
        (None, 16, '0.0'),
        ('xyzzy', 2, '#NUM!'),
        ('a', 8, '#NUM!'),
        ('f10000001f', 16, -64424509409),
        (True, 2, '#VALUE!'),
        (True, 8, '#VALUE!'),
        (True, 16, '#VALUE!'),
    )
)
def test_base2dec(value, base, expected):
    assert compare_result(expected, binary._base2dec(value, base))

    mapped = {2: binary.bin2dec, 8: binary.oct2dec, 16: binary.hex2dec}
    assert compare_result(expected, mapped[base](value))


@pytest.mark.parametrize('value', tuple(ERROR_CODES))
def test_base2dec_errors(value):
    for base in (2, 8, 16):
        assert compare_result(value, binary._base2dec(value, base))


@pytest.mark.parametrize(
    'value, base, expected', (
        (MAX_BASE_2, 2, '#NUM!'),
        (MAX_BASE_8, 8, '#NUM!'),
        (MAX_BASE_16, 16, '#NUM!'),
        (MAX_BASE_2 - 1, 2, '111111111'),
        (MAX_BASE_8 - 1, 8, '3777777777'),
        (MAX_BASE_16 - 1, 16, '7FFFFFFFFF'),
        (-MAX_BASE_2, 2, '1000000000'),
        (-MAX_BASE_8, 8, '4000000000'),
        (-MAX_BASE_16, 16, '8000000000'),
        (-MAX_BASE_2 - 1, 2, '#NUM!'),
        (-MAX_BASE_8 - 1, 8, '#NUM!'),
        (-MAX_BASE_16 - 1, 16, '#NUM!'),
        ('xyzzy', 2, '#VALUE!'),
        ('xyzzy', 8, '#VALUE!'),
        ('xyzzy', 16, '#VALUE!'),
        (True, 2, '#VALUE!'),
        (True, 8, '#VALUE!'),
        (True, 16, '#VALUE!'),
    )
)
def test_dec2base(value, base, expected):
    assert compare_result(expected, binary._dec2base(value, base=base))

    mapped = {2: binary.dec2bin, 8: binary.dec2oct, 16: binary.dec2hex}
    assert compare_result(expected, mapped[base](value))


@pytest.mark.parametrize(
    'value, base, places, expected', (
        (100, 2, 1, '#NUM!'),
        (100, 8, 1, '#NUM!'),
        (100, 16, 1, '#NUM!'),
        (None, 2, 3, '000'),
        (None, 8, 0, '#NUM!'),
        (None, 16, 1, '0'),
    )
)
def test_dec2base_places(value, base, places, expected):
    assert compare_result(expected,
                          binary._dec2base(value, base=base, places=places))


@pytest.mark.parametrize('value', tuple(ERROR_CODES))
def test_dec2base_errors(value):
    for base in (2, 8, 16):
        assert compare_result(value, binary._dec2base(value, base=base))


@pytest.mark.parametrize(
    'value, bases, expected', (
        ('111111111', (2, 8), '777'),
        ('111111111', (2, 16), '1FF'),
        ('7777777777', (8, 2), '1111111111'),
        ('7777777777', (8, 16), 'FFFFFFFFFF'),
        ('9999999999', (16, 2), '#NUM!'),
        ('9999999999', (16, 8), '#NUM!'),
        ('1000000000', (2, 8), '7777777000'),
        ('1000000000', (2, 16), 'FFFFFFFE00'),
        ('11111111111', (8, 2), '#NUM!'),
        ('11111111111', (8, 16), '#NUM!'),
        ('11111111111', (16, 2), '#NUM!'),
        ('11111111111', (16, 8), '#NUM!'),
        (None, (2, 8), '#NUM!'),
        (None, (2, 16), '#NUM!'),
        (None, (8, 2), '0'),
        (None, (8, 16), '0'),
        (None, (16, 2), '0'),
        (None, (16, 8), '0'),
        ('fffffffffe', (2, 8), '#NUM!'),
        ('fffffffffe', (2, 16), '#NUM!'),
        ('a', (8, 2), '#NUM!'),
        ('a', (8, 16), '#NUM!'),
        ('fffffffffe', (16, 2), '1111111110'),
        ('fffffffffe', (16, 8), '7777777776'),
    )
)
def test_base2base(value, bases, expected):
    base_in, base_out = bases
    assert compare_result(expected, binary._base2base(
        value, base_in=base_in, base_out=base_out))

    mapped = {
        (2, 8): binary.bin2oct,
        (2, 16): binary.bin2hex,
        (8, 2): binary.oct2bin,
        (8, 16): binary.oct2hex,
        (16, 2): binary.hex2bin,
        (16, 8): binary.hex2oct,
    }
    assert compare_result(expected, mapped[bases](value))


@pytest.mark.parametrize(
    'value, expected', (
        (True, '#VALUE!'),
        (-1, '#NUM!'),
        ('-1', '#NUM!'),
        ('3.5', '#NUM!'),
        (3.5, '#NUM!'),
        ('0', '0'),
        (0, '0'),
    )
)
def test_base2base_all_bases(value, expected):
    for base_in in (2, 8, 16):
        for base_out in (2, 8, 16):
            if base_in != base_out:
                assert compare_result(
                    expected, binary._base2base(
                        value, base_in=base_in, base_out=base_out))


@pytest.mark.parametrize('value', tuple(ERROR_CODES))
def test_base2base_errors(value):
    for base_in in (2, 8, 16):
        for base_out in (2, 8, 16):
            assert compare_result(value, binary._base2base(
                value, base_in=base_in, base_out=base_out))
