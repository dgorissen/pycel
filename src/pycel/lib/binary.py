# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of binary base number excel functions (hex2dec, etc.)
"""

import functools

from pycel.excelutil import EMPTY, ERROR_CODES, flatten, NUM_ERROR, VALUE_ERROR

_SIZE_MASK = {2: 512, 8: 0x20000000, 16: 0x8000000000}
_BASE_TO_FUNC = {2: bin, 8: oct, 16: hex}


def _base2dec(value, base):
    value = list(flatten(value))
    if len(value) != 1 or isinstance(value[0], bool):
        return VALUE_ERROR

    value = value[0]
    if value in ERROR_CODES:
        return value

    if value in (None, EMPTY):
        value = '0'
    elif isinstance(value, (int, float)) and value >= 0:
        if int(value) == value:
            value = str(int(value))

    if isinstance(value, str) and len(value) <= 10:
        try:
            value, mask = int(value, base), _SIZE_MASK[base]
            if value >= 0:
                return (value & ~mask) - (value & mask)
        except ValueError:
            return NUM_ERROR
    return NUM_ERROR


def _dec2base(value, places=None, base=16):
    value = list(flatten(value))
    if len(value) != 1 or isinstance(value[0], bool):
        return VALUE_ERROR

    value = value[0]
    if value in ERROR_CODES:
        return value

    if value in (None, EMPTY):
        if base == 8:
            return NUM_ERROR
        value = 0

    try:
        value = int(value)
    except ValueError:
        return VALUE_ERROR

    mask = _SIZE_MASK[base]
    if not (-mask <= value < mask):
        return NUM_ERROR

    if value < 0:
        value += mask << 1

    value = _BASE_TO_FUNC[base](value)[2:].upper()
    if places is None:
        places = 0
    else:
        places = int(places)
        if places < len(value):
            return NUM_ERROR
    return value.zfill(int(places))


def _base2base(value, places=None, base_in=16, base_out=16):
    if value is None:
        if base_out == 10 or base_in != 2:
            value = 0
        else:
            return NUM_ERROR
    return _dec2base(_base2dec(value, base_in), places=places, base=base_out)


# Excel reference: https://support.office.com/en-us/article/
#   HEX2DEC-function-8C8C3155-9F37-45A5-A3EE-EE5379EF106E
hex2dec = functools.partial(_base2dec, base=16)

# Excel reference: https://support.office.com/en-us/article/
#  OCT2DEC-function-87606014-CB98-44B2-8DBB-E48F8CED1554
oct2dec = functools.partial(_base2dec, base=8)

# Excel reference: https://support.office.com/en-us/article/
#   BIN2DEC-function-63905B57-B3A0-453D-99F4-647BB519CD6C
bin2dec = functools.partial(_base2dec, base=2)


# Excel reference: https://support.office.com/en-us/article/
#   DEC2HEX-function-6344EE8B-B6B5-4C6A-A672-F64666704619
dec2hex = functools.partial(_dec2base, base=16)

# Excel reference: https://support.office.com/en-us/article/
#   DEC2OCT-function-C9D835CA-20B7-40C4-8A9E-D3BE351CE00F
dec2oct = functools.partial(_dec2base, base=8)

# Excel reference: https://support.office.com/en-us/article/
#   DEC2BIN-function-0F63DD0E-5D1A-42D8-B511-5BF5C6D43838
dec2bin = functools.partial(_dec2base, base=2)


# Excel reference: https://support.office.com/en-us/article/
#   OCT2HEX-function-912175B4-D497-41B4-A029-221F051B858F
oct2hex = functools.partial(_base2base, base_in=8, base_out=16)

# Excel reference: https://support.office.com/en-us/article/
#   BIN2HEX-function-0375E507-F5E5-4077-9AF8-28D84F9F41CC
bin2hex = functools.partial(_base2base, base_in=2, base_out=16)

# Excel reference: https://support.office.com/en-us/article/
#   HEX2OCT-function-54D52808-5D19-4BD0-8A63-1096A5D11912
hex2oct = functools.partial(_base2base, base_in=16, base_out=8)

# Excel reference: https://support.office.com/en-us/article/
#   HEX2BIN-function-A13AAFAA-5737-4920-8424-643E581828C1
hex2bin = functools.partial(_base2base, base_in=16, base_out=2)

# Excel reference: https://support.office.com/en-us/article/
#   BIN2OCT-function-0A4E01BA-AC8D-4158-9B29-16C25C4C23FD
bin2oct = functools.partial(_base2base, base_in=2, base_out=8)

# Excel reference: https://support.office.com/en-us/article/
#   OCT2BIN-function-55383471-3C56-4D27-9522-1A8EC646C589
oct2bin = functools.partial(_base2base, base_in=8, base_out=2)
