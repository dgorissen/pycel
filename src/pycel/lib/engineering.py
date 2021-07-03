# -*- coding: UTF-8 -*-
#
# Copyright 2011-2021 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of Engineering library functions
"""
import functools

from pycel.excelutil import EMPTY, ERROR_CODES, flatten, NUM_ERROR, VALUE_ERROR
from pycel.lib.function_helpers import (
    excel_math_func,
)


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


# def besseli(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   besseli-function-8d33855c-9a8d-444b-98e0-852267b1c0df


# def besselj(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   besselj-function-839cb181-48de-408b-9d80-bd02982d94f7


# def besselk(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   besselk-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70


# def bessely(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   bessely-function-f3a356b3-da89-42c3-8974-2da54d6353a2


# Excel reference: https://support.microsoft.com/en-us/office/
#   BIN2DEC-function-63905B57-B3A0-453D-99F4-647BB519CD6C
bin2dec = functools.partial(_base2dec, base=2)


# Excel reference: https://support.microsoft.com/en-us/office/
#   BIN2HEX-function-0375E507-F5E5-4077-9AF8-28D84F9F41CC
bin2hex = functools.partial(_base2base, base_in=2, base_out=16)


# Excel reference: https://support.microsoft.com/en-us/office/
#   BIN2OCT-function-0A4E01BA-AC8D-4158-9B29-16C25C4C23FD
bin2oct = functools.partial(_base2base, base_in=2, base_out=8)


@excel_math_func
def bitand(op_x, op_y):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   bitand-function-8a2be3d7-91c3-4b48-9517-64548008563a
    if op_x < 0 or op_y < 0:
        return NUM_ERROR
    return op_x & op_y


@excel_math_func
def bitlshift(number, pos):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   bitlshift-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c
    if number < 0 or abs(pos) > 53 or number >= 2**48:
        return NUM_ERROR
    if pos < 0:
        return bitrshift(number, abs(pos))
    return number << pos


@excel_math_func
def bitor(op_x, op_y):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   bitor-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2
    if op_x < 0 or op_y < 0:
        return NUM_ERROR
    return op_x | op_y


@excel_math_func
def bitrshift(number, pos):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   bitrshift-function-274d6996-f42c-4743-abdb-4ff95351222c
    if number < 0 or abs(pos) > 53 or number >= 2 ** 48:
        return NUM_ERROR
    if pos < 0:
        return bitlshift(number, abs(pos))
    return number >> pos


@excel_math_func
def bitxor(op_x, op_y):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   bitxor-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4
    if op_x < 0 or op_y < 0:
        return NUM_ERROR
    return op_x ^ op_y


# def complex(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   complex-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128


# def convert(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   convert-function-d785bef1-808e-4aac-bdcd-666c810f9af2


# Excel reference: https://support.microsoft.com/en-us/office/
#   DEC2BIN-function-0F63DD0E-5D1A-42D8-B511-5BF5C6D43838
dec2bin = functools.partial(_dec2base, base=2)


# Excel reference: https://support.microsoft.com/en-us/office/
#   DEC2HEX-function-6344EE8B-B6B5-4C6A-A672-F64666704619
dec2hex = functools.partial(_dec2base, base=16)


# Excel reference: https://support.microsoft.com/en-us/office/
#   DEC2OCT-function-C9D835CA-20B7-40C4-8A9E-D3BE351CE00F
dec2oct = functools.partial(_dec2base, base=8)


# def delta(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   delta-function-2f763672-c959-4e07-ac33-fe03220ba432


# def erf(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   erf-function-c53c7e7b-5482-4b6c-883e-56df3c9af349


# def erf.precise(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   erf-precise-function-9a349593-705c-4278-9a98-e4122831a8e0


# def erfc(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   erfc-function-736e0318-70ba-4e8b-8d08-461fe68b71b3


# def erfc.precise(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   erfc-precise-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273


# def gestep(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   gestep-function-f37e7d2a-41da-4129-be95-640883fca9df


# Excel reference: https://support.microsoft.com/en-us/office/
#   HEX2BIN-function-A13AAFAA-5737-4920-8424-643E581828C1
hex2bin = functools.partial(_base2base, base_in=16, base_out=2)


# Excel reference: https://support.microsoft.com/en-us/office/
#   HEX2DEC-function-8C8C3155-9F37-45A5-A3EE-EE5379EF106E
hex2dec = functools.partial(_base2dec, base=16)


# Excel reference: https://support.microsoft.com/en-us/office/
#   HEX2OCT-function-54D52808-5D19-4BD0-8A63-1096A5D11912
hex2oct = functools.partial(_base2base, base_in=16, base_out=8)


# def imabs(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imabs-function-b31e73c6-d90c-4062-90bc-8eb351d765a1


# def imaginary(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imaginary-function-dd5952fd-473d-44d9-95a1-9a17b23e428a


# def imargument(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imargument-function-eed37ec1-23b3-4f59-b9f3-d340358a034a


# def imconjugate(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imconjugate-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42


# def imcos(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imcos-function-dad75277-f592-4a6b-ad6c-be93a808a53c


# def imcosh(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imcosh-function-053e4ddb-4122-458b-be9a-457c405e90ff


# def imcot(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imcot-function-dc6a3607-d26a-4d06-8b41-8931da36442c


# def imcsc(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imcsc-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323


# def imcsch(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imcsch-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9


# def imdiv(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imdiv-function-a505aff7-af8a-4451-8142-77ec3d74d83f


# def imexp(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imexp-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f


# def imln(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imln-function-32b98bcf-8b81-437c-a636-6fb3aad509d8


# def imlog10(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imlog10-function-58200fca-e2a2-4271-8a98-ccd4360213a5


# def imlog2(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imlog2-function-152e13b4-bc79-486c-a243-e6a676878c51


# def impower(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   impower-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39


# def improduct(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   improduct-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba


# def imreal(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imreal-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366


# def imsec(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsec-function-6df11132-4411-4df4-a3dc-1f17372459e0


# def imsech(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsech-function-f250304f-788b-4505-954e-eb01fa50903b


# def imsin(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsin-function-1ab02a39-a721-48de-82ef-f52bf37859f6


# def imsinh(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsinh-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d


# def imsqrt(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsqrt-function-e1753f80-ba11-4664-a10e-e17368396b70


# def imsub(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsub-function-2e404b4d-4935-4e85-9f52-cb08b9a45054


# def imsum(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imsum-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f


# def imtan(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   imtan-function-8478f45d-610a-43cf-8544-9fc0b553a132


# Excel reference: https://support.microsoft.com/en-us/office/
#   OCT2BIN-function-55383471-3C56-4D27-9522-1A8EC646C589
oct2bin = functools.partial(_base2base, base_in=8, base_out=2)


# Excel reference: https://support.microsoft.com/en-us/office/
#  OCT2DEC-function-87606014-CB98-44B2-8DBB-E48F8CED1554
oct2dec = functools.partial(_base2dec, base=8)


# Excel reference: https://support.microsoft.com/en-us/office/
#   OCT2HEX-function-912175B4-D497-41B4-A029-221F051B858F
oct2hex = functools.partial(_base2base, base_in=8, base_out=16)
