# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of text excel functions (lower, upper, etc.)
"""
import re
from datetime import datetime

from pycel.excelutil import (
    coerce_to_string,
    ERROR_CODES,
    flatten,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import excel_helper

RE_MULTI_SPACE = re.compile(' +')


# def asc(text):
# Excel reference: https://support.office.com/en-us/article/
#   asc-function-0b6abf1c-c663-4004-a964-ebc00b723266


# def bahttext(text):
# Excel reference: https://support.office.com/en-us/article/
#   bahttext-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c


# def char(text):
# Excel reference: https://support.office.com/en-us/article/
#   char-function-bbd249c8-b36e-4a91-8017-1c133f9b837a


# def clean(text):
# Excel reference: https://support.office.com/en-us/article/
#   clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41


# def code(text):
# Excel reference: https://support.office.com/en-us/article/
#   code-function-c32b692b-2ed0-4a04-bdd9-75640144b928


def concat(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2
    return concatenate(*tuple(flatten(args)))


def concatenate(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   CONCATENATE-function-8F8AE884-2CA8-4F7A-B093-75D702BEA31D
    if tuple(flatten(args)) != args:
        return VALUE_ERROR

    error = next((x for x in args if x in ERROR_CODES), None)
    if error:
        return error

    return ''.join(coerce_to_string(a) for a in args)


# def dbcs(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   dbcs-function-a4025e73-63d2-4958-9423-21a24794c9e5


# def dollar(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   dollar-function-a6cd05d9-9740-4ad3-a469-8109d18ff611


# def exact(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   exact-function-d3087698-fc15-4a15-9631-12575cf29926


@excel_helper(cse_params=(0, 1, 2), number_params=2)
def find(find_text, within_text, start_num=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   FIND-FINDB-functions-C7912941-AF2A-4BDF-A553-D0D89B0A0628
    find_text = coerce_to_string(find_text)
    within_text = coerce_to_string(within_text)
    found = within_text.find(find_text, start_num - 1)
    if found == -1:
        return VALUE_ERROR
    else:
        return found + 1


# def findb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628


# def fixed(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   fixed-function-ffd5723c-324c-45e9-8b96-e41be2a8274a


# def jis(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   jis-function-b72fb1a7-ba52-448a-b7d3-d2610868b7e2


@excel_helper(cse_params=(0, 1), number_params=1)
def left(text, num_chars=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   LEFT-LEFTB-functions-9203D2D2-7960-479B-84C6-1EA52B99640C
    if num_chars < 0:
        return VALUE_ERROR
    else:
        return str(text)[:int(num_chars)]


# def leftb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c


@excel_helper(cse_params=0)
def x_len(arg):
    # Excel reference: https://support.office.com/en-us/article/
    #   len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb
    return 0 if arg is None else len(str(arg))


# def lenb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb


@excel_helper(cse_params=0)
def lower(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4
    return coerce_to_string(text).lower()


@excel_helper(cse_params=-1, number_params=(1, 2))
def mid(text, start_num, num_chars):
    # Excel reference: https://support.office.com/en-us/article/
    #   MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028

    if start_num < 1 or num_chars < 0:
        return VALUE_ERROR

    start_num = int(start_num) - 1

    return str(text)[start_num:start_num + int(num_chars)]


# def midb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028


# def numbervalue(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   numbervalue-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879


# def phonetic(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   phonetic-function-9a329dac-0c0f-42f8-9a55-639086988554


# def proper(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   proper-function-52a5a283-e8b2-49be-8506-b2887b889f94


@excel_helper(cse_params=-1, number_params=(1, 2))
def replace(old_text, start_num, num_chars, new_text):
    # Excel reference: https://support.office.com/en-us/article/
    #   replace-replaceb-functions-8d799074-2425-4a8a-84bc-82472868878a
    old_text = coerce_to_string(old_text)
    new_text = coerce_to_string(new_text)
    start_num = int(start_num) - 1
    num_chars = int(num_chars)
    if start_num < 0 or num_chars < 0:
        return VALUE_ERROR
    return '{}{}{}'.format(
        old_text[:start_num], new_text, old_text[start_num + num_chars:])


# def replaceb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   replace-replaceb-functions-8d799074-2425-4a8a-84bc-82472868878a


# def rept(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   rept-function-04c4d778-e712-43b4-9c15-d656582bb061


@excel_helper(cse_params=(0, 1), number_params=1)
def right(text, num_chars=1):
    # Excel reference:  https://support.office.com/en-us/article/
    #   RIGHT-RIGHTB-functions-240267EE-9AFA-4639-A02B-F19E1786CF2F

    if num_chars < 0:
        return VALUE_ERROR
    elif num_chars == 0:
        return ''
    else:
        return str(text)[-int(num_chars):]


# def rightb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f


# def search(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495


# def searchb(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495


# def substitute(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332


# def t(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   t-function-fb83aeec-45e7-4924-af95-53e073541228


@excel_helper(cse_params=0)
def text(text_value, value_format):
    # Excel reference: https://support.office.com/en-us/article/
    #   text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c
    def _get_datetime_format(excel_format):
        fmt = excel_format.lower()
        fmt.replace('a/p', 'am/pm', 1)
        hour_fmt = '%H' if 'am/pm' not in fmt else '%I'
        py_fmt = {
            'dddd': '%A',
            'ddd': '%a',
            'dd': '%d',
            ':mm': ':%M',
            'mm:': '%M:',
            ':mm:': ':%M:',
            'mmmmm': '%b',
            'mmmm': '%B',
            'mmm': '%b',
            'mm': '%m',
            'am/pm': '%p',
            'yyyy': '%Y',
            'yyy': '%Y',
            'yy': '%y',
            'hh:': hour_fmt + ":",
            'h:': hour_fmt + ":",
            'hh': hour_fmt,
            '[h]': hour_fmt,
            'h': hour_fmt,
            ':ss': ':%S',
            'ss': '%S',
            ':s': ':%S',
            's': '%S',
            'd': '%d',
            'm': '%m',
        }

        replaced = set()
        for fmt_excel, fmt_python in py_fmt.items():
            if fmt_excel in fmt:
                if fmt.find(fmt_excel) in replaced or fmt_python in fmt:
                    continue
                fmt = fmt.replace(fmt_excel, fmt_python, 1)
                s = fmt.find(fmt_python)
                replaced.update(set([x for x in range(s, s + len(fmt_python))]))
        return fmt

    date_format = _get_datetime_format(value_format)
    if isinstance(text_value, str):
        if any(x in text_value for x in ('-', '/', ':', 'am', 'pm')):
            date_value = None
            time_value = None
            tokens = text_value.split(" ")
            add_locale = ''
            hour_fmt = 'H'
            if 'am' in tokens or 'pm' in tokens:
                add_locale = ' %p'
                hour_fmt = 'I'

            python_time_formats = set()
            adds = ('', '-')
            for h in adds:
                for m in adds:
                    for s in adds:
                        python_time_formats.add(
                            f'%{h}{hour_fmt}:%{m}M:%{s}S{add_locale}'
                        )

            python_time_formats.update(
                set([fmt[:fmt.index('M:') + 1:] + add_locale
                     for fmt in python_time_formats])
            )

            for token in tokens:
                if '/' in token or '-' in token:
                    for python_fmt in (
                            '%d/%m/%y',
                            '%d/%m/%Y',
                            '%m/%d/%y',
                            '%m/%d/%Y',
                            '%Y-%m-%d'
                    ):
                        try:
                            date_value = datetime.strptime(token, python_fmt)
                            break
                        except ValueError:
                            continue
                elif ':' in token:
                    if 'am' in tokens:
                        token += ' am'
                    elif 'pm' in tokens:
                        token += ' pm'
                    for python_fmt in python_time_formats:
                        try:
                            time_value = datetime.strptime(token, python_fmt)
                            break
                        except ValueError:
                            continue
            if isinstance(time_value, datetime):
                if isinstance(date_value, datetime):
                    date_value = datetime.combine(date_value, time_value.time())
                else:
                    date_value = time_value

            if isinstance(date_value, datetime):
                return date_value.strftime(date_format)

        is_pcnt = '%' in value_format

        if '#' in value_format or '0' in value_format or is_pcnt:
            if "#,#" not in value_format and "0,0" not in value_format:
                thousand_sep = ""
            else:
                thousand_sep = ","
            decimals = 0
            dec_sep = value_format.find('.')
            if dec_sep >= 0:
                decimals = value_format[dec_sep::].count('0')
            num_v = float("".join(
                [x for x in text_value if x.isdecimal() or x == '.']
            ))
            if is_pcnt:
                num_v *= 100
            num_v = round(num_v, decimals)
            if decimals == 0:
                num_v = int(num_v)
            res = f'{num_v:{thousand_sep}.{decimals}f}{"%" if is_pcnt else ""}'
            if not value_format[0] in ('#', '.', ',', '0'):
                res = value_format[0] + res
            return res

    return str(text_value)

# def textjoin(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c


@excel_helper(cse_params=0)
def trim(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9
    return RE_MULTI_SPACE.sub(' ', coerce_to_string(text))


# def unichar(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   unichar-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8


# def unicode(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   unicode-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f


@excel_helper(cse_params=0)
def upper(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6
    return coerce_to_string(text).upper()


@excel_helper(cse_params=0)
def value(text):
    # Excel reference: https://support.office.com/en-us/article/
    #   VALUE-function-257D0108-07DC-437D-AE1C-BC2D3953D8C2
    if isinstance(text, bool):
        return VALUE_ERROR
    try:
        return float(text)
    except ValueError:
        return VALUE_ERROR
