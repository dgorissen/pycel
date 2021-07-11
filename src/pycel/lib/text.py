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
import collections
import itertools as it
import locale
import re
from enum import Enum
from typing import Iterable, List

from pycel.excelutil import (
    coerce_to_number,
    coerce_to_string,
    ERROR_CODES,
    flatten,
    VALUE_ERROR,
)
from pycel.lib.date_time import DateTimeFormatter
from pycel.lib.function_helpers import excel_helper

RE_MULTI_SPACE = re.compile(' +')


class TextFormat:
    Element = collections.namedtuple('Element', 'position code next_code char')
    Token = collections.namedtuple('Token', 'token type position')
    Tokenized = collections.namedtuple('Tokenized', 'tokens types decimal thousands percents')

    FORMAT_MISC = set("$+(:^\'{<=-/)!&~}> ")
    FORMAT_NUMBER = set('0#?.,%')
    FORMAT_PLACEHOLDER = set('#0?')
    DIGITS = set('0123456789')
    NUMBER_TOKEN_MATCH = {'#': None, '0': '0', '?': '0'}

    class TokenType(Enum):
        STRING = 1
        NUMBER = 2
        DATETIME = 3
        AM_PM = 4
        REPLACE = 5

    FORMAT_TYPES = {TokenType.DATETIME, TokenType.NUMBER, TokenType.REPLACE}

    def __init__(self, format: str):
        self.format = format
        try:
            self.tokenized_formats = tuple(self._tokenize_format(format))
        except ValueError:
            self.tokenized_formats = VALUE_ERROR

        self.thousands_format = ',' if locale.setlocale(locale.LC_NUMERIC) == 'C' else 'n'

    @classmethod
    def _find_am_pm(cls, element, format, stream):
        if element.code == 'a' and element.next_code in 'm/':
            if element.next_code == 'm':
                to_match = 'am/pm'
            else:
                to_match = 'a/p'
            matched = format[element.position:element.position + len(to_match)]
            if matched.lower() == to_match:
                for i in range(len(to_match) - 1):
                    next(stream)
                return matched if to_match == 'a/p' else to_match
        return None

    @classmethod
    def _get_matching_codes(cls, element, stream, eos_allowed=True):
        elements = [element]
        while elements[-1].code == elements[-1].next_code:
            elements.append(next(stream))

        if not eos_allowed and elements[-1].next_code is None:
            raise ValueError

        return "".join(e.code for e in elements)

    def _tokenize_format(self, format: str) -> Iterable[Tokenized]:
        """Break up the tokens by type and section"""

        tokens = []
        last_date = None
        have_decimal = False
        have_thousands = False
        percents = 0

        # amend cls.Token stream to ease code production
        stream = iter(self.Element(i, *e) for i, e in enumerate(
            zip(format.lower(), list(format[1:].lower()) + [None], format)))

        for element in stream:
            if element.char == '"':
                tokens.append(self.Token(''.join(
                    e.char for e in it.takewhile(lambda x: x.code != '"', stream)),
                    self.TokenType.STRING, element.position))

            elif element.char == '\\':
                if element.next_code is None:
                    raise ValueError
                tokens.append(self.Token(
                    next(stream).char, self.TokenType.STRING, element.position))

            elif element.char == ';':
                yield self._finalize_tokenize(tokens, have_decimal, have_thousands, percents)
                tokens = []
                have_decimal = False
                have_thousands = False
                percents = 0

            elif element.char == '@':
                tokens.append(self.Token(element.char, self.TokenType.REPLACE, element.position))

            elif element.code in self.FORMAT_NUMBER and not (
                    last_date and (
                        (last_date[0].token[0] == 's' or last_date[0].token == '[s]') and
                        element.code == '.' or element.code == ',')):

                need_emit = True
                if element.code == ',':
                    need_emit = False
                    if (have_decimal or have_thousands or
                            element.position == 0 or element.next_code is None or
                            format[element.position - 1] not in self.FORMAT_PLACEHOLDER or
                            format[element.position + 1] not in self.FORMAT_PLACEHOLDER):
                        # just a regular comma, not 1000's indicator
                        if element.position == 0 or (
                                format[element.position - 1] not in self.FORMAT_PLACEHOLDER):
                            tokens.append(self.Token(
                                element.code, self.TokenType.STRING, element.position))
                    else:
                        have_thousands = True

                elif element.code == '.':
                    if have_decimal:
                        need_emit = False
                        tokens.append(self.Token(
                            element.code, self.TokenType.STRING, element.position))
                    else:
                        have_decimal = True

                elif element.code == '%':
                    percents += 1
                    need_emit = False
                    tokens.append(self.Token(
                        element.code, self.TokenType.STRING, element.position))

                if need_emit:
                    tokens.append(self.Token(self._get_matching_codes(element, stream),
                                             self.TokenType.NUMBER, element.position))

            elif element.code == '[' and element.next_code in set('hms'):
                element = next(stream)
                tokens.append(self.Token(
                    f'[{self._get_matching_codes(element, stream, eos_allowed=False)[0]}]',
                    self.TokenType.DATETIME, element.position
                ))
                last_date = tokens[-1], len(tokens)
                if next(stream).code != ']':
                    raise ValueError

            elif element.code in DateTimeFormatter.FORMAT_DATETIME_CONVERSION_LOOKUP:
                am_pm = self._find_am_pm(element, format, stream)
                if am_pm is not None:
                    tokens.append(self.Token(am_pm, self.TokenType.AM_PM, element.position))
                elif element.code == 'a':
                    tokens.append(self.Token(
                        element.code, self.TokenType.STRING, element.position))
                else:
                    code = self._get_matching_codes(element, stream)
                    # Search previous actual token not punctuation
                    if code in {'m', 'mm'} and last_date and last_date[0].token[0] in 'hs[':
                        # this is minutes not months
                        code = code.upper()

                    elif code[0] == 's' and last_date and last_date[0].token in {'m', 'mm'}:
                        # the previous minutes not months
                        prev = last_date[1] - 1
                        tokens[prev] = self.Token(
                            tokens[prev].token.upper(), tokens[prev].type, tokens[prev].position)

                    elif code == '.' and element.next_code == '0':
                        # if we are here with '.', then this is subseconds: ss.000
                        code += self._get_matching_codes(next(stream), stream)

                    tokens.append(self.Token(code, self.TokenType.DATETIME, element.position))
                last_date = tokens[-1], len(tokens)

            elif element.code == '*':
                if element.next_code is None:
                    raise ValueError
                # we don't support filling, so drop the character following '*'
                next(stream)

            else:
                tokens.append(self.Token(element.char, self.TokenType.STRING, element.position))

        if tokens:
            yield self._finalize_tokenize(tokens, have_decimal, have_thousands, percents)

    @classmethod
    def _finalize_tokenize(cls, tokens: List[Token], decimal, thousands, percents) -> Tokenized:
        types = {token.type for token in tokens}
        if cls.TokenType.AM_PM in types:
            # replace with the 12 hours version of hours
            tokens = [t if t.token[0] != 'h' else cls.Token(t.token.upper(), t.type, t.position)
                      for t in tokens]
            types.remove(cls.TokenType.AM_PM)
            types.add(cls.TokenType.DATETIME)
        if len(types.intersection(cls.FORMAT_TYPES)) > 1:
            raise ValueError
        return cls.Tokenized(tuple(tokens), frozenset(types), decimal, thousands, percents)

    def format_value(self, data) -> str:
        tokenized_formats = self.tokenized_formats
        if isinstance(tokenized_formats, str):
            return tokenized_formats

        # check for only one string replace field, and in the last field if present
        string_replace_token_count = sum(int(self.TokenType.REPLACE in tokens.types)
                                         for tokens in tokenized_formats)
        if string_replace_token_count and (
                string_replace_token_count > 1 or
                self.TokenType.REPLACE not in tokenized_formats[-1].types):
            return VALUE_ERROR

        # (attempt to) convert the data into a date (serial number) or number
        convertor = DateTimeFormatter.new(data)
        if convertor is not None:
            # The data was a convertable date
            data = convertor.serial_number
        elif data is None:
            data = 0
        else:
            data = coerce_to_number(data)

        # Process strings first
        if isinstance(data, str):
            # '@' is not required in the fourth field to use the field
            if string_replace_token_count or len(tokenized_formats) == 4:
                tokens, token_types = tokenized_formats[-1][:2]
                return ''.join(data if t.type == self.TokenType.REPLACE else t.token
                               for t in tokens)
            else:
                # if no specific string formatter, then pass through
                return data

        if not tokenized_formats:
            return '-' if data < 0 else ''

        if self.TokenType.REPLACE in tokenized_formats[-1].types:
            # remove the string formatter on the end if present
            tokenized_formats = tokenized_formats[:-1]

        if data == 0 and len(tokenized_formats) > 2:
            tokenized_format = tokenized_formats[2]
        elif data < 0 and len(tokenized_formats) > 1:
            tokenized_format = tokenized_formats[1]
        else:
            tokenized_format = tokenized_formats[0]

        if data < 0 and self.TokenType.DATETIME not in tokenized_format.types:
            data = -data
            if len(tokenized_formats) < 2:
                amended_tokens = (
                    self.Token('-', self.TokenType.STRING, -1), *tokenized_format.tokens)
                tokenized_format = self.Tokenized(
                    tokens=amended_tokens,
                    types=tokenized_format.types,
                    decimal=tokenized_format.decimal,
                    thousands=tokenized_format.thousands,
                    percents=tokenized_format.percents,
                )

        format_tokens, format_types = tokenized_format[:2]
        if self.TokenType.DATETIME in format_types:
            if convertor is None:
                convertor = DateTimeFormatter(data)
            tokens = tuple(token.token if token.type == self.TokenType.STRING
                           else convertor.format(token.token)
                           for token in format_tokens)
            if any(t in ERROR_CODES for t in tokens):
                return VALUE_ERROR
            else:
                return ''.join(tokens)
        elif self.TokenType.NUMBER in format_types:
            return self._number_converter(data, tokenized_format)
        else:
            # return the format directly
            return ''.join(t.token for t in tokenized_format.tokens)

    def _number_converter(self, number_value, tokenized: Tokenized):
        number_value *= 100 ** tokenized.percents
        number_format = ''.join(
            t.token for t in tokenized.tokens if t.type == self.TokenType.NUMBER)
        thousands = self.thousands_format if tokenized.thousands else ''

        if tokenized.decimal:
            left_num_format, right_num_format = number_format.split('.', 1)
            decimals = len(right_num_format)
            left_side, right_side = f'{number_value:#{thousands}.{decimals}f}'.split('.')
            right_side = right_side.rstrip('0')
        else:
            left_side = f'{int(round(number_value, 0)):{thousands}}'
            right_side = None
        left_side = left_side.lstrip('0')

        tokens_iter = iter(tokenized.tokens)
        left_side_tokens = tuple(it.takewhile(lambda t: t.token != '.', tokens_iter))
        right_side_tokens = tuple(tokens_iter)

        left = tuple(self._number_token_converter(left_side_tokens, left_side, left_side=True))
        if tokenized.decimal:
            right_side = "".join(self._number_token_converter(right_side_tokens, right_side))
            return f'{"".join(left[::-1])}.{right_side}'
        else:
            return ''.join(left[::-1])

    def _number_token_converter(self, tokens, number, left_side=False):
        digits_iter = iter(number[::-1] if left_side else number)
        result = []
        filler = []
        for token in (tokens[::-1] if left_side else tokens):
            if token.type == self.TokenType.STRING:
                filler.extend(iter(token.token[::-1] if left_side else token.token))
            else:
                result.extend(filler)
                filler = []
                for i in range(len(token.token)):
                    c = next(digits_iter, self.NUMBER_TOKEN_MATCH[token.token[0]])
                    if c is not None:
                        result.append(c)
                        if c not in self.DIGITS:
                            c = next(digits_iter, self.NUMBER_TOKEN_MATCH[token.token[0]])
                            if c is not None:  # pragma: no cover
                                result.append(c)
        result.extend(digits_iter)
        result.extend(filler)
        return result


# def asc(text):
# Excel reference: https://support.microsoft.com/en-us/office/
#   asc-function-0b6abf1c-c663-4004-a964-ebc00b723266


# def bahttext(text):
# Excel reference: https://support.microsoft.com/en-us/office/
#   bahttext-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c


# def char(text):
# Excel reference: https://support.microsoft.com/en-us/office/
#   char-function-bbd249c8-b36e-4a91-8017-1c133f9b837a


# def clean(text):
# Excel reference: https://support.microsoft.com/en-us/office/
#   clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41


# def code(text):
# Excel reference: https://support.microsoft.com/en-us/office/
#   code-function-c32b692b-2ed0-4a04-bdd9-75640144b928


def concat(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2
    return concatenate(*tuple(flatten(args)))


def concatenate(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   CONCATENATE-function-8F8AE884-2CA8-4F7A-B093-75D702BEA31D
    if tuple(flatten(args)) != args:
        return VALUE_ERROR

    error = next((x for x in args if x in ERROR_CODES), None)
    if error:
        return error

    return ''.join(coerce_to_string(a) for a in args)


# def dbcs(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   dbcs-function-a4025e73-63d2-4958-9423-21a24794c9e5


# def dollar(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   dollar-function-a6cd05d9-9740-4ad3-a469-8109d18ff611


# def exact(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   exact-function-d3087698-fc15-4a15-9631-12575cf29926


@excel_helper(cse_params=(0, 1, 2), number_params=2, str_params=(0, 1))
def find(find_text, within_text, start_num=1):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   FIND-FINDB-functions-C7912941-AF2A-4BDF-A553-D0D89B0A0628
    found = within_text.find(find_text, start_num - 1)
    if found == -1:
        return VALUE_ERROR
    else:
        return found + 1


# def findb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628


# def fixed(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   fixed-function-ffd5723c-324c-45e9-8b96-e41be2a8274a


# def jis(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   jis-function-b72fb1a7-ba52-448a-b7d3-d2610868b7e2


@excel_helper(cse_params=(0, 1), number_params=1, str_params=0)
def left(text, num_chars=1):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   LEFT-LEFTB-functions-9203D2D2-7960-479B-84C6-1EA52B99640C
    if num_chars < 0:
        return VALUE_ERROR
    else:
        return str(text)[:int(num_chars)]


# def leftb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c


@excel_helper(cse_params=0)
def len_(arg):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb
    return 0 if arg is None else len(str(arg))


# def lenb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb


@excel_helper(cse_params=0, str_params=0)
def lower(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4
    return text.lower()


@excel_helper(cse_params=-1, number_params=(1, 2), str_params=0)
def mid(text, start_num, num_chars):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028

    if start_num < 1 or num_chars < 0:
        return VALUE_ERROR

    start_num = int(start_num) - 1

    return str(text)[start_num:start_num + int(num_chars)]


# def midb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028


# def numbervalue(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   numbervalue-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879


# def phonetic(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   phonetic-function-9a329dac-0c0f-42f8-9a55-639086988554


# def proper(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   proper-function-52a5a283-e8b2-49be-8506-b2887b889f94


@excel_helper(cse_params=-1, number_params=(1, 2), str_params=(0, 3))
def replace(old_text, start_num, num_chars, new_text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   replace-replaceb-functions-8d799074-2425-4a8a-84bc-82472868878a
    start_num = int(start_num) - 1
    num_chars = int(num_chars)
    if start_num < 0 or num_chars < 0:
        return VALUE_ERROR
    return f'{old_text[:start_num]}{new_text}{old_text[start_num + num_chars:]}'


# def replaceb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   replace-replaceb-functions-8d799074-2425-4a8a-84bc-82472868878a


# def rept(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   rept-function-04c4d778-e712-43b4-9c15-d656582bb061


@excel_helper(cse_params=(0, 1), number_params=1, str_params=0)
def right(text, num_chars=1):
    # Excel reference:  https://support.microsoft.com/en-us/office/
    #   RIGHT-RIGHTB-functions-240267EE-9AFA-4639-A02B-F19E1786CF2F

    if num_chars < 0:
        return VALUE_ERROR
    elif num_chars == 0:
        return ''
    else:
        return str(text)[-int(num_chars):]


# def rightb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f


# def search(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495


# def searchb(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495


@excel_helper(cse_params=-1, str_params=(0, 1, 2))
def substitute(text, old_text, new_text, instance_num=None):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332
    if instance_num is None:
        return text.replace(old_text, new_text)

    if isinstance(instance_num, bool):
        return VALUE_ERROR

    try:
        instance_num = int(instance_num)
    except ValueError:
        return VALUE_ERROR

    if instance_num <= 0:
        return VALUE_ERROR

    start = 0
    while instance_num > 1:
        new_start = text[start:].find(old_text)
        if new_start == -1:
            return text
        instance_num -= 1
        start += new_start + len(old_text)
    replaced = text[start:].replace(old_text, new_text, 1)
    return f'{text[:start]}{replaced}'


# def t(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-function-fb83aeec-45e7-4924-af95-53e073541228


@excel_helper(cse_params=0, str_params=1)
def text(text_value, value_format):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c
    if isinstance(text_value, bool):
        text_value = 'TRUE' if text_value else 'FALSE'
    return TextFormat(value_format).format_value(text_value)


# def textjoin(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c


@excel_helper(cse_params=0, str_params=0)
def trim(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9
    return RE_MULTI_SPACE.sub(' ', text)


# def unichar(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   unichar-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8


# def unicode(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   unicode-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f


@excel_helper(cse_params=0, str_params=0)
def upper(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6
    return text.upper()


@excel_helper(cse_params=0)
def value(text):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   VALUE-function-257D0108-07DC-437D-AE1C-BC2D3953D8C2
    if isinstance(text, bool):
        return VALUE_ERROR
    if text is None:
        return 0
    try:
        return float(text)
    except ValueError:
        return VALUE_ERROR


# Older mappings for excel functions that match Python built-in and keywords
x_len = len_
