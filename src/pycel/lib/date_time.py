# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of Date and Time library functions
"""

import calendar
import datetime as dt
import functools
import math

import dateutil.parser

from pycel.excelutil import (
    coerce_to_number,
    ERROR_CODES,
    is_number,
    NUM_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import (
    excel_helper,
)


DATE_ZERO = dt.datetime(1899, 12, 30)
DATE_MAX = dt.datetime(9999, 12, 31)  # last legal value
DATE_MAX_INT = (DATE_MAX - DATE_ZERO).days + 1  # first illegal value
SECOND = 1 / 24 / 60 / 60
MICROSECOND = SECOND / 1E6
LEAP_1900_SERIAL_NUMBER = 60  # magic number for non-existent 1900/02/29
LEAP_1900_TUPLE = 1900, 2, 29

TIME_CHARS = set('0123456789')
SECS_CHARS = TIME_CHARS | {'.'}


def serial_number_wrapper(f):
    """Validations and conversions for date-time serial numbers"""
    @functools.wraps(f)
    @excel_helper(number_params=0)
    def wrapped(date_serial_number):
        if date_serial_number < 0:
            return NUM_ERROR
        return f(date_serial_number)
    return wrapped


def time_value_wrapper(f):
    """Validations and conversions for date values"""
    @functools.wraps(f)
    def wrapped(a_timevalue):
        if isinstance(a_timevalue, str):
            try:
                a_timevalue = float(a_timevalue)
            except ValueError:
                a_timevalue = timevalue(a_timevalue)
            if a_timevalue in ERROR_CODES:
                return a_timevalue
        if a_timevalue is None:
            a_timevalue = 0
        elif a_timevalue < 0:
            return NUM_ERROR
        return f(a_timevalue)
    return wrapped


def date_from_int(datestamp):

    if datestamp == LEAP_1900_SERIAL_NUMBER:
        # excel thinks 1900 is a leap year
        return LEAP_1900_TUPLE

    if datestamp == 0:
        # excel thinks Jan 1900 starts at day 0
        return 1900, 1, 0

    date = DATE_ZERO + dt.timedelta(days=datestamp)
    if datestamp < LEAP_1900_SERIAL_NUMBER:
        date += dt.timedelta(days=1)

    return date.year, date.month, date.day


def time_from_serialnumber_with_microseconds(serialnumber):
    at_hours = (serialnumber + MICROSECOND / 1.5) * 24
    hours = math.floor(at_hours)
    at_mins = (at_hours - hours) * 60
    mins = math.floor(at_mins)
    at_secs = (at_mins - mins) * 60
    secs = math.floor(at_secs)
    microseconds = (at_secs - secs) * 1E6
    return hours % 24, mins, secs, int(microseconds - 0.5)


def time_from_serialnumber(serialnumber):
    at_hours = (serialnumber + MICROSECOND) * 24
    hours = math.floor(at_hours)
    at_mins = (at_hours - hours) * 60
    mins = math.floor(at_mins)
    secs = (at_mins - mins) * 60
    return hours % 24, mins, int(round(secs - 1.1E-6, 0))


def is_leap_year(year):
    if not is_number(year):
        raise TypeError(f"{year} must be a number")
    if year <= 0:
        raise TypeError(f"{year} must be strictly positive")

    # Watch out, 1900 is a leap according to Excel =>
    # https://support.microsoft.com/en-us/kb/214326
    return year % 4 == 0 and year % 100 != 0 or year % 400 == 0 or year == 1900


def max_days_in_month(month, year):
    if month == 2 and is_leap_year(year):
        return 29

    return calendar.monthrange(year, month)[1]


def normalize_year(y, m, d):
    """taking into account negative month and day values"""
    if not (1 <= m <= 12):
        y_plus = math.floor((m - 1) / 12)
        y += y_plus
        m -= y_plus * 12

    if d <= 0:
        d += max_days_in_month(m, y)
        m -= 1
        y, m, d = normalize_year(y, m, d)

    else:
        days_in_month = max_days_in_month(m, y)
        if d > days_in_month:
            m += 1
            d -= days_in_month
            y, m, d = normalize_year(y, m, d)

    return y, m, d


def yearfrac_basis_0(beg, end):
    # https://github.com/dgorissen/pycel/issues/111
    y1, m1, d1 = beg
    y2, m2, d2 = end

    # Change day-of-month for purposes of calculation.
    if d1 == 31:
        d1 = 30
        if d2 == 31:
            d2 = 30

    elif d1 == 30 and d2 == 31:
        # Note: If d2==31, it STAYS 31 if d1 < 30.
        d2 = 30

    # Special fixes for February:
    elif m1 == 2 and d1 == calendar.monthrange(y1, m1)[1]:
        d1 = 30
        if m2 == 2 and d2 == calendar.monthrange(y2, m2)[1]:
            d2 = 30

    return ((d2 + m2 * 30 + y2 * 360) - (d1 + m1 * 30 + y1 * 360)) / 360


def yearfrac_basis_1(beg, end):
    # http://svn.finmath.net/finmath%20lib/trunk/src/main/java/net/
    #   finmath/time/daycount/DayCountConvention_ACT_ACT_YEARFRAC.java
    delta = date(*end) - date(*beg)

    if delta <= 365:
        if (is_leap_year(beg[0]) and date(*beg) <= date(beg[0], 2, 29) or
            is_leap_year(end[0]) and date(*end) >= date(end[0], 2, 29) or
                is_leap_year(beg[0]) and is_leap_year(end[0])):
            denom = 366
        else:
            denom = 365
    else:
        year_range = range(beg[0], end[0] + 1)
        nb = 0

        for y in year_range:
            nb += 366 if is_leap_year(y) else 365

        denom = nb / len(year_range)

    return delta / denom


class DateTimeFormatter:
    """Using the Excel Formatting language, format a date time, one token at a time

    class TextFormat contains code to tokenize the format string
    """

    # Excel reference: https://support.microsoft.com/en-us/office/
    # review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5

    FORMAT_DATETIME_CONVERSIONS = {
        'yyyy': lambda d: d._strftime('%Y'),
        'yy': lambda d: d._strftime('%y'),
        'mmmmm': lambda d: d._strftime('%b')[0],
        'mmmm': lambda d: d._strftime('%B'),
        'mmm': lambda d: d._strftime('%b'),
        'mm': lambda d: "{:02d}".format(d.month),
        'm': lambda d: str(d.month),
        'dddd': lambda d: d._strftime('%A'),
        'ddd': lambda d: d._strftime('%a'),
        'dd': lambda d: "{:02d}".format(d.day),
        'd': lambda d: str(d.day),
        'hh': lambda d: "{:02d}".format(d.hour),
        'h': lambda d: str(d.hour),
        'HH': lambda d: d._strftime('%I'),  # 12 Hour (AM/PM)
        'H': lambda d: str(int(d._strftime('%I'))),  # 12 Hour (AM/PM)
        'MM': lambda d: "{:02d}".format(d.minute),
        'M': lambda d: str(d.minute),
        'ss': lambda d: "{:02d}".format(d.second),
        's': lambda d: str(d.second),
        '.': lambda d: ".",
        '.0': lambda d: ".{:01d}".format(round(d.microsecond / 100000)),
        '.00': lambda d: ".{:02d}".format(round(d.microsecond / 10000)),
        '.000': lambda d: ".{:03d}".format(round(d.microsecond / 1000)),
        '[h]': lambda d: str(d._elapsed('h')),
        '[m]': lambda d: str(d._elapsed('m')),
        '[s]': lambda d: str(d._elapsed('s')),
        'am/pm': lambda d: d._strftime('%p'),
        'a/p': lambda d: d._strftime('%p')[0].lower(),
        'A/P': lambda d: d._strftime('%p')[0].upper(),
        'A/p': lambda d: 'A' if d._strftime('%p').lower() == 'am' else 'p',
        'a/P': lambda d: 'a' if d._strftime('%p').lower() == 'am' else 'P',
    }

    def FORMAT_DATETIME_CONVERSION_LOOKUP(FORMAT_DATETIME_CONVERSIONS):
        return {
            'e': lambda code: FORMAT_DATETIME_CONVERSIONS['yyyy'],
            'y': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 'yy',
                2: 'yy'
            }.get(len(code), 'yyyy')],
            'm': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 'm',
                2: 'mm',
                3: 'mmm',
                4: 'mmmm',
                5: 'mmmmm',
            }.get(len(code), 'mmmm')],
            'd': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 'd',
                2: 'dd',
                3: 'ddd',
            }.get(len(code), 'dddd')],
            'h': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 'h',
            }.get(len(code), 'hh')],
            'H': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 'H',
            }.get(len(code), 'HH')],
            'M': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 'M',
            }.get(len(code), 'MM')],
            's': lambda code: FORMAT_DATETIME_CONVERSIONS[{
                1: 's',
            }.get(len(code), 'ss')],
            '.': lambda code: FORMAT_DATETIME_CONVERSIONS[code],
            'a': lambda code: FORMAT_DATETIME_CONVERSIONS[code],
            'A': lambda code: FORMAT_DATETIME_CONVERSIONS[code],
            '[': lambda code: FORMAT_DATETIME_CONVERSIONS[code],
        }
    FORMAT_DATETIME_CONVERSION_LOOKUP = FORMAT_DATETIME_CONVERSION_LOOKUP(
        FORMAT_DATETIME_CONVERSIONS)

    def format(self, format_str):
        """Format datetime using a single token from a custom format"""
        try:
            return self.FORMAT_DATETIME_CONVERSION_LOOKUP[format_str[0]](format_str)(self)
        except (KeyError, ValueError, AttributeError):
            return VALUE_ERROR

    def __init__(self, serial_number, time=None):
        """Init formatter using a datetime serial number

        Use the .new() method to init from a date time string or serial number

        :param serial_number: An Excel datetime serial number
        :param time: An optional datetime.time object instance
        """
        if 0 <= serial_number < DATE_MAX_INT:
            # only if the serial number is not OOR can we do date conversion
            self.time = time
            datestamp = int(serial_number)
            self.year, self.month, self.day = date_from_int(datestamp)
            if time is None:
                self.hour, self.minute, self.second, self.microsecond = \
                    time_from_serialnumber_with_microseconds(serial_number)
            else:
                assert serial_number == datestamp
                serial_number += dt.timedelta(
                    hours=time.hour, minutes=time.minute, seconds=time.second
                ).total_seconds() * SECOND
                self.hour = time.hour
                self.minute = time.minute
                self.second = time.second
                self.microsecond = time.microsecond

        self.serial_number = serial_number
        self._cached_datetime = None

    @classmethod
    def new(cls, excel_date_time):
        """Create a cls instance if the parameter is convertible to an excel date time

        :param excel_date_time: An excel datatype that might be a valid date/time
        :return: cls instance if convertible, else None
        """
        if isinstance(excel_date_time, bool):
            return None

        try:
            serial_number = float(excel_date_time)
            time = None
        except (ValueError, TypeError):
            if not isinstance(excel_date_time, str):
                return None

            try:
                time = dateutil.parser.parse(
                    excel_date_time, parserinfo=DateutilParserInfo()).time()
                serial_number = datevalue(excel_date_time)
            except (TypeError, ValueError):
                # if we get here, then dateutil can't parse date, try for just a time
                time = None
                serial_number = timevalue(excel_date_time)

            if isinstance(serial_number, str):
                # failed to convert
                return None

        if serial_number < 0 or DATE_MAX_INT <= serial_number:
            return None
        return cls(serial_number, time)

    def _strftime(self, format):
        return self._datetime.strftime(format)

    @property
    def _datetime(self):
        if self._cached_datetime is None:
            try:
                self._cached_datetime = dt.datetime(
                    self.year, self.month, self.day, self.hour,
                    self.minute, self.second, self.microsecond)
            except ValueError:
                if (self.year, self.month, self.day) == (1900, 1, 0):
                    # preserve day of the week for 1900-01-00
                    self._cached_datetime = dt.datetime(
                        self.year, self.month, self.day + 7, self.hour,
                        self.minute, self.second, self.microsecond)
                elif (self.year, self.month, self.day) == (1900, 2, 29):
                    # preserve day of the week for 1900-02-29
                    self._cached_datetime = dt.datetime(
                        self.year, self.month, self.day - 7, self.hour,
                        self.minute, self.second, self.microsecond)
                else:  # pragma: no cover
                    # this should not be a possibility
                    assert False
        return self._cached_datetime

    def _elapsed(self, units):
        if hasattr(self, 'hour'):
            elapsed = int(self.serial_number) * 24 + self.hour
            if units == 'm':
                elapsed = elapsed * 60 + self.minute
            elif units == 's':
                elapsed = (elapsed * 60 + self.minute) * 60 + self.second
        else:
            elapsed = self.serial_number * 24
            if units == 'm':
                elapsed *= 60
            elif units == 's':
                elapsed *= 60 * 60

        return int(elapsed)


@excel_helper(number_params=-1)
def date(year, month_, day):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

    if not (0 <= year <= 9999):
        return NUM_ERROR

    if year < 1900:
        year += 1900

    # taking into account negative month and day values
    year, month_, day = normalize_year(year, month_, day)

    try:
        result = (dt.datetime(year, month_, day) - DATE_ZERO).days
        if result <= 60:
            result -= 1
    except ValueError:
        assert (year, month_, day) == LEAP_1900_TUPLE
        result = 60.0

    if result < 0:
        return NUM_ERROR
    return result


# def datedif(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c


class DateutilParserInfo(dateutil.parser.parserinfo):
    """Hook into dateutil parser and fix number strings and 1900/02/29"""

    def __init__(self):
        super().__init__()
        self.is_leap_day_1900 = False

    def validate(self, res):
        if res.day is None or res.month is None:
            return False
        if (res.year, res.month, res.day) == LEAP_1900_TUPLE:
            self.is_leap_day_1900 = True
        return super().validate(res)


def datevalue(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   datevalue-function-df8b07d4-7761-4a93-bc33-b7471bbff252
    parserinfo = DateutilParserInfo()
    try:
        a_date = dateutil.parser.parse(value, parserinfo=parserinfo).date()
    except (TypeError, ValueError):
        if parserinfo.is_leap_day_1900:
            return LEAP_1900_SERIAL_NUMBER
        elif value in ERROR_CODES:
            return value
        else:
            return VALUE_ERROR

    serial_number = (a_date - DATE_ZERO.date()).days
    if serial_number <= LEAP_1900_SERIAL_NUMBER:
        serial_number -= 1
        if serial_number < 1:
            return VALUE_ERROR
    return serial_number


@serial_number_wrapper
def day(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101
    return date_from_int(math.floor(serial_number))[2]


# def days(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   days-function-57740535-d549-4395-8728-0f07bff0b9df


# def days360(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   days360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a


@excel_helper(err_str_params=-1)
def edate(start_date, months):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5
    return months_inc(start_date, months)


@excel_helper(err_str_params=-1)
def eomonth(start_date, months):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628
    return months_inc(start_date, months, eomonth=True)


def months_inc(start_date, months, eomonth=False):
    if isinstance(start_date, bool) or isinstance(months, bool):
        return VALUE_ERROR
    start_date = coerce_to_number(start_date, convert_all=True)
    months = coerce_to_number(months, convert_all=True)
    if isinstance(start_date, str) or isinstance(months, str):
        return VALUE_ERROR
    if start_date < 0:
        return NUM_ERROR
    y, m, d = date_from_int(start_date)
    if eomonth:
        return date(y, m + months + 1, 1) - 1
    else:
        return date(y, m + months, d)


@time_value_wrapper
def hour(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   hour-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7
    return time_from_serialnumber(serial_number)[0]


# def isoweeknum(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   isoweeknum-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e


@time_value_wrapper
def minute(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   minute-function-af728df0-05c4-4b07-9eed-a84801a60589
    return time_from_serialnumber(serial_number)[1]


@serial_number_wrapper
def month(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   month-function-579a2881-199b-48b2-ab90-ddba0eba86e8
    return date_from_int(math.floor(serial_number))[1]


# def networkdays(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7


# def networkdays.intl(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28


def now():
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   now-function-3337fd29-145a-4347-b2e6-20c904739c46
    delta = dt.datetime.now() - DATE_ZERO
    return delta.days + delta.seconds * SECOND


@time_value_wrapper
def second(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   second-function-740d1cfc-553c-4099-b668-80eaa24e8af1
    return time_from_serialnumber(serial_number)[2]


# def time(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   time-function-9a5aff99-8f7d-4611-845e-747d0b8d5457


def timevalue(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   timevalue-function-0b615c12-33d8-4431-bf3d-f3eb6d186645
    if not isinstance(value, str):
        return VALUE_ERROR

    if value in ERROR_CODES:
        return value

    fields = value.lower().replace(':', ' ').split()
    colons = value.count(':')
    have_secs = True
    if colons == 1:
        if '.' in fields[1][:-1]:
            # a decimal is seconds
            fields.insert(0, '0')
        else:
            if fields[1][-1] == '.':
                fields[1] = fields[1][:-1]
            fields.insert(2, '0')
            have_secs = False
    elif colons != 2:
        return VALUE_ERROR

    # validate characters present
    if set(fields[0]) - TIME_CHARS or set(fields[1]) - TIME_CHARS or set(fields[2]) - SECS_CHARS:
        return VALUE_ERROR

    try:
        time_tuple = list(map(float, fields[:3]))
    except ValueError:
        return VALUE_ERROR
    if time_tuple[0] > 23 or \
            time_tuple[1] > (59 if have_secs else 9999) or \
            time_tuple[2] >= 10000:
        return VALUE_ERROR
    if time_tuple[0] == 12 and len(fields) == 4:
        time_tuple[0] = 0
    serial_number = ((
        time_tuple[0] * 60 + time_tuple[1]) * 60 + time_tuple[2]) / 86400

    if len(fields) == 4:
        if fields[3][0] == 'p':
            serial_number += 0.5
        elif fields[3][0] != 'a':
            return VALUE_ERROR

    return serial_number


def today():
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
    return (dt.date.today() - DATE_ZERO.date()).days


@serial_number_wrapper
def weekday(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a
    return (math.floor(serial_number) - 1) % 7 + 1


# def weeknum(serial_number, return_Type=1):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   weeknum-function-e5c43a03-b4ab-426c-b411-b18c13c75340


# def workday(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   workday-function-f764a5b7-05fc-4494-9486-60d494efbf33


# def workday.intl(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d


@serial_number_wrapper
def year(serial_number):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   year-function-c64f017a-1354-490d-981f-578e8ec8d3b9
    return date_from_int(math.floor(serial_number))[0]


@excel_helper(cse_params=-1, err_str_params=2, number_params=None)
def yearfrac(start_date, end_date, basis=0):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8
    if isinstance(basis, (bool, str)):
        return VALUE_ERROR
    basis = 0 if basis is None else int(basis)
    if basis not in {0, 1, 2, 3, 4}:
        return NUM_ERROR

    if start_date in ERROR_CODES:
        return start_date

    if end_date in ERROR_CODES:
        return end_date

    try:
        if not (0 <= start_date < DATE_MAX_INT and 0 <= end_date < DATE_MAX_INT):
            return NUM_ERROR
    except TypeError:
        return VALUE_ERROR

    if start_date > end_date:  # switch dates if start_date > end_date
        start_date, end_date = end_date, start_date

    y1, m1, d1 = date_from_int(start_date)
    y2, m2, d2 = date_from_int(end_date)

    if basis == 0:  # US 30/360
        result = yearfrac_basis_0((y1, m1, d1), (y2, m2, d2))

    elif basis == 1:  # Actual/actual
        result = yearfrac_basis_1((y1, m1, d1), (y2, m2, d2))

    elif basis == 2:  # Actual/360
        result = (end_date - start_date) / 360

    elif basis == 3:  # Actual/365
        result = (end_date - start_date) / 365

    else:  # basis == 4:  # Eurobond 30/360
        d2 = min(d2, 30)
        d1 = min(d1, 30)

        day_count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = day_count / 360

    return result
