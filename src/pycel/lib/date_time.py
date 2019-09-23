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
    excel_math_func,
)


DATE_ZERO = dt.datetime(1899, 12, 30)
SECOND = 1 / 24 / 60 / 60
MICROSECOND = SECOND / 1E6
LEAP_1900_SERIAL_NUMBER = 60  # magic number for non-existent 1900/02/29


def serial_number_wrapper(f):
    """Validations and conversions for date-time serial numbers"""
    @functools.wraps(f)
    @excel_helper(number_params=0)
    def wrapped(date_serial_number, *args, **kwargs):
        if date_serial_number < 0:
            return NUM_ERROR
        return f(date_serial_number, *args, **kwargs)
    return wrapped


def time_value_wrapper(f):
    """Validations and conversions for date values"""
    @functools.wraps(f)
    def wrapped(a_timevalue, *args, **kwargs):
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
        return 1900, 2, 29

    if datestamp == 0:
        # excel thinks Jan 1900 starts at day 0
        return 1900, 1, 0

    date = DATE_ZERO + dt.timedelta(days=datestamp)
    if datestamp < LEAP_1900_SERIAL_NUMBER:
        date += dt.timedelta(days=1)

    return date.year, date.month, date.day


def time_from_serialnumber(serialnumber):
    at_hours = (serialnumber + MICROSECOND) * 24
    hours = math.floor(at_hours)
    at_mins = (at_hours - hours) * 60
    mins = math.floor(at_mins)
    secs = (at_mins - mins) * 60
    return hours % 24, mins, int(round(secs - 1.1E-6, 0))


def is_leap_year(year):
    if not is_number(year):
        raise TypeError("%s must be a number" % str(year))
    if year <= 0:
        raise TypeError("%s must be strictly positive" % str(year))

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


@excel_helper(number_params=-1)
def date(year, month_, day):
    # Excel reference: https://support.office.com/en-us/article/
    #   DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

    if not (0 <= year <= 9999):
        return NUM_ERROR

    if year < 1900:
        year += 1900

    # taking into account negative month and day values
    year, month_, day = normalize_year(year, month_, day)

    result = (dt.datetime(year, month_, day) - DATE_ZERO).days

    if result <= 0:
        return NUM_ERROR
    return result


# def datedif(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c


class DateutilParserInfo(dateutil.parser.parserinfo):
    """Hook into dateutil parser and fix number strings and 1900/01/29"""

    def __init__(self):
        super().__init__()
        self.is_leap_day_1900 = False

    def validate(self, res):
        if res.day is None or res.month is None:
            return False
        if (res.year, res.month, res.day) == (1900, 2, 29):
            self.is_leap_day_1900 = True
        return super().validate(res)


def datevalue(value):
    # Excel reference: https://support.office.com/en-us/article/
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
    # Excel reference: https://support.office.com/en-us/article/
    #   day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101
    return date_from_int(math.floor(serial_number))[2]


# def days(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   days-function-57740535-d549-4395-8728-0f07bff0b9df


# def days360(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   days360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a


@excel_helper(err_str_params=-1)
def edate(start_date, months):
    # Excel reference: https://support.office.com/en-us/article/
    #   edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5
    return months_inc(start_date, months)


@excel_helper(err_str_params=-1)
def eomonth(start_date, months):
    # Excel reference: https://support.office.com/en-us/article/
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
    # Excel reference: https://support.office.com/en-us/article/
    #   hour-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7
    return time_from_serialnumber(serial_number)[0]


# def isoweeknum(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   isoweeknum-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e


@time_value_wrapper
def minute(serial_number):
    # Excel reference: https://support.office.com/en-us/article/
    #   minute-function-af728df0-05c4-4b07-9eed-a84801a60589
    return time_from_serialnumber(serial_number)[1]


@serial_number_wrapper
def month(serial_number):
    # Excel reference: https://support.office.com/en-us/article/
    #   month-function-579a2881-199b-48b2-ab90-ddba0eba86e8
    return date_from_int(math.floor(serial_number))[1]


# def networkdays(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7


# def networkdays.intl(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28


def now():
    # Excel reference: https://support.office.com/en-us/article/
    #   now-function-3337fd29-145a-4347-b2e6-20c904739c46
    delta = dt.datetime.now() - DATE_ZERO
    return delta.days + delta.seconds * SECOND


@time_value_wrapper
def second(serial_number):
    # Excel reference: https://support.office.com/en-us/article/
    #   second-function-740d1cfc-553c-4099-b668-80eaa24e8af1
    return time_from_serialnumber(serial_number)[2]


# def time(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   time-function-9a5aff99-8f7d-4611-845e-747d0b8d5457


def timevalue(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   timevalue-function-0b615c12-33d8-4431-bf3d-f3eb6d186645
    if not isinstance(value, str):
        return VALUE_ERROR

    if value in ERROR_CODES:
        return value

    fields = value.lower().replace(':', ' ').split()
    colons = value.count(':')
    if colons == 1:
        fields.insert(2, 0)
    elif colons != 2:
        return VALUE_ERROR

    try:
        time_tuple = list(map(int, fields[:3]))
        if time_tuple[0] == 12 and len(fields) == 4:
            time_tuple[0] = 0
        serial_number = ((
            time_tuple[0] * 60 + time_tuple[1]) * 60 + time_tuple[2]) / 86400
    except ValueError:
        return VALUE_ERROR

    if len(fields) == 4:
        if fields[3][0] == 'p':
            serial_number += 0.5
        elif fields[3][0] != 'a':
            return VALUE_ERROR

    return serial_number


def today():
    # Excel reference: https://support.office.com/en-us/article/
    #   today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
    return (dt.date.today() - DATE_ZERO.date()).days


@serial_number_wrapper
def weekday(serial_number):
    # Excel reference: https://support.office.com/en-us/article/
    #   weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a
    return (math.floor(serial_number) - 1) % 7 + 1


# def weeknum(serial_number, return_Type=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   weeknum-function-e5c43a03-b4ab-426c-b411-b18c13c75340


# def workday(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   workday-function-f764a5b7-05fc-4494-9486-60d494efbf33


# def workday.intl(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d


@serial_number_wrapper
def year(serial_number):
    # Excel reference: https://support.office.com/en-us/article/
    #   year-function-c64f017a-1354-490d-981f-578e8ec8d3b9
    return date_from_int(math.floor(serial_number))[0]


@excel_math_func
def yearfrac(start_date, end_date, basis=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8

    def actual_nb_days_afb_alter(beg, end):
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

    if start_date < 0 or end_date < 0:
        return NUM_ERROR

    if start_date > end_date:  # switch dates if start_date > end_date
        start_date, end_date = end_date, start_date

    y1, m1, d1 = date_from_int(start_date)
    y2, m2, d2 = date_from_int(end_date)

    if basis == 0:  # US 30/360
        d1 = min(d1, 30)
        d2 = max(d2, 30) if d1 == 30 else d2

        day_count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = day_count / 360

    elif basis == 1:  # Actual/actual
        result = actual_nb_days_afb_alter((y1, m1, d1), (y2, m2, d2))

    elif basis == 2:  # Actual/360
        result = (end_date - start_date) / 360

    elif basis == 3:  # Actual/365
        result = (end_date - start_date) / 365

    elif basis == 4:  # Eurobond 30/360
        d2 = min(d2, 30)
        d1 = min(d1, 30)

        day_count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = day_count / 360

    else:
        return NUM_ERROR

    return result
