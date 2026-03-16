# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of Excel financial functions
"""
import numpy_financial as npf

from pycel.excelutil import flatten


def irr(values, guess=None):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc

    # currently guess is not used
    return npf.irr(list(flatten(values)))


def pmt(rate, nper, pv, fv=0, when=0):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   pmt-function-0214da64-9a63-4996-bc20-214433fa6441
    return npf.pmt(rate, nper, pv, fv=fv, when=when)


def ppmt(rate, per, nper, pv, fv=0, when=0):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   ppmt-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b
    return npf.ppmt(rate, per, nper, pv, fv=fv, when=when)
