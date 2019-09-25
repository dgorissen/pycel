# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import os
import shutil
from unittest import mock

import pytest
from openpyxl.utils import column_index_from_string

from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import AddressCell
from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl


@pytest.fixture('session')
def ATestCell():

    class ATestCell:

        def __init__(self, col, row, sheet='', excel=None, value=None):
            self.row = row
            self.col = col
            self.col_idx = column_index_from_string(col)
            self.sheet = sheet
            self.excel = excel
            self.address = AddressCell('{}{}'.format(col, row), sheet=sheet)
            self.value = value

    return ATestCell


@pytest.fixture('session')
def fixture_dir():
    return os.path.join(os.path.dirname(__file__), 'fixtures')


@pytest.fixture('session')
def tmpdir(tmpdir_factory):
    return tmpdir_factory.mktemp('fixtures')


@pytest.fixture('session')
def serialization_override_path(tmpdir):
    return os.path.join(str(tmpdir), 'excelcompiler_serialized.yml')


def copy_fixture_xls_path(fixture_dir, tmpdir, filename):
    src = os.path.join(fixture_dir, filename)
    dst = os.path.join(str(tmpdir), filename)
    shutil.copy(src, dst)
    return dst


@pytest.fixture('session')
def fixture_xls_copy(fixture_dir, tmpdir):
    def wrapped(filename):
        return copy_fixture_xls_path(fixture_dir, tmpdir, filename)
    return wrapped


@pytest.fixture('session')
def fixture_xls_path(fixture_xls_copy):
    return fixture_xls_copy('excelcompiler.xlsx')


@pytest.fixture('session')
def fixture_xls_path_circular(fixture_xls_copy):
    return fixture_xls_copy('circular.xlsx')


@pytest.fixture('session')
def unconnected_excel(fixture_xls_path):
    import openpyxl.worksheet._reader as orw
    old_warn = orw.warn

    def new_warn(msg, *args, **kwargs):
        if 'Unknown' not in msg:
            old_warn(msg, *args, **kwargs)

    # quiet the warnings about unknown extensions
    with mock.patch('openpyxl.worksheet._reader.warn', new_warn):
        yield ExcelWrapperImpl(fixture_xls_path)


@pytest.fixture()
def excel(unconnected_excel):
    unconnected_excel.load()
    return unconnected_excel


@pytest.fixture('session')
def basic_ws(fixture_xls_copy):
    return ExcelCompiler(fixture_xls_copy('basic.xlsx'))


@pytest.fixture('session')
def cond_format_ws(fixture_xls_copy):
    return ExcelCompiler(fixture_xls_copy('cond-format.xlsx'))


@pytest.fixture
def circular_ws(fixture_xls_path_circular):
    return ExcelCompiler(fixture_xls_path_circular, cycles=True)


@pytest.fixture
def excel_compiler(excel):
    return ExcelCompiler(excel=excel)
