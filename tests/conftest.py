import os
import shutil

import pytest
from unittest import mock

from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl


@pytest.fixture('session')
def fixture_dir():
    return os.path.dirname(__file__)


@pytest.fixture('session')
def tmpdir(tmpdir_factory):
    return tmpdir_factory.mktemp('fixtures')


@pytest.fixture('session')
def example_xls_path(fixture_dir, tmpdir):
    src = os.path.join(fixture_dir, "fixtures/excelcompiler.xlsx")
    dst = os.path.join(str(tmpdir), "excelcompiler.xlsx")
    shutil.copy(src, dst)
    return dst


@pytest.fixture('session')
def unconnected_excel(example_xls_path):
    import openpyxl.reader.worksheet as orw
    old_warn = orw.warn

    def new_warn(msg, *args, **kwargs):
        if 'Unknown' not in msg:
            old_warn(msg, *args, **kwargs)
            
    with mock.patch('openpyxl.reader.worksheet.warn', new_warn):
        yield ExcelWrapperImpl(example_xls_path)
        
        
@pytest.fixture()
def excel(unconnected_excel):
    unconnected_excel.connect()
    return unconnected_excel
