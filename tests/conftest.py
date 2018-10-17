import os
import pytest
from unittest import mock

from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl


@pytest.fixture('session')
def fixture_dir():
    return os.path.dirname(__file__)


@pytest.fixture('session')
def unconnected_excel(fixture_dir):
    import openpyxl.reader.worksheet as orw
    old_warn = orw.warn

    def new_warn(msg, *args, **kwargs):
        if 'Unknown' not in msg:
            old_warn(msg, *args, **kwargs)
            
    with mock.patch('openpyxl.reader.worksheet.warn', new_warn):
        yield ExcelWrapperImpl(
            os.path.join(fixture_dir, "../example/example.xlsx"))
        
        
@pytest.fixture()
def excel(unconnected_excel):
    unconnected_excel.connect()
    return unconnected_excel
