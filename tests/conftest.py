import os
import shutil

import pytest
from unittest import mock

from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl
from pycel.excelcompiler import ExcelCompiler


@pytest.fixture('session')
def fixture_dir():
    return os.path.join(os.path.dirname(__file__), 'fixtures')


@pytest.fixture('session')
def tmpdir(tmpdir_factory):
    return tmpdir_factory.mktemp('fixtures')


def copy_fixture_xls_path(fixture_dir, tmpdir, filename):
    src = os.path.join(fixture_dir, filename)
    dst = os.path.join(str(tmpdir), filename)
    shutil.copy(src, dst)
    return dst


@pytest.fixture('session')
def fixture_xls_path(fixture_dir, tmpdir):
    return copy_fixture_xls_path(fixture_dir, tmpdir, 'excelcompiler.xlsx')


@pytest.fixture('session')
def serialization_override_path(tmpdir):
    return os.path.join(str(tmpdir), 'excelcompiler_serialized.yml')


@pytest.fixture('session')
def fixture_xls_path_circular(fixture_dir, tmpdir):
    return copy_fixture_xls_path(fixture_dir, tmpdir, 'circular.xlsx')


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
    unconnected_excel.connect()
    return unconnected_excel


@pytest.fixture('session')
def basic_ws(fixture_dir, tmpdir):
    return ExcelCompiler(
        copy_fixture_xls_path(fixture_dir, tmpdir, 'basic.xlsx'))


@pytest.fixture('session')
def cond_format_ws(fixture_dir, tmpdir):
    return ExcelCompiler(
        copy_fixture_xls_path(fixture_dir, tmpdir, 'cond-format.xlsx'))


@pytest.fixture
def circular_ws(fixture_xls_path_circular):
    return ExcelCompiler(fixture_xls_path_circular, cycles=True)


@pytest.fixture
def excel_compiler(excel):
    return ExcelCompiler(excel=excel)
