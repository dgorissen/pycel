from pycel.excelutil import *


def test_is_range():

    assert is_range('a1:b2')
    assert not is_range('a1')


def test_split_range():

    assert (None, 'B1', None) == split_range('B1')
    assert ('sheet', 'B1', None) == split_range('sheet!B1')
    assert (None, 'B1', 'C2') == split_range('B1:C2')
    assert ('sheet', 'B1', 'C2') == split_range('sheet!B1:C2')
