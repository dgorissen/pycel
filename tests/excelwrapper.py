# We will choose our wrapper with os compatibility
try:
    import win32com.client
    import pythoncom
    from pycel.excelwrapper import ExcelComWrapper as ExcelWrapperImpl
except:
    print "Can\'t import win32com -> switch from Com to Openpyxl wrapping implementation"
    from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl

import os
import sys

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../src')
sys.path.insert(0, path)

# RUN AT THE ROOT LEVEL
excel = ExcelWrapperImpl(os.path.join(dir, "../example/example.xlsx"))

def connect():
    connected = True
    try:
        excel.connect()
    except Exception as inst:
        print inst
        connected = False
    assert connected == True

def save_as():
    excel.connect()
    path_copy = os.path.join(dir, "../example/exampleCopy.xlsx")
    if os.path.exists(path_copy):
        os.remove(path_copy)
    excel.save_as(path_copy)
    assert os.path.exists(path_copy) == True

def set_and_get_active_sheet():
    excel.connect()
    excel.set_sheet("Sheet3")
    assert excel.get_active_sheet() == 'Sheet3'

def get_range():
    excel.connect()
    excel.set_sheet("Sheet2")
    excel_range = excel.get_range('Sheet2!A5:B7')
    assert sum(map(len,excel_range.cells)) == 6

def get_used_range():
    excel.connect()
    excel.set_sheet("Sheet1")
    assert sum(map(len,excel.get_used_range())) == 72

def get_value():
    excel.connect()
    excel.set_sheet("Sheet1")
    assert excel.get_value(2,2) == 9L

def get_formula():
    excel.connect()
    excel.set_sheet("Sheet1")
    assert excel.get_formula(2,2) == "=SUM(A2:A4)"
    assert excel.get_formula(3,12) == None

def has_formula():
    excel.connect()
    excel.set_sheet("Sheet1")
    assert excel.has_formula("Sheet1!C2:C5") == True
    assert excel.has_formula("Sheet1!A2:A5") == False

def get_formula_from_range():
    excel.connect()
    excel.set_sheet("Sheet1")
    formulas = excel.get_formula_from_range("Sheet1!C2:C5")
    assert len(formulas) == 4
    assert formulas[1] == "=SIN(B3*A3^2)"

def get_formula_or_value():
    excel.connect()
    excel.set_sheet("Sheet1")
    list = excel.get_formula_or_value("Sheet1!A2:C2")
    assert list == ((u'2', u'=SUM(A2:A4)', u'=SIN(B2*A2^2)'),)
    list = excel.get_formula_or_value("Sheet1!A1:A3")
    assert list == ((u'1',), (u'2',), (u'3',))

def get_row():
    excel.connect()
    assert len(excel.get_row(2)) == 4

def get_ranged_names():
    excel.connect()
    assert excel.rangednames == [[(1,'SINUS','Sheet1!$C$1:$C$18')]]

connect()
#save_as() # to disable with COM instance running 
set_and_get_active_sheet()
get_range()
get_used_range()
get_formula()
get_value()
has_formula()
get_formula_from_range()

get_formula_or_value()
get_row()
get_ranged_names()