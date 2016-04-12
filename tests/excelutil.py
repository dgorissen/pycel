import os
import sys

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../src')
sys.path.insert(0, path)

from pycel.excelutil import Cell
from pycel.excelcompiler import ExcelCompiler

# RUN AT THE ROOT LEVEL
excel = ExcelCompiler(os.path.join(dir, "../example/example.xlsx")).excel
cursheet = excel.get_active_sheet()

def make_cells():
    global excel, cursheet

    my_input = ['A1', 'A2:B3']
    output_cells = Cell.make_cells(excel, my_input, sheet=cursheet)
    assert len(output_cells) == 5  

make_cells()
