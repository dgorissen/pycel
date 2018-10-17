from pycel.excelutil import Cell


def make_cells(excel):
    cursheet = excel.get_active_sheet()

    my_input = ['A1', 'A2:B3']
    output_cells = Cell.make_cells(excel, my_input, sheet=cursheet)
    assert len(output_cells) == 5  
