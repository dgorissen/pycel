from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import AddressRange


# ::TODO:: need some rectangular ranges for testing


def test_end_2_end(excel, example_xls_path):
    # load & compile the file to a graph, starting from D1
    excel_compiler = ExcelCompiler(excel=excel)

    # test evaluation
    assert -0.02286 == round(excel_compiler.evaluate('Sheet1!D1'), 5)

    excel_compiler.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel_compiler.evaluate('Sheet1!D1'), 5)

    # show the graph usisng matplotlib
    # sp.plot_graph()

    # export the graph, can be loaded by a viewer like gephi
    # sp.export_to_gexf(fname + ".gexf")

    # Serializing to disk...
    # excel_compiler.save_to_file(example_xls_path + ".pickle")


def test_round_trip_through_json(excel, example_xls_path):
    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.evaluate('Sheet1!D1')
    excel_compiler.extra_data = {1: 3}
    excel_compiler.to_json()

    # read the spreadsheet from json
    excel_compiler = ExcelCompiler.from_json(excel.filename)

    # test evaluation
    assert -0.02286 == round(excel_compiler.evaluate('Sheet1!D1'), 5)

    excel_compiler.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel_compiler.evaluate('Sheet1!D1'), 5)


def make_cells(excel):
    # ::TODO:: finish/fix this
    cursheet = excel.get_active_sheet()

    my_input = ['A1', 'A2:B3']
    output_cells = ExcelCompiler.make_cells(my_input, sheet=cursheet)
    assert len(output_cells) == 5


def test_trim_cells(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = [AddressRange('trim-range!D5')]
    output_addrs = [AddressRange('trim-range!B2')]

    old_value = excel_compiler.evaluate(output_addrs[0])

    excel_compiler.trim_graph(input_addrs, output_addrs)
    excel_compiler.to_json()

    new_value = ExcelCompiler.from_json(
        excel_compiler.filename).evaluate(output_addrs[0])

    assert old_value == new_value
