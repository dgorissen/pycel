from pycel.excelcompiler import ExcelCompiler



def test_end_2_end(excel, example_xls_path):
    # load & compile the file to a graph, starting from D1
    excel = ExcelCompiler(excel=excel)

    # test evaluation
    assert -0.02286 == round(excel.evaluate('Sheet1!D1'), 5)

    excel.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel.evaluate('Sheet1!D1'), 5)

    # show the graph usisng matplotlib
    # sp.plot_graph()

    # export the graph, can be loaded by a viewer like gephi
    # sp.export_to_gexf(fname + ".gexf")

    # Serializing to disk...
    excel.save_to_file(example_xls_path + ".pickle")


def make_cells(excel):
    # ::TODO:: finish/fix this
    cursheet = excel.get_active_sheet()

    my_input = ['A1', 'A2:B3']
    output_cells = ExcelCompiler.make_cells(my_input, sheet=cursheet)
    assert len(output_cells) == 5

