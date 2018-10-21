"""
Simple example file showing how a spreadsheet can be translated to python
and executed
"""
from __future__ import division, print_function

import os
from pycel.excelcompiler import ExcelCompiler

if __name__ == '__main__':
    path = os.path.dirname(__file__)
    fname = os.path.join(path, "../example/example.xlsx")

    print("Loading %s..." % fname)

    # load  & compile the file to a graph, starting from D1
    excel = ExcelCompiler(filename=fname)

    # test evaluation
    print("D1 is %s" % excel.evaluate('Sheet1!D1'))

    print("Setting A1 to 200")
    excel.set_value('Sheet1!A1', 200)

    print("D1 is now %s (the same should happen in Excel)" % excel.evaluate(
        'Sheet1!D1'))

    # show the graph usisng matplotlib
    print("Plotting using matplotlib...")
    excel.plot_graph()

    # export the graph, can be loaded by a viewer like gephi
    print("Exporting to gexf...")
    excel.export_to_gexf(fname + ".gexf")

    print("Serializing to disk...")
    excel.save_to_file(fname + ".pickle")

    print("Done")
