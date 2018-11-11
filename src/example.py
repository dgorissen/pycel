"""
Simple example file showing how a spreadsheet can be translated to python
and executed
"""
import logging
import os
import sys

from pycel.excelcompiler import ExcelCompiler


def pycel_logging_to_console(enable=True):
    if enable:
        logger = logging.getLogger('pycel')
        logger.setLevel('INFO')

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        logger.addHandler(console)


if __name__ == '__main__':
    pycel_logging_to_console()

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
    excel.to_json(fname)

    print("Done")
