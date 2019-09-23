# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Simple example file showing how a spreadsheet can be translated to python
and executed
"""
import logging
import os
import sys

from pycel import ExcelCompiler


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
    fname = os.path.join(path, "example.xlsx")

    print("Loading %s..." % fname)

    # load & compile the file to a graph
    excel = ExcelCompiler(filename=fname)

    # test evaluation
    print("D1 is %s" % excel.evaluate('Sheet1!D1'))

    print("Setting A1 to 200")
    excel.set_value('Sheet1!A1', 200)

    print("D1 is now %s (the same should happen in Excel)" % excel.evaluate(
        'Sheet1!D1'))

    # show the graph using matplotlib if installed
    print("Plotting using matplotlib...")
    try:
        excel.plot_graph()
    except ImportError:
        pass

    # export the graph, can be loaded by a viewer like gephi
    print("Exporting to gexf...")
    excel.export_to_gexf(fname + ".gexf")

    # As an alternative to using evaluate to put cells in the graph and
    # as a way to trim down the size of the file to just that needed.
    excel.trim_graph(input_addrs=['Sheet1!A1'], output_addrs=['Sheet1!D1'])

    # As a sanity check, validate that the compiled code can produce
    # the current cell values.
    assert {} == excel.validate_calcs(output_addrs=['Sheet1!D1'])

    print("Serializing to disk...")
    excel.to_file(fname)

    # To reload the file later...

    print("Loading from compiled file...")
    excel = ExcelCompiler.from_file(fname)

    # test evaluation
    print("D1 is %s" % excel.evaluate('Sheet1!D1'))

    print("Setting A1 to 1")
    excel.set_value('Sheet1!A1', 1)

    print("D1 is now %s (the same should happen in Excel)" % excel.evaluate(
        'Sheet1!D1'))

    print("Done")
