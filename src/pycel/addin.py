# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Simple Excel addin, requires www.pyxll.com
"""
import os
import webbrowser

import win32api
import win32com.client
from pyxll import (
    get_active_object,
    get_config,
    xl_menu
)

from pycel import AddressRange, ExcelCompiler


@xl_menu("Open log file", menu="PyXLL")
def on_open_logfile():
    # the PyXLL config is accessed as a ConfigParser.ConfigParser object
    config = get_config()
    if config.has_option("LOG", "path") and config.has_option("LOG", "file"):
        path = os.path.join(
            config.get("LOG", "path"), config.get("LOG", "file"))
        webbrowser.open("file://%s" % path)


def xl_app():
    xl_window = get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application
    return xl_app


@xl_menu("Compile selection", menu="Pycel")
def compile_selection_menu():
    curfile = xl_app().ActiveWorkbook.FullName
    newfile = curfile + ".pickle"
    selection = xl_app().Selection
    seed = selection.Address

    if not selection or seed.find(',') > 0:
        win32api.MessageBox(
            0, "You must select a cell or rectangular range of cells", "Pycel")
        return

    res = win32api.MessageBox(
        0, "Going to compile %s to %s starting from %s" % (
            curfile, newfile, seed), "Pycel", 1)
    if res == 2:
        return

    sp = do_compilation(curfile, seed)
    win32api.MessageBox(
        0, "Compilation done, graph has %s nodes and %s edges" % (
            len(sp.dep_graph.nodes()), len(sp.dep_graph.edges())), "Pycel")


def do_compilation(fname, seed, sheet=None):
    sp = ExcelCompiler(filename=fname)
    sp.evaluate(AddressRange(seed, sheet=sheet))
    sp.to_file()
    sp.export_to_gexf()
    return sp
