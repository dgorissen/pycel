Pycel
=====

.. image:: https://mybinder.org/badge.svg 
  :target: https://mybinder.org/v2/gh/dgorissen/pycel/master

Pycel is a small python library that can translate an Excel spreadsheet into executable python code which can be run independently of Excel.
The python code is based on a graph and uses caching & lazy evaluation to ensure (relatively) fast execution.  The graph can be exported and analyzed using
tools like `Gephi <http://www.gephi.org>`_. See the contained example for an illustration.

Required python libraries: `networkx <http://networkx.lanl.gov/>`_, `numpy <http://numpy.scipy.org/>`_, `matplotlib <http://matplotlib.sourceforge.net/>`_ (optional)

The full motivation behind pycel including some examples & screenshots is described in this `blog post <http://www.dirkgorissen.com/2011/10/19/pycel-compiling-excel-spreadsheets-to-python-and-making-pretty-pictures/>`_.

Usage
======

Download the library and run the example file, the initial compilation uses COM so an instance of Excel must be available (i.e., the compilation needs to be run on Windows).  

**Quick start:**
You can use binder to see and explore the tool quickly and interactively in the browser `Binder Example <https://mybinder.org/v2/gh/kmader/pycel/patch-1?filepath=notebooks%2Fexample.ipynb>`_

**The good:**

All the main mathematical functions (sin, cos, atan2, ...) and operators (+,/,^, ...) are supported as are ranges (A5:D7), and functions like MIN, MAX, INDEX, LOOKUP, and LINEST.
The codebase is small, relatively fast and should be easy to understand and extend.  

I have tested it extensively on spreadsheets with 10 sheets & more than 10000 formulae.  In that case calculation of the equations takes about 50ms and agrees with Excel up to 5 decimal places.

**The bad:**

My development is driven by the particular spreadsheets I need to handle so I have only added support for functions that I need.  However, it is should be straightforward to add support
for others.

The code does currently not support cell references so a function like OFFSET would take some more work to implement.  Not inherently difficult, its just that I have had no
need for references yet.  Also, for obvious reasons, any VBA code is not compiled but needs to be re-implemented manually on the python side.

**The Ugly:**

The resulting graph-based code is fast enough for my purposes but to make it truly fast you would probably replace the graph with a dependency tracker based on sparse matrices
or something similar.

Communicating with Excel over COM for the initial compilation is very slow.  I looked at various python libraries for parsing the xlsx files directly but none deal with formulas
correctly in all cases.  If that changes I will definitely add a file based compilation backend.

Using `OpenOpt <http://openopt.org/>`_ I also coded a python replacement for the Excel `solver plugin <http://www.solver.com/suppstdsolver.htm>`_.  However, since its quite closely linked with our spreadsheet structure it is not generic enough
to be released (yet).

Excel Addin
===========

Its possible to run pycel as an excel addin using `PyXLL <http://www.pyxll.com/>`_. Simply place pyxll.xll and pyxll.py in the lib directory and add the xll file to the Excel Addins list as explained in the pyxll documentation.

Acknowledgements
================

This code was made possible thanks to the python port of Eric Bachtal's `Excel formula parsing code <http://ewbi.blogs.com/develops/popular/excelformulaparsing.html>`_ by Robin Macharg.
