Pycel
=====

|build-state| |coverage| |requirements|

|pypi| |pypi-pyversions| |repo-size| |code-size|

Pycel is a small python library that can translate an Excel spreadsheet into
executable python code which can be run independently of Excel.

The python code is based on a graph and uses caching & lazy evaluation to
ensure (relatively) fast execution.  The graph can be exported and analyzed
using tools like `Gephi <http://www.gephi.org>`_. See the contained example
for an illustration.

Required python libraries:
    `dateutil <https://dateutil.readthedocs.io/en/stable/>`_,
    `networkx <https://networkx.github.io/>`_,
    `numpy <https://www.numpy.org/>`_,
    `openpyxl <https://openpyxl.readthedocs.io/en/stable/>`_,
    `ruamel.yaml <https://yaml.readthedocs.io/en/latest/>`_, and optionally:
    `matplotlib <https://matplotlib.org/>`_,
    `pydot <https://github.com/pydot/pydot>`_

The full motivation behind pycel including some examples & screenshots is
described in this `blog post <http://www.dirkgorissen.com/2011/10/19/
pycel-compiling-excel-spreadsheets-to-python-and-making-pretty-pictures/>`_.

Usage
======

Download the library and run the example file.

**Quick start:**
You can use binder to see and explore the tool quickly and interactively in the
browser: |notebook|

**The good:**

All the main mathematical functions (sin, cos, atan2, ...) and operators
(+,/,^, ...) are supported as are ranges (A5:D7), and functions like
MIN, MAX, INDEX, LOOKUP, and LINEST.

The codebase is small, relatively fast and should be easy to understand
and extend.

I have tested it extensively on spreadsheets with 10 sheets & more than
10000 formulae.  In that case calculation of the equations takes about 50ms
and agrees with Excel up to 5 decimal places.

**The bad:**

My development is driven by the particular spreadsheets I need to handle so
I have only added support for functions that I need.  However, it is should be
straightforward to add support for others.

The code does currently not support cell references so a function like OFFSET
would take some more work to implement.  Not inherently difficult, its just
that I have had no need for references yet.  Also, for obvious reasons, any
VBA code is not compiled but needs to be re-implemented manually on the
python side.

**The Ugly:**

The resulting graph-based code is fast enough for my purposes but to make it
truly fast you would probably replace the graph with a dependency tracker
based on sparse matrices or something similar.

Excel Addin
===========

It's possible to run pycel as an excel addin using
`PyXLL <http://www.pyxll.com/>`_. Simply place pyxll.xll and pyxll.py in the
lib directory and add the xll file to the Excel Addins list as explained in
the pyxll documentation.

Acknowledgements
================

This code was originally made possible thanks to the python port of
Eric Bachtal's `Excel formula parsing code
<http://ewbi.blogs.com/develops/popular/excelformulaparsing.html>`_
by Robin Macharg.

The code currently uses a tokenizer of similar origin from the
`openpyxl library.
<https://bitbucket.org/openpyxl/openpyxl/src/default/openpyxl/formula/>`_

.. Image links

.. |build-state| image:: https://travis-ci.org/stephenrauch/pycel.svg?branch=master
  :target: https://travis-ci.org/stephenrauch/pycel
  :alt: Build Status

.. |coverage| image:: https://codecov.io/gh/stephenrauch/pycel/branch/master/graph/badge.svg
  :target: https://codecov.io/gh/stephenrauch/pycel/list/master
  :alt: Code Coverage

.. |pypi| image:: https://img.shields.io/pypi/v/pycel.svg
  :target: https://pypi.org/project/pycel/
  :alt: Latest Release

.. |pypi-pyversions| image:: https://img.shields.io/pypi/pyversions/pycel.svg
    :target: https://pypi.python.org/pypi/pycel

.. |requirements| image:: https://requires.io/github/stephenrauch/pycel/requirements.svg?branch=master
  :target: https://requires.io/github/stephenrauch/pycel/requirements/?branch=master
  :alt: Requirements Status

.. |repo-size| image:: https://img.shields.io/github/repo-size/stephenrauch/pycel.svg
  :target: https://github.com/stephenrauch/pycel
  :alt: Repo Size

.. |code-size| image:: https://img.shields.io/github/languages/code-size/stephenrauch/pycel.svg
  :target: https://github.com/stephenrauch/pycel
  :alt: Code Size

.. |notebook| image:: https://mybinder.org/badge.svg
  :target: https://mybinder.org/v2/gh/stephenrauch/pycel/master?filepath=notebooks%2Fexample.ipynb
  :alt: Open Notebook
