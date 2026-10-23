Pycel
=====

|build-state| |coverage| |license|

|pypi| |pypi-pyversions| |repo-size| |code-size|

Pycel is a small Python library that translates an Excel spreadsheet into
executable Python code, which can then be run independently of Excel.

The generated code is built around a dependency graph and uses caching &
lazy evaluation to ensure (relatively) fast execution. The graph can be
exported and analyzed using tools like `Gephi <http://www.gephi.org>`_.
See the ``example`` directory for a runnable illustration.

Required Python libraries:
    `python-dateutil <https://dateutil.readthedocs.io/en/stable/>`_,
    `networkx <https://networkx.github.io/>`_,
    `numpy <https://www.numpy.org/>`_,
    `openpyxl <https://openpyxl.readthedocs.io/en/stable/>`_,
    `ruamel.yaml <https://yaml.readthedocs.io/en/latest/>`_,
    and optionally: `matplotlib <https://matplotlib.org/>`_,
    `pydot <https://github.com/pydot/pydot>`_ (for graph plotting/export).

The original motivation behind pycel, along with some examples & screenshots,
is described in this `blog post
<http://elazungu.wordpress.com/2011/10/19/pycel-compiling-excel-spreadsheets-to-python-and-making-pretty-pictures>`_.

Installation
============

.. code-block:: bash

    pip install pycel

Or, for development, clone the repo and install it in editable mode:

.. code-block:: bash

    git clone https://github.com/dgorissen/pycel.git
    cd pycel
    pip install -e .

Usage
=====

**Quick start:** you can use Binder to explore pycel interactively in the
browser, no install required: |notebook|

**A minimal example:**

.. code-block:: python

    from pycel import ExcelCompiler

    # Load & compile a workbook into a dependency graph
    excel = ExcelCompiler(filename='example.xlsx')

    # Evaluate a cell
    print(excel.evaluate('Sheet1!D1'))

    # Change an input and see the result recalculate, just like in Excel
    excel.set_value('Sheet1!A1', 200)
    print(excel.evaluate('Sheet1!D1'))

    # Persist the compiled graph so it can be reloaded without Excel/openpyxl
    excel.to_file('example.xlsx')
    excel = ExcelCompiler.from_file('example.xlsx')

See ``example/example.py`` for a more complete walkthrough, including
graph plotting/export, trimming the graph down to only the cells needed
for a given input/output, and validating compiled calculations against
the values already stored in the workbook.

**The good:**

All the main mathematical functions (``SIN``, ``COS``, ``ATAN2``, ...) and
operators (``+``, ``/``, ``^``, ...) are supported, as are ranges
(``A5:D7``) and a wide range of built-in functions such as ``MIN``,
``MAX``, ``INDEX``, ``VLOOKUP``, ``SUMIFS``, ``SWITCH``, and ``LINEST``
(see ``src/pycel/excellib.py`` and ``src/pycel/lib/`` for the full list).

The codebase is relatively small and fast, and should be easy to understand
and extend. It has been tested on spreadsheets with 10+ sheets and more
than 10,000 formulae, where calculation of the equations takes on the
order of tens of milliseconds and agrees with Excel to several decimal
places.

**The bad:**

Development has historically been driven by the particular spreadsheets
contributors needed to handle, so not every Excel function is implemented.
Adding support for a missing function is usually straightforward though
- contributions are welcome.

Cell references are supported, so functions like ``OFFSET`` work, but the
compiler can fail if it encounters a cell that has not already been pulled
into the graph. Also, for obvious reasons, VBA macros are not compiled and
need to be re-implemented manually on the Python side.

**The ugly:**

The resulting graph-based code is fast enough for most purposes, but to
make it truly fast you would probably want to replace the graph with a
dependency tracker based on sparse matrices or something similar.

Security
========

Pycel compiles formulas from a workbook into Python source and executes
them. Only compile/evaluate workbooks you trust - a maliciously crafted
formula in an untrusted ``.xlsx`` file could otherwise attempt to run
arbitrary code. See the `changelog <CHANGES.rst>`_ for past advisories
(e.g. GHSA-pw67-xjhq-389w / CVE-2024-53924) and the fixes that landed
for them.

Excel Addin
===========

It's possible to run pycel as an Excel addin using
`PyXLL <http://www.pyxll.com/>`_. See ``src/pycel/addin.py`` for an example
addin module, and the ``lib/pyxll`` directory for a sample PyXLL config.
Place the compiled ``pyxll.xll``/``pyxll.py`` in the ``lib`` directory and
add the ``.xll`` file to the Excel Addins list, as explained in the PyXLL
documentation.

Development & Testing
======================

Pycel uses `tox <https://tox.wiki/>`_ to run the test suite (pytest,
flake8, and coverage) across supported Python versions:

.. code-block:: bash

    pip install tox
    tox

See `CHANGES.rst <CHANGES.rst>`_ for the full changelog, and
`.github/CONTRIBUTING.md <.github/CONTRIBUTING.md>`_ before opening an
issue or pull request.

Acknowledgements
=================

This code was originally made possible thanks to the Python port of
Eric Bachtal's `Excel formula parsing code
<http://ewbi.blogs.com/develops/popular/excelformulaparsing.html>`_
by Robin Macharg.

The code currently uses a tokenizer of similar origin from the
`openpyxl library
<https://foss.heptapod.net/openpyxl/openpyxl/-/tree/branch/default/openpyxl/formula/>`_.

Thanks also to the many contributors who have kept pycel going over the
years - see the project history on GitHub for the full list.

.. Image links

.. |build-state| image:: https://travis-ci.com/dgorissen/pycel.svg?branch=master
  :target: https://travis-ci.com/dgorissen/pycel
  :alt: Build Status

.. |coverage| image:: https://codecov.io/gh/dgorissen/pycel/branch/master/graph/badge.svg
  :target: https://codecov.io/gh/dgorissen/pycel/list/master
  :alt: Code Coverage

.. |license| image:: https://img.shields.io/badge/license-GPLv3-blue.svg
  :target: LICENSE.txt
  :alt: License: GPLv3

.. |pypi| image:: https://img.shields.io/pypi/v/pycel.svg
  :target: https://pypi.org/project/pycel/
  :alt: Latest Release

.. |pypi-pyversions| image:: https://img.shields.io/pypi/pyversions/pycel.svg
    :target: https://pypi.python.org/pypi/pycel

.. |repo-size| image:: https://img.shields.io/github/repo-size/dgorissen/pycel.svg
  :target: https://github.com/dgorissen/pycel
  :alt: Repo Size

.. |code-size| image:: https://img.shields.io/github/languages/code-size/dgorissen/pycel.svg
  :target: https://github.com/dgorissen/pycel
  :alt: Code Size

.. |notebook| image:: https://mybinder.org/badge.svg
  :target: https://mybinder.org/v2/gh/dgorissen/pycel/master?filepath=notebooks%2Fexample.ipynb
  :alt: Open Notebook
