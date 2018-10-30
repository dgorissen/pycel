import logging
import pickle
import sys

import networkx as nx
from networkx.drawing.nx_pydot import write_dot
from networkx.readwrite.gexf import write_gexf
from pycel.excelformula import ExcelFormula
from pycel.excelutil import (
    col2num,
    flatten,
    is_range,
    num2col,
    resolve_range,
    split_address,
    split_range,
)


# We will choose our wrapper with os compatibility
#       ExcelComWrapper : Must be run on Windows as it requires a
#                         COM link to an Excel instance.
#       ExcelOpxWrapper : Can be run anywhere but only with post
#                         2010 Excel formats
# ::TODO:: if keeping this move to __init__ or someplace so only needed once

if sys.platform in ('win32', 'cygwin'):
    try:
        import win32com.client
        import pythoncom
        from pycel.excelwrapper import ExcelComWrapper as ExcelWrapperImpl
    except ImportError:
        ExcelWrapperImpl = None
else:
    ExcelWrapperImpl = None

if ExcelWrapperImpl is None:
    from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl

__version__ = list(filter(str.isdigit, "$Revision: 2524 $"))
__date__ = list(filter(str.isdigit,
                       "$Date: 2011-09-06 17:05:00 +0100 (Tue, 06 Sep 2011) $"))
__author__ = list(filter(str.isdigit, "$Author: dg2d09 $"))


class CompilerError(Exception):
    """"Base class for Compiler errors"""


class ExcelCompiler(object):
    """Class responsible for taking an Excel spreadsheet and compiling it
    to a Spreadsheet instance that can be serialized to disk, and executed
    independently of excel.
    """

    def __init__(self, filename=None, excel=None):

        self.filename = filename
        self.eval = None

        if excel:
            # if we are running as an excel addin, this gets passed to us
            self.excel = excel
        else:
            # TODO: use a proper interface so we can (eventually) support
            # loading from file (much faster)  Still need to find a good lib.
            self.excel = ExcelWrapperImpl(filename=filename)
            self.excel.connect()

        self.log = logging.getLogger(
            "decode.{0}".format(self.__class__.__name__))

        # directed graph for cell dependencies
        self.dep_graph = nx.DiGraph()

        # cell address to Cell mapping
        self.cellmap = {}

    @staticmethod
    def load_from_file(fname):
        with open(fname, 'rb') as f:
            return pickle.load(f)

    def save_to_file(self, fname):
        self.excel = None
        self.log = None
        self.eval = None
        f = open(fname, 'wb')
        pickle.dump(self, f, protocol=2)
        f.close()

    def export_to_dot(self, fname):
        write_dot(self.dep_graph, fname)

    def export_to_gexf(self, fname):
        write_gexf(self.dep_graph, fname)

    def plot_graph(self, layout_type='spring_layout'):
        import matplotlib.pyplot as plt

        pos = getattr(nx, layout_type)(self.dep_graph, iterations=2000)
        nx.draw_networkx_nodes(self.dep_graph, pos)
        nx.draw_networkx_edges(self.dep_graph, pos, arrows=True)
        nx.draw_networkx_labels(self.dep_graph, pos)
        plt.show()

    def set_value(self, cell, val, is_addr=True):
        if is_addr:
            cell = self.cellmap[cell]

        if cell.value != val:
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = val

    def reset(self, cell):
        if cell.value is None:
            return
        print("resetting {}".format(cell.address()))
        cell.value = None
        for cell in self.dep_graph.successors(cell):
            self.reset(cell)

    def print_value_tree(self, addr, indent):
        cell = self.cellmap[addr]
        print("%s %s = %s" % (" " * indent, addr, cell.value))
        for c in self.dep_graph.predecessors(cell):
            self.print_value_tree(c.address(), indent + 1)

    def recalculate(self):
        for c in self.cellmap.values():
            if isinstance(c, CellRange):
                self.evaluate_range(c, is_addr=False)
            else:
                self.evaluate(c, is_addr=False)

    def resolve_cell(self, address, sheet=None):
        r = self.excel.get_range(address)
        f = r.Formula if r.Formula.startswith('=') else None
        v = r.Value

        sh, c, r = split_address(address)

        # use the sheet specified in the cell, else the passed sheet
        sheet = sh or sheet

        c = Cell(address, sheet, value=v, formula=f, excel=self.excel)
        return c

    def make_cells(self, rng, sheet=None):
        cells = []

        def convert_range(rng, sheet=None):
            cells = []

            # use the sheet specified in the range, else the passed sheet
            sh, start, end = split_range(rng)
            if sh:
                sheet = sh

            ads, numrows, numcols = resolve_range(rng)
            # ensure in the same nested format as fs/vs will be
            if numrows == 1:
                ads = [ads]
            elif numcols == 1:
                ads = [[x] for x in ads]

            # get everything in blocks, is faster
            r = self.excel.get_range(rng)
            fs = r.Formula
            vs = r.Value

            for it in (list(zip(*x)) for x in zip(ads, fs, vs)):
                row = []
                for c in it:
                    a = c[0]
                    f = c[1] if c[1] and c[1].startswith('=') else None
                    v = c[2]
                    cl = Cell(a, sheet, value=v, formula=f, excel=self.excel)
                    row.append(cl)
                cells.append(row)

            # return as vector
            if numrows == 1:
                cells = cells[0]
            elif numcols == 1:
                cells = [x[0] for x in cells]
            else:
                pass

            return cells, numrows, numcols

        if isinstance(rng, list):  # if a list of cells
            for cell in rng:
                if is_range(cell):
                    cs_in_range, nr, nc = convert_range(cell, sheet)
                    cells.append(cs_in_range)
                else:
                    c = self.resolve_cell(cell, sheet=sheet)
                    cells.append(c)

            cells = list(flatten(cells))

            # numrows and numcols are irrelevant here, so we return nr=nc=-1
            return cells, -1, -1

        else:
            if is_range(rng):
                cells, numrows, numcols = convert_range(rng, sheet)

            else:
                c = self.resolve_cell(rng, sheet=sheet)
                cells.append(c)

                numrows = 1
                numcols = 1

            return cells, numrows, numcols

    def evaluate_range(self, rng, is_addr=True):

        if is_addr:
            rng = self.cellmap[rng]
        else:
            assert isinstance(rng, CellRange)

        # it's important that [] gets treated as false here
        if rng.value:
            return rng.value

        cells, nrows, ncols = rng.celladdrs, rng.nrows, rng.ncols

        if nrows == 1 or ncols == 1:
            data = [self.evaluate(c) for c in cells]
        else:
            data = [[self.evaluate(c) for c in cells[i]] for i in
                    range(len(cells))]

        rng.value = data

        return data

    def evaluate(self, cell, is_addr=True):

        if is_addr:
            if cell not in self.cellmap:
                self.gen_graph(cell)
            cell = self.cellmap[cell]
        else:
            assert isinstance(cell, Cell)
            assert cell.address() in self.cellmap

        # no formula, fixed value
        if not cell.formula or cell.value is not None:
            # print "  returning constant or cached value for ", cell.address()
            return cell.value

        try:
            print("Evaluating: %s, %s" % (cell.address(), cell.python_code))
            if self.eval is None:
                self.eval = ExcelFormula.build_eval_context(
                    self.evaluate, self.evaluate_range)
            value = self.eval(cell.compiled_python)
            print("Cell %s evaluated to %s" % (cell.address(), value))
            if value is None:
                print("WARNING %s is None" % (cell.address()))
            cell.value = value
            
        except Exception as exc:
            if str(exc).startswith("Problem evaluating"):
                raise
            raise CompilerError("Problem evaluating: %s for %s, %s" % (
                exc, cell.address(), cell.python_code))

        return cell.value

    def gen_graph(self, seed, sheet=None):
        """Given a starting point (e.g., A6, or A3:B7) on a particular sheet,
        generate a Spreadsheet instance that captures the logic and control
        flow of the equations.
        """

        def add_node_to_graph(node):
            self.dep_graph.add_node(node)
            self.dep_graph.node[node]['sheet'] = node.sheet

            if isinstance(node, Cell):
                self.dep_graph.node[node]['label'] = node.col + str(node.row)
            else:
                # strip the sheet
                self.dep_graph.node[node]['label'] = \
                    node.address()[node.address().find('!') + 1:]

        # starting points
        cursheet = sheet or self.excel.get_active_sheet()
        self.excel.set_sheet(cursheet)

        # no need to output nr and nc here, since seed can be a list of unlinked cells
        seeds = self.make_cells(seed, sheet=cursheet)[0]

        # only keep seeds with formulas or numbers
        seeds = [s for s in flatten(seeds)
                 if s.formula or isinstance(s.value, (int, float))]

        # cells to analyze: only formulas
        todo = [s for s in seeds if s.formula if s not in self.cellmap]

        # map of all new cells
        for cell in seeds:
            if cell.address() not in self.cellmap:
                self.cellmap[cell.address()] = cell
                add_node_to_graph(cell)

        while todo:
            c1 = todo.pop()

            print("Handling ", c1.address())

            # set the current sheet so relative addresses resolve properly
            if c1.sheet != cursheet:
                cursheet = c1.sheet
                self.excel.set_sheet(cursheet)

            for addr in c1.compiled_python.needed_addresses:

                # if the dependency is a multi-cell range, create a range object
                if is_range(addr):
                    # this will make sure we always have an absolute address
                    rng = CellRange(addr, sheet=cursheet)

                    if rng.address() in self.cellmap:
                        # already dealt with this range
                        # add an edge from the range to the parent
                        self.dep_graph.add_edge(
                            self.cellmap[rng.address()],
                            self.cellmap[c1.address()]
                        )
                        continue
                    else:
                        # turn into cell objects
                        cells, nrows, ncols = self.make_cells(
                            addr, sheet=cursheet)

                        # get the values so we can set the range value
                        if nrows == 1 or ncols == 1:
                            rng.value = [c.value for c in cells]
                        else:
                            rng.value = [[c.value for c in cells[i]]
                                         for i in range(len(cells))]

                        # save the range
                        self.cellmap[rng.address()] = rng

                        # add an edge from the range to the parent
                        add_node_to_graph(rng)
                        self.dep_graph.add_edge(rng, self.cellmap[c1.address()])

                        # cells in the range have the range as their parent
                        target = rng
                else:
                    # not a range, create the cell object
                    cells = [self.resolve_cell(addr, sheet=cursheet)]
                    target = self.cellmap[c1.address()]

                # process each cell
                for c2 in flatten(cells):
                    # if we haven't treated this cell already
                    if c2.address() not in self.cellmap:
                        if c2.formula:
                            # cell with a formula, add to the todo list
                            todo.append(c2)

                        # save in the self.cellmap
                        self.cellmap[c2.address()] = c2
                        # add to the graph
                        add_node_to_graph(c2)

                    # add an edge from the cell to the parent (range or cell)
                    self.dep_graph.add_edge(self.cellmap[c2.address()], target)

        print(
            "Graph construction done, %s nodes, %s edges, %s self.cellmap entries" % (
                len(self.dep_graph.nodes()), len(self.dep_graph.edges()),
                len(self.cellmap)))


class CellRange(object):
    # TODO: only supports rectangular ranges

    def __init__(self, address, sheet=None):

        self.__address = address.replace('$', '')

        sh, start, end = split_range(address)
        if not sh and not sheet:
            raise Exception("Must pass in a sheet")

        # make sure the address is always prefixed with the range
        if sh:
            sheet = sh
        else:
            self.__address = sheet + "!" + self.__address

        addr, nrows, ncols = resolve_range(address, sheet=sheet)

        # don't allow messing with these params
        self.__celladdr = addr
        self.__nrows = nrows
        self.__ncols = ncols
        self.__sheet = sheet

        self.value = None

    def __repr__(self):
        return self.__address

    def __str__(self):
        return self.__address

    def address(self):
        return self.__address

    @property
    def celladdrs(self):
        return self.__celladdr

    @property
    def nrows(self):
        return self.__nrows

    @property
    def ncols(self):
        return self.__ncols

    @property
    def sheet(self):
        return self.__sheet


class Cell(object):
    ctr = 0

    @classmethod
    def next_id(cls):
        cls.ctr += 1
        return cls.ctr

    def __init__(self, address, sheet, value=None, formula=None, excel=None):

        sheet, c, r = split_address(address.replace('$', ''), sheet=sheet)

        # we assume a cell's location can never change
        self._sheet = sheet
        self._excel_formula = ExcelFormula(
            formula, context=Context(self, excel))

        self._col = c
        self._row = int(r)
        self._col_idx = col2num(c)

        self.value = value

        # every cell has a unique id
        self._id = Cell.next_id()

    def __repr__(self):
        return self.address()

    def __str__(self):
        if self.formula:
            return "%s%s" % (self.address(), self.formula)
        else:
            return "%s=%s" % (self.address(), self.value)

    @property
    def sheet(self):
        return self._sheet

    @property
    def row(self):
        return self._row

    @property
    def col(self):
        return self._col

    @property
    def formula(self):
        return self._excel_formula.base_formula

    @property
    def id(self):
        return self._id

    @property
    def python_code(self):
        return self._excel_formula.python_code

    @property
    def compiled_python(self):
        return self._excel_formula

    def clean_name(self):
        return self.address().replace('!', '_').replace(' ', '_')

    def address(self, absolute=True):
        if absolute:
            return "%s!%s%s" % (self._sheet, self._col, self._row)
        else:
            return "%s%s" % (self._col, self._row)

    def address_parts(self):
        return self._sheet, self._col, self._row, self._col_idx

    @staticmethod
    def inc_col_address(address, inc):
        sh, col, row = split_address(address)
        return "%s!%s%s" % (sh, num2col(col2num(col) + inc), row)

    @staticmethod
    def inc_row_address(address, inc):
        sh, col, row = split_address(address)
        return "%s!%s%s" % (sh, col, row + inc)


class Context(object):
    """A small context object that nodes can use to emit code"""

    def __init__(self, curcell, excel):
        # the current cell for which we are generating code
        self.curcell = curcell

        # a handle to athe excel instance
        self.excel = excel
