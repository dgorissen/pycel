import collections
import itertools as it
import logging
import pickle
import sys

import networkx as nx
from networkx.drawing.nx_pydot import write_dot
from networkx.readwrite.gexf import write_gexf
from pycel.excelformula import ExcelFormula
from pycel.excelutil import (
    AddressCell,
    AddressRange,
    flatten,
    resolve_range,
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

        # cell address to Cell mapping, cells and ranges already built
        self.cell_map = {}

        # cells, ranges and graph_edges that need to be built
        self.address_todos = []
        self.graph_todos = []
        self.range_todos = []

    @staticmethod
    def load_from_file(fname):
        with open(fname, 'rb') as f:
            return pickle.load(f)

    def save_to_file(self, fname):
        self.excel = None
        self.log = None
        self.eval = None
        with open(fname, 'wb') as f:
            pickle.dump(self, f, protocol=pickle.HIGHEST_PROTOCOL)

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
            address = AddressRange(cell)
            cell = self.cell_map[address]

        if cell.value != val:
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = val

    def reset(self, cell):
        if cell.value is None:
            return
        print("Resetting {}".format(cell.address))
        cell.value = None
        for child_cell in self.dep_graph.successors(cell):
            if child_cell.value is not None:
                self.reset(child_cell)

    def print_value_tree(self, addr, indent):
        cell = self.cell_map[addr]
        print("%s %s = %s" % (" " * indent, addr, cell.value))
        for c in self.dep_graph.predecessors(cell):
            self.print_value_tree(c.address(), indent + 1)

    def recalculate(self):
        for cell in self.cell_map.values():
            if isinstance(cell, CellRange) or cell.formula:
                cell.value = None

        for cell in self.cell_map.values():
            if isinstance(cell, CellRange):
                self.evaluate_range(cell)
            else:
                self.evaluate(cell)

    def make_cells(self, address):
        """Given an AddressRange or AddressCell generate compiler Cells"""

        # from here don't build cells that are already in the cellmap
        # ::TODO:: remove this when refactoring done
        assert address not in self.cell_map

        def add_node_to_graph(node):
            self.dep_graph.add_node(node)
            self.dep_graph.node[node]['sheet'] = node.sheet
            self.dep_graph.node[node]['label'] = node.address.coordinate

            # stick in queue to add edges
            self.graph_todos.append(node)

        def build_cell(addr):
            if address in self.cell_map:
                return

            excel_range = self.excel.get_range(addr)
            formula = None
            if excel_range.Formula.startswith('='):
                formula = excel_range.Formula

            self.cell_map[addr] = Cell(addr, value=excel_range.Value,
                                       formula=formula, excel=self.excel)
            return self.cell_map[addr]

        def build_range(rng):

            # ensure in the same nested format as fs/vs will be
            height, width = address.size
            addrs = resolve_range(rng)
            if height == 1:
                addrs = [addrs]
            elif width == 1:
                addrs = [[x] for x in addrs]

            # get everything in blocks, as it is faster
            excel_range = self.excel.get_range(rng)
            excel_cells = zip(addrs, excel_range.Formula, excel_range.Value)

            cells = []
            for row in (zip(*x) for x in excel_cells):
                for cell_address, f, value in row:
                    if cell_address not in self.cell_map:
                        formula = f if f and f.startswith('=') else None
                        if None not in (f, value):
                            cl = Cell(cell_address, value=value,
                                      formula=formula, excel=self.excel)
                            self.cell_map[cell_address] = cl
                            cells.append(cl)
            return cells

        if address.is_range:
            cell_range = CellRange(address)
            self.range_todos.append(cell_range)
            self.cell_map[address] = cell_range
            add_node_to_graph(cell_range)

            new_cells = build_range(address)
        else:
            new_cells = [build_cell(address)]

        for cell in new_cells:
            if cell.formula:
                # cells to analyze: only formulas have precedents
                add_node_to_graph(cell)

    def evaluate_range(self, cell_range):

        if isinstance(cell_range, CellRange):
            assert cell_range.address in self.cell_map
        else:
            cell_range = AddressRange(cell_range)
            assert AddressRange.has_sheet, \
                "{} missing sheetname".format(cell_range)
            if cell_range not in self.cell_map:
                self.gen_graph(cell_range)
            cell_range = self.cell_map[cell_range]

        if cell_range.value is None:
            cells = cell_range.addresses

            if 1 in cell_range.address.size:
                data = [self.evaluate(c) for c in cells]
            else:
                data = [[self.evaluate(c) for c in cells[i]] for i in
                        range(len(cells))]

            cell_range.value = data

        return cell_range.value

    def evaluate(self, cell):

        if isinstance(cell, Cell):
            assert cell.address in self.cell_map
        else:
            address = AddressRange.create(cell)
            if address not in self.cell_map:
                self.gen_graph(address)
            cell = self.cell_map[address]

        # calculate the cell value for formulas
        if cell.compiled_python and cell.value is None:

            try:
                print("Evaluating: %s, %s" % (cell.address, cell.python_code))
                if self.eval is None:
                    self.eval = ExcelFormula.build_eval_context(
                        self.evaluate, self.evaluate_range)
                value = self.eval(cell.formula)
                print("Cell %s evaluated to '%s' (%s)" % (
                    cell.address, value, type(value).__name__))
                if value is None:
                    print("WARNING %s is None" % cell.address)
                cell.value = value

            except Exception as exc:
                if str(exc).startswith("Problem evaluating"):
                    raise
                raise CompilerError("Problem evaluating: %s for %s, %s" % (
                    exc, cell.address, cell.python_code))

        return cell.value

    def gen_graph(self, seed, recursed=False, sheet=None):
        """Given a starting point (e.g., A6, or A3:B7) on a particular sheet,
        generate a Spreadsheet instance that captures the logic and control
        flow of the equations.
        """
        if not isinstance(seed, (AddressCell, AddressRange)):
            if isinstance(seed, str):
                seed = AddressRange(seed, sheet=sheet)
            elif isinstance(seed, collections.Iterable):
                for s in seed:
                    self.gen_graph(s, sheet=sheet)
                return
            else:
                raise ValueError('Unknown seed: {}'.format(seed))

        # get/set the current sheet
        if not seed.has_sheet:
            seed = AddressRange(seed, self.excel.get_active_sheet())
        else:
            self.excel.set_sheet(seed.sheet)

        if seed in self.cell_map:
            # already did this cell/range
            return

        # queue the work for the seed
        self.address_todos.append(seed)

        while self.address_todos or self.graph_todos:
            if self.address_todos:
                self.make_cells(self.address_todos.pop())

            elif recursed:
                # entered to queue up the cell / cellrange creation, so exit
                return

            else:
                # connect the dependant cells in the graph
                dependant = self.graph_todos.pop()

                # print("Handling ", dependant.address)

                for precedent_address in dependant.needed_addresses:
                    if precedent_address not in self.cell_map:
                        self.gen_graph(precedent_address, recursed=True)

                    self.dep_graph.add_edge(
                        self.cell_map[precedent_address], dependant)

        # calc the values for ranges
        for range_todo in reversed(self.range_todos):
            self.evaluate_range(range_todo)

        print(
            "Graph construction done, %s nodes, %s edges, %s self.cellmap entries" % (
                len(self.dep_graph.nodes()), len(self.dep_graph.edges()),
                len(self.cell_map)))


class CellRange(object):
    # TODO: only supports rectangular ranges

    def __init__(self, address):
        self.address = AddressRange(address)
        if not self.address.sheet:
            raise Exception("Must pass in a sheet")

        self.addresses = resolve_range(self.address)
        self.size = self.address.size
        self.value = None

    def __repr__(self):
        return str(self.address)

    __str__ = __repr__

    def __iter__(self):
        if 1 in self.size:
            return iter(self.addresses)
        else:
            return it.chain.from_iterable(self.addresses)

    @property
    def needed_addresses(self):
        return iter(self)

    @property
    def sheet(self):
        return self.address.sheet


class Cell(object):
    ctr = 0

    @classmethod
    def next_id(cls):
        cls.ctr += 1
        return cls.ctr

    def __init__(self, address, value=None, formula=None, excel=None):
        if not value and not formula:
            x = 1
        self.address = address
        self.excel = excel
        self.formula = ExcelFormula(formula, cell=self)
        self.value = value

        # every cell has a unique id
        self.id = Cell.next_id()

    def __repr__(self):
        return str(self)

    def __str__(self):
        return "{} -> {}".format(self.address, self.formula or self.value)

    @property
    def needed_addresses(self):
        return self.formula.needed_addresses

    @property
    def sheet(self):
        return self.address.sheet

    @property
    def python_code(self):
        return self.formula.python_code

    @property
    def compiled_python(self):
        return self.formula.compiled_python
