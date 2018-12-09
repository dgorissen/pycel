import collections
import hashlib
import itertools as it
import json
import logging
import os
import pickle

import networkx as nx
from networkx.drawing.nx_pydot import write_dot
from networkx.readwrite.gexf import write_gexf
from pycel.excelformula import ExcelFormula
from pycel.excelutil import (
    AddressCell,
    AddressRange,
    resolve_range,
)
from pycel.excelwrapper import ExcelOpxWrapper
from ruamel.yaml import YAML


__version__ = list(filter(str.isdigit, "$Revision: 2524 $"))
__date__ = list(filter(str.isdigit,
                       "$Date: 2011-09-06 17:05:00 +0100 (Tue, 06 Sep 2011) $"))
__author__ = list(filter(str.isdigit, "$Author: dg2d09 $"))


class ExcelCompiler(object):
    """Class responsible for taking an Excel spreadsheet and compiling it
    to a Spreadsheet instance that can be serialized to disk, and executed
    independently of excel.
    """

    def __init__(self, filename=None, excel=None):

        self.eval = None

        if excel:
            # if we are running as an excel addin, this gets passed to us
            self.excel = excel
            self.filename = excel.filename
            self.hash = None
        else:
            # TODO: use a proper interface so we can (eventually) support
            # loading from file (much faster)  Still need to find a good lib.
            self.excel = ExcelOpxWrapper(filename=filename)
            self.excel.connect()
            self.filename = filename

        # grab a copy of the current hash
        self._excel_file_md5_digest = self._compute_excel_file_md5_digest

        self.log = logging.getLogger('pycel')

        # directed graph for cell dependencies
        self.dep_graph = nx.DiGraph()

        # cell address to Cell mapping, cells and ranges already built
        self.cell_map = {}

        # cells, ranges and graph_edges that need to be built
        self.graph_todos = []
        self.range_todos = []

        self.extra_data = None

    def __getstate__(self):
        # code objects are not serializable
        state = dict(self.__dict__)
        for to_remove in 'eval excel log'.split():
            if to_remove in state:    # pragma: no branch
                state[to_remove] = None
        return state

    def __setstate__(self, d):
        self.__dict__.update(d)
        self.log = logging.getLogger('pycel')

    @property
    def _compute_excel_file_md5_digest(self):
        hash_md5 = hashlib.md5()
        with open(self.filename, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    @property
    def hash_matches(self):
        current_hash = self._compute_excel_file_md5_digest
        return self._excel_file_md5_digest == current_hash

    def to_json(self, filename=None):
        self.to_yaml(filename=filename, is_json=True)

    def to_yaml(self, filename=None, is_json=False):
        """Serialize to a json file"""
        extra_data = {} if self.extra_data is None else self.extra_data

        def cell_value(a_cell):
            if a_cell.formula and a_cell.formula.python_code:
                return '=' + a_cell.formula.python_code
            else:
                return a_cell.value

        extra_data.update(dict(
            excel_hash=self._excel_file_md5_digest,
            cell_map=dict(
                (addr.address, cell_value(cell))
                for addr, cell in self.cell_map.items() if not addr.is_range
            ),
        ))
        filename = filename or self.filename

        if not is_json:
            if not filename.split('.')[-1].startswith('y'):
                filename += '.yml'

            with open(filename, 'w') as f:
                ymlo = YAML()
                ymlo.width = 120
                ymlo.dump(extra_data, f)
        else:
            if not filename.endswith('.json'):
                filename += '.json'
            with open(filename, 'w') as f:
                json.dump(extra_data, f, indent=4)

        del extra_data['cell_map']

    @classmethod
    def from_json(cls, filename):
        return cls.from_yaml(filename, is_json=True)

    @classmethod
    def from_yaml(cls, filename, is_json=False):

        if not is_json:
            if not filename.split('.')[-1].startswith('y'):
                filename += '.yml'
        else:
            if not filename.endswith('.json'):  # pragma: no branch
                filename += '.json'

        with open(filename, 'r') as f:
            data = YAML().load(f)

        excel = CompiledImporter(filename)
        excel_compiler = cls(excel=excel)
        excel.compiler = excel_compiler

        for address, python_code in data['cell_map'].items():
            lineno = data['cell_map'].lc.data[address][0] + 1
            address = AddressRange(address)
            excel.value = python_code
            excel_compiler.make_cells(address)
            formula = excel_compiler.cell_map[address].formula
            if formula is not None:
                formula.lineno = lineno
                formula.filename = filename

        excel_compiler.process_gen_graph()
        del data['cell_map']

        excel_compiler._excel_file_md5_digest = data['excel_hash']
        del data['excel_hash']

        excel_compiler.extra_data = data
        return excel_compiler

    def to_file(self, filename=None):
        """"""
        filename = filename or self.filename

        # round trip through yaml to strip out junk
        yaml_name = self.filename + '.yml'
        self.to_yaml(yaml_name)
        excel_compiler = self.from_yaml(yaml_name)
        os.unlink(yaml_name)

        if not filename.split('.')[-1].startswith('p'):
            filename += '.pkl'
        with open(filename, 'wb') as f:
            pickle.dump(excel_compiler, f)

    @classmethod
    def from_file(cls, filename):
        if not filename.split('.')[-1].startswith('p'):  # pragma: no branch
            filename += '.pkl'

        with open(filename, 'rb') as f:
            excel_compiler = pickle.load(f)
        return excel_compiler

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

    def set_value(self, cell, value, is_addr=True):
        if is_addr:  # pragma: no branch
            address = AddressRange(cell)
            cell = self.cell_map[address]

        if cell.value != value:  # pragma: no branch
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = value

    def reset(self, cell):
        if cell.value is None:
            return
        self.log.info("Resetting {}".format(cell.address))
        cell.value = None
        for child_cell in self.dep_graph.successors(cell):
            if child_cell.value is not None:
                self.reset(child_cell)

    def value_tree_str(self, address, indent=0):
        cell = self.cell_map[AddressRange(address)]
        yield "{}{} = {}".format(" " * indent, address, cell.value)
        for children in sorted(self.dep_graph.predecessors(cell),
                               key=lambda a: a.address.address):
            yield from self.value_tree_str(children.address, indent + 1)

    def recalculate(self):
        for cell in self.cell_map.values():
            if isinstance(cell, CellRange) or cell.formula:
                cell.value = None

        for cell in self.cell_map.values():
            if isinstance(cell, CellRange):
                self.evaluate_range(cell)
            else:
                self.evaluate(cell)

    def trim_graph(self, input_addrs, output_addrs):
        """Remove unneeded cells from the graph"""
        output_addrs = tuple(AddressRange(addr) for addr in output_addrs)

        # build network for all needed outputs
        self.gen_graph(output_addrs)

        # walk the dependant tree and find needed nodes
        needed_cells = set()

        def walk_dependents(cell):
            for child_cell in self.dep_graph.successors(cell):
                if child_cell.address not in needed_cells:  # pragma: no branch
                    needed_cells.add(child_cell.address)
                    walk_dependents(child_cell)

        try:
            for addr in (AddressRange(a) for a in input_addrs):
                if addr in self.cell_map:
                    walk_dependents(self.cell_map[addr])
                else:
                    self.log.warning(
                        'Address {} not found in cell_map'.format(addr))
        except nx.exception.NetworkXError as exc:
            raise ValueError('{}: which usually means no outputs are dependant '
                             'on it.'.format(exc))

        for addr in output_addrs:
            needed_cells.add(addr)

        # now walk the precedent tree and prune unneeded cells
        processed_cells = set()

        def walk_precedents(cell):
            for child_address in cell.needed_addresses:
                if child_address not in processed_cells:  # pragma: no branch
                    processed_cells.add(child_address)
                    child_cell = self.cell_map[child_address]
                    if child_address in needed_cells or child_address.is_range:
                        walk_precedents(child_cell)
                    else:
                        # trim this cell, now we will need only its value
                        needed_cells.add(child_address)
                        child_cell.formula = None
                        self.log.debug('Trimming {}'.format(child_address))

        for addr in output_addrs:
            walk_precedents(self.cell_map[addr])

        cells_to_remove = tuple(addr for addr in self.cell_map
                                if addr not in needed_cells)
        for addr in cells_to_remove:
            del self.cell_map[addr]

    def validate_calcs(self, output_addrs):
        """for each address, calc the value, and verify that it matches

        This is a debugging tool which will show which cells evaluate
        differently than they do for excel.

        :param output_addrs: The cells to evaluate from
        :return: list of addresses that failed to verify
        """
        to_verify = list(AddressCell(addr) for addr in output_addrs)
        verified = set()
        failed = []
        while to_verify:
            addr = to_verify.pop()
            cell = self.cell_map[addr]
            if isinstance(cell, Cell) and cell.python_code:
                original_value = cell.value
                cell.value = None
                self.evaluate(cell)
                if original_value != cell.value:  # pragma: no branch
                    failed.append(addr)
                    print('{} mismatch  {} -> {}'.format(
                        addr, original_value, cell.value))

                    # do it again to allow easy breakpointing
                    cell.value = None
                    self.evaluate(cell)

            verified.add(addr)
            for addr in cell.needed_addresses:
                if addr not in verified:  # pragma: no branch
                    to_verify.append(addr)

        return failed

    def make_cells(self, address):
        """Given an AddressRange or AddressCell generate compiler Cells"""

        # from here don't build cells that are already in the cellmap
        assert address not in self.cell_map

        def add_node_to_graph(node):
            self.dep_graph.add_node(node)
            self.dep_graph.node[node]['sheet'] = node.sheet
            self.dep_graph.node[node]['label'] = node.address.coordinate

            # stick in queue to add edges
            self.graph_todos.append(node)

        def build_cell(addr):
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
                            cell = Cell(cell_address, value=value,
                                        formula=formula, excel=self.excel)
                            self.cell_map[cell_address] = cell
                            cells.append(cell)
            return cells

        if address.is_range:
            cell_range = CellRange(address, self)
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
            cells = cell_range.cells

            if 1 == min(cell_range.address.size):
                data = [self.evaluate(cell) for cell in cells]
            else:
                data = [[self.evaluate(cell) for cell in cells[i]] for i in
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
        if cell.python_code and cell.value is None:

            self.log.debug(
                "Evaluating: %s, %s" % (cell.address, cell.python_code))
            if self.eval is None:
                self.eval = ExcelFormula.build_eval_context(
                    self.evaluate, self.evaluate_range, self.log)
            value = self.eval(cell.formula)
            if value is None:
                value = '#EMPTY!'
            self.log.info("Cell %s evaluated to '%s' (%s)" % (
                cell.address, value, type(value).__name__))
            cell.value = value

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
                    self.gen_graph(s, recursed=True, sheet=sheet)
                self.process_gen_graph()
                return
            else:
                raise ValueError('Unknown seed: {}'.format(seed))

        # get/set the current sheet
        if not seed.has_sheet:
            seed = AddressRange(seed, sheet=self.excel.get_active_sheet_name())
        else:
            # ::TODO:: Is this needed?!?
            self.excel.set_sheet(seed.sheet)

        if seed in self.cell_map:
            # already did this cell/range
            return

        # process the seed
        self.make_cells(seed)

        if not recursed:
            # if not entered to process one cell / cellrange process other work
            self.process_gen_graph()

    def process_gen_graph(self):

        while self.graph_todos:
            # connect the dependant cells in the graph
            dependant = self.graph_todos.pop()

            self.log.debug("Handling {}".format(dependant.address))

            for precedent_address in dependant.needed_addresses:
                if precedent_address not in self.cell_map:
                    self.gen_graph(precedent_address, recursed=True)

                self.dep_graph.add_edge(
                    self.cell_map[precedent_address], dependant)

        # calc the values for ranges
        for range_todo in reversed(self.range_todos):
            self.evaluate_range(str(range_todo))

        self.log.info(
            "Graph construction done, %s nodes, "
            "%s edges, %s self.cellmap entries" % (
                len(self.dep_graph.nodes()),
                len(self.dep_graph.edges()),
                len(self.cell_map))
        )


class CellRange(object):
    # TODO: only supports rectangular ranges

    def __init__(self, address, excel):
        self.address = AddressRange(address)
        self.excel = excel
        if not self.address.sheet:
            raise ValueError("Must pass in a sheet: {}".format(address))

        self.addresses = resolve_range(self.address)
        self._cells = None
        self.size = self.address.size
        self.value = None

    def __repr__(self):
        return str(self.address)

    __str__ = __repr__

    def __iter__(self):
        if 1 == min(self.size):
            return iter(self.addresses)
        else:
            return it.chain.from_iterable(self.addresses)

    @property
    def cells(self):
        if self._cells is None:
            cell_map = self.excel.cell_map
            if 1 == min(self.size):
                self._cells = [cell_map[addr] for addr in self.addresses]
            else:
                self._cells = [[cell_map[addr] for addr in row]
                               for row in self.addresses]
        return self._cells

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
        self.address = address
        if isinstance(excel, CompiledImporter):
            excel = None

        self.excel = excel
        self.formula = formula and ExcelFormula(
            formula, cell=self, formula_is_python_code=(excel is None))
        self.value = value

        # every cell has a unique id
        self.id = Cell.next_id()

    def __repr__(self):
        return "{} -> {}".format(self.address, self.formula or self.value)

    __str__ = __repr__

    @property
    def needed_addresses(self):
        return self.formula and self.formula.needed_addresses or ()

    @property
    def sheet(self):
        return self.address.sheet

    @property
    def python_code(self):
        return self.formula and self.formula.python_code


class CompiledImporter:
    def __init__(self, filename):
        self.filename = filename.rsplit('.', maxsplit=1)[0]
        self.value = None
        self.compiler = None

    CellValue = collections.namedtuple('CellValue', 'Formula Value')

    def get_range(self, address):
        if address in self.compiler.cell_map:
            cell_map = self.compiler.cell_map
            cell = cell_map[address]
            addrs = cell.addresses

            height, width = address.size
            if height == 1:
                addrs = [addrs]
            elif width == 1:  # pragma: no branch
                addrs = [[x] for x in addrs]

            cells = [[cell_map[addr] for addr in row] for row in addrs]
            formulas = [[c.formula for c in row] for row in cells]
            values = [[c.value for c in row] for row in cells]

            return self.CellValue(formulas, values)

        elif isinstance(self.value, str) and self.value.startswith('='):
            return self.CellValue(self.value, None)
        else:
            return self.CellValue('', self.value)

    def set_sheet(self, *args):
        return
