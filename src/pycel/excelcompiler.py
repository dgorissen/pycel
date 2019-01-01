import collections
import hashlib
import itertools as it
import json
import logging
import os
import pickle

import networkx as nx
from pycel.excelformula import ExcelFormula
from pycel.excelutil import (
    AddressCell,
    AddressRange,
    resolve_range,
)
from pycel.excelwrapper import ExcelOpxWrapper
from ruamel.yaml import YAML


class ExcelCompiler:
    """Class responsible for taking an Excel spreadsheet and compiling it
    to an instance that can be serialized to disk, and executed
    independently of excel.
    """

    save_file_extensions = ('pkl', 'pickle', 'yml', 'yaml', 'json')

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
        self._formula_cells_list = None

    def __getstate__(self):
        # code objects are not serializable
        state = dict(self.__dict__)
        for to_remove in 'eval excel log graph_todos range_todos'.split():
            if to_remove in state:    # pragma: no branch
                state[to_remove] = None
        return state

    def __setstate__(self, d):
        self.__dict__.update(d)
        self.log = logging.getLogger('pycel')

    @staticmethod
    def _compute_file_md5_digest(filename):
        hash_md5 = hashlib.md5()
        with open(filename, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    @property
    def _compute_excel_file_md5_digest(self):
        return self._compute_file_md5_digest(self.filename)

    @property
    def hash_matches(self):
        current_hash = self._compute_excel_file_md5_digest
        return self._excel_file_md5_digest == current_hash

    @classmethod
    def _filename_has_extension(cls, filename):
        return next((extension for extension in cls.save_file_extensions
                     if filename.endswith(extension)), None)

    def _to_text(self, filename=None, is_json=False):
        """Serialize to a json/yaml file"""
        extra_data = {} if self.extra_data is None else self.extra_data

        def cell_value(a_cell):
            if a_cell.formula and a_cell.formula.python_code:
                return '=' + a_cell.formula.python_code
            else:
                return a_cell.value

        extra_data.update(dict(
            excel_hash=self._excel_file_md5_digest,
            cell_map=dict(sorted(
                ((addr, cell_value(cell))
                 for addr, cell in self.cell_map.items() if ':' not in addr),
                key=lambda x: AddressCell(x[0]).sort_key
            )),
        ))
        if not filename:
            filename = self.filename + ('.json' if is_json else '.yml')

        # hash the current file to see if this function makes any changes
        existing_hash = (self._compute_file_md5_digest(filename)
                         if os.path.exists(filename) else None)

        if not is_json:
            with open(filename, 'w') as f:
                ymlo = YAML()
                ymlo.width = 120
                ymlo.dump(extra_data, f)
        else:
            with open(filename, 'w') as f:
                json.dump(extra_data, f, indent=4)

        del extra_data['cell_map']

        # hash the newfile, return True if it changed, this is only reliable
        # on pythons which have ordered dict (CPython 3.6 & python 3.7+)
        return (existing_hash is None or
                existing_hash != self._compute_file_md5_digest(filename))

    @classmethod
    def _from_text(cls, filename, is_json=False):
        """deserialize from a json/yaml file"""

        if not is_json:
            if not filename.split('.')[-1].startswith('y'):
                filename += '.yml'
        else:
            if not filename.endswith('.json'):  # pragma: no branch
                filename += '.json'

        with open(filename, 'r') as f:
            data = YAML().load(f)

        excel = _CompiledImporter(filename)
        excel_compiler = cls(excel=excel)
        excel.compiler = excel_compiler

        for address, python_code in data['cell_map'].items():
            lineno = data['cell_map'].lc.data[address][0] + 1
            address = AddressRange(address)
            excel.value = python_code
            excel_compiler._make_cells(address)
            formula = excel_compiler.cell_map[address.address].formula
            if formula is not None:
                formula.lineno = lineno
                formula.filename = filename

        excel_compiler._process_gen_graph()
        del data['cell_map']

        excel_compiler._excel_file_md5_digest = data['excel_hash']
        del data['excel_hash']

        excel_compiler.extra_data = data
        excel_compiler.excel = None
        return excel_compiler

    def to_file(self, filename=None, file_types=('pkl', 'yml')):
        """ Save the spreadsheet to a file so it can be loaded later w/o excel

        :param filename: filename to save as, defaults to xlsx_name + file_type
        :param file_types: one or more of: pkl, pickle, yml, yaml, json

        If the filename has one of the expected extensions, then this
        parameter is ignored.

        The text file formats (yaml and json) provide the benefits of:
            1. Can `diff` subsequent version of xlsx to monitor changes.
            2. Can "debug" the generated code.
                Since the compiled code is marked with the line number in
                the text file and will be shown by debuggers and stack traces.
            3. The file size on disk is somewhat smaller than pickle files

        The pickle file format provides the benefits of:
            1. Much faster to load (5x to 10x)
            2. ...  (no #2, speed is the thing)
        """

        filename = filename or self.filename
        extension = self._filename_has_extension(filename)
        if extension:
            file_types = (extension, )
        elif isinstance(file_types, str):
            file_types = (file_types, )

        unknown_types = tuple(ft for ft in file_types
                              if ft not in self.save_file_extensions)
        if unknown_types:
            raise ValueError('Unknown file types: {}'.format(
                ' '.join(unknown_types)))

        pickle_extension = next((ft for ft in file_types
                                 if ft.startswith('p')), None)
        non_pickle_extension = next((ft for ft in file_types
                                     if not ft.startswith('p')), None)
        extra_extensions = tuple(ft for ft in file_types if ft not in (
            pickle_extension, non_pickle_extension))

        if extra_extensions:
            raise ValueError(
                'Only allowed one pickle extension and one text extension. '
                'Extras: {}'.format(extra_extensions))

        is_json = non_pickle_extension and non_pickle_extension[0] == 'j'

        # round trip through yaml/json to strip out junk
        text_name = filename
        if not text_name.endswith(non_pickle_extension or '.yml'):
            text_name += '.' + (non_pickle_extension or 'yml')
        text_changed = self._to_text(text_name, is_json=is_json)

        # save pickle file if requested and has changed
        if pickle_extension:
            if not filename.endswith(pickle_extension):
                filename += '.' + pickle_extension

            if text_changed or not os.path.exists(filename):
                excel_compiler = self._from_text(text_name, is_json=is_json)
                if non_pickle_extension not in file_types:
                    os.unlink(text_name)

                with open(filename, 'wb') as f:
                    pickle.dump(excel_compiler, f)

    @classmethod
    def from_file(cls, filename):
        """ Load the spreadsheet saved by `to_file`

        :param filename: filename to load from, can be xlsx_name
        """

        extension = cls._filename_has_extension(filename) or next(
            (ext for ext in cls.save_file_extensions
             if os.path.exists(filename + '.' + ext)), None)

        if not extension:
            raise ValueError("Unrecognized file type or compiled file not found"
                             ": '{}'".format(filename))

        if not filename.endswith(extension):
            filename += '.' + extension

        if extension[0] == 'p':
            with open(filename, 'rb') as f:
                excel_compiler = pickle.load(f)
        else:
            excel_compiler = cls._from_text(
                filename, is_json=extension == 'json')

        return excel_compiler

    def export_to_dot(self, filename=None):
        try:
            # test pydot is importable  (optionally installed)
            import pydot  # noqa: F401
        except ImportError:
            raise ImportError("Package 'pydot' is not installed")

        from networkx.drawing.nx_pydot import write_dot
        filename = filename or (self.filename + '.dot')
        write_dot(self.dep_graph, filename)

    def export_to_gexf(self, filename=None):
        from networkx.readwrite.gexf import write_gexf
        filename = filename or (self.filename + '.gexf')
        write_gexf(self.dep_graph, filename)

    def plot_graph(self, layout_type='spring_layout'):
        try:
            # test matplotlib is importable  (optionally installed)
            import matplotlib.pyplot as plt
        except ImportError:
            raise ImportError("Package 'matplotlib' is not installed")

        pos = getattr(nx, layout_type)(self.dep_graph, iterations=2000)
        nx.draw_networkx_nodes(self.dep_graph, pos)
        nx.draw_networkx_edges(self.dep_graph, pos, arrows=True)
        nx.draw_networkx_labels(self.dep_graph, pos)
        plt.show()

    def set_value(self, address, value):
        """ Set the value of one or more cells or ranges

        :param address: `str`, `AddressRange`, `AddressCell` or a tuple, list
            or an iterable of these three
        :param value: value to set.  This can be a value or a tuple/list
            which matches the shapes needed for the given address/addresses
        """

        if (not isinstance(address, (AddressRange, AddressCell)) and
                isinstance(address, (tuple, list))):
            assert isinstance(value, (tuple, list))
            assert len(address) == len(value)
            for addr, val in zip(address, value):
                self.set_value(addr, val)
            return

        elif address not in self.cell_map:
            address = AddressRange.create(address).address
            assert address in self.cell_map

        cell_or_range = self.cell_map[address]

        if cell_or_range.value != value:  # pragma: no branch
            # need to be able to 'set' an empty cell
            if cell_or_range.value is None:
                cell_or_range.value = value

            # reset the node + its dependencies
            self._reset(cell_or_range)

            # set the value
            cell_or_range.value = value

    def _reset(self, cell):
        if cell.value is None:
            return
        self.log.info("Resetting {}".format(cell.address))
        cell.value = None

        if cell in self.dep_graph:
            for child_cell in self.dep_graph.successors(cell):
                if child_cell.value is not None:
                    self._reset(child_cell)

    def value_tree_str(self, address, indent=0):
        """Generator which returns a formatted dependency graph"""
        cell = self.cell_map[address]
        yield "{}{} = {}".format(" " * indent, address, cell.value)
        for children in sorted(self.dep_graph.predecessors(cell),
                               key=lambda a: a.address.address):
            yield from self.value_tree_str(children.address.address, indent + 1)

    def recalculate(self):
        """Recalculate all of the known cells"""
        for cell in self.cell_map.values():
            if isinstance(cell, _CellRange) or cell.formula:
                cell.value = None

        for cell in self.cell_map.values():
            if isinstance(cell, _CellRange):
                self._evaluate_range(cell.address.address)
            else:
                self._evaluate(cell.address.address)

    def trim_graph(self, input_addrs, output_addrs):
        """Remove unneeded cells from the graph"""
        input_addrs = tuple(AddressRange(addr).address for addr in input_addrs)
        output_addrs = tuple(AddressRange(addr) for addr in output_addrs)

        # 1) build graph for all needed outputs
        self._gen_graph(output_addrs)

        # 2) walk the dependant tree (from the inputs) and find needed cells
        needed_cells = set()

        def walk_dependents(cell):
            """passed in a _Cell or _CellRange"""
            for child_cell in self.dep_graph.successors(cell):
                child_addr = child_cell.address.address
                if child_addr not in needed_cells:
                    needed_cells.add(child_addr)
                    walk_dependents(child_cell)

        try:
            for addr in input_addrs:
                if addr in self.cell_map:
                    walk_dependents(self.cell_map[addr])
                else:
                    self.log.warning(
                        'Address {} not found in cell_map'.format(addr))
        except nx.exception.NetworkXError as exc:
            if AddressRange(addr) not in output_addrs:
                raise ValueError('{}: which usually means no outputs '
                                 'are dependant on it.'.format(exc))

        # even unconnected output addresses are needed
        for addr in output_addrs:
            needed_cells.add(addr.address)

        # 3) walk the precedent tree (from the output) and trim unneeded cells
        processed_cells = set()

        def walk_precedents(cell):
            for child_address in (a.address for a in cell.needed_addresses):
                if child_address not in processed_cells:  # pragma: no branch
                    processed_cells.add(child_address)
                    child_cell = self.cell_map[child_address]
                    if child_address in needed_cells or ':' in child_address:
                        walk_precedents(child_cell)
                    else:
                        # trim this cell, now we will need only its value
                        needed_cells.add(child_address)
                        child_cell.formula = None
                        self.log.debug('Trimming {}'.format(child_address))

        for addr in output_addrs:
            walk_precedents(self.cell_map[addr.address])

        # 4) check for any buried (not leaf node) inputs
        for addr in input_addrs:
            cell = self.cell_map.get(addr)
            if cell and getattr(cell, 'formula', None):
                self.log.info("{} is not a leaf node".format(addr))

        # 5) remove unneeded cells
        cells_to_remove = tuple(addr for addr in self.cell_map
                                if addr not in needed_cells)
        for addr in cells_to_remove:
            del self.cell_map[addr]

    def validate_calcs(self, output_addrs=None):
        """For each address, calc the value, and verify that it matches

        This is a debugging tool which will show which cells evaluate
        differently than they do for excel.

        :param output_addrs: The cells to evaluate from (defaults to all)
        :return: dict of addresses with good/bad values that failed to verify
        """
        def close_enough(val1, val2):
            import pytest
            if isinstance(val1, (int, float)):
                return val2 == pytest.approx(val1)
            else:
                return val1 == val2

        Mismatch = collections.namedtuple('Mismatch', 'original calced formula')

        if output_addrs is None:
            to_verify = self._formula_cells
        else:
            to_verify = list(AddressCell(addr) for addr in output_addrs)
        verified = set()
        failed = {}
        while to_verify:
            addr = to_verify.pop()
            try:
                self._gen_graph(addr)
                cell = self.cell_map[addr.address]
                if isinstance(cell, _Cell) and cell.python_code:
                    original_value = cell.value
                    cell.value = None
                    self._evaluate(cell.address.address)

                    # pragma: no branch
                    if not close_enough(original_value, cell.value):
                        failed.setdefault('mismatch', {})[str(addr)] = Mismatch(
                            original_value, cell.value,
                            cell.formula.base_formula)
                        print('{} mismatch  {} -> {}  {}'.format(
                            addr, original_value, cell.value,
                            cell.formula.base_formula))

                        # do it again to allow easy breakpointing
                        cell.value = None
                        self._evaluate(cell.address.address)

                verified.add(addr)
                for addr in cell.needed_addresses:
                    if addr not in verified:  # pragma: no branch
                        to_verify.append(addr)
            except Exception as exc:   # pragma: no cover
                cell = self.cell_map.get(addr.address, None)
                formula = cell and cell.formula.base_formula
                exc_str = str(exc)
                exc_str_split = exc_str.split('\n')
                if len(exc_str_split) == 1:
                    exc_str_key = '{}: {}'.format(type(exc).__name__, exc_str)
                else:
                    exc_str_key = exc_str_split[-2]

                if ('NameError: name ' in exc_str
                        or exc_str_key.startswith('NotImplementedError: ')):
                    failed.setdefault('not-implemented', {}).setdefault(
                        exc_str_key, []).append((str(addr), formula, exc_str))
                else:
                    failed.setdefault('exceptions', {}).setdefault(
                        exc_str_key, []).append((str(addr), formula, exc_str))

        return failed

    @property
    def _formula_cells(self):
        """Iterate all cells and find cells with formulas"""

        if self._formula_cells_list is None:
            self._formula_cells_list = [
                AddressCell.create(cell.coordinate, ws.title)
                for ws in self.excel.workbook
                for row in ws.iter_rows()
                for cell in row
                if isinstance(cell.value, str) and cell.value.startswith('=')
            ]
        return self._formula_cells_list

    def _make_cells(self, address):
        """Given an AddressRange or AddressCell generate compiler Cells"""

        # from here don't build cells that are already in the cell_map
        assert address.address not in self.cell_map

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

            self.cell_map[addr.address] = _Cell(
                addr, value=excel_range.Value,
                formula=formula, excel=self.excel)
            return self.cell_map[addr.address]

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
                    if cell_address.address not in self.cell_map:
                        formula = f if f and f.startswith('=') else None
                        if None not in (f, value):
                            cell = _Cell(cell_address, value=value,
                                         formula=formula, excel=self.excel)
                            self.cell_map[cell_address.address] = cell
                            cells.append(cell)
            return cells

        if address.is_range:
            cell_range = _CellRange(address, self)
            self.range_todos.append(address.address)
            self.cell_map[address.address] = cell_range
            add_node_to_graph(cell_range)

            new_cells = build_range(address)
        else:
            new_cells = [build_cell(address)]

        for cell in new_cells:
            if cell.formula:
                # cells to analyze: only formulas have precedents
                add_node_to_graph(cell)

    def _evaluate_range(self, address):

        cell_range = self.cell_map.get(address)
        if cell_range is None:
            # we don't save the _CellRange values in the text format files
            assert '!' in address, "{} missing sheetname".format(address)
            self._gen_graph(address)
            cell_range = self.cell_map[address]

        if cell_range.value is None:
            self.log.debug("Evaluating: {}".format(cell_range.address))
            cells = cell_range.cells

            if 1 == min(cell_range.address.size):
                data = [self._evaluate(cell.address.address) for cell in cells]
            else:
                data = [
                    [self._evaluate(cell.address.address) for cell in cells[i]]
                    for i in range(len(cells))
                ]

            cell_range.value = data

        return cell_range.value

    def _evaluate(self, address):
        cell = self.cell_map[address]

        # calculate the cell value for formulas and ranges
        if cell.value is None:
            if isinstance(cell, _CellRange):
                self._evaluate_range(cell.address.address)

            elif cell.python_code:
                self.log.debug(
                    "Evaluating: %s, %s".format(cell.address, cell.python_code))
                if self.eval is None:
                    self.eval = ExcelFormula.build_eval_context(
                        self._evaluate, self._evaluate_range, self.log)
                value = self.eval(cell.formula)
                self.log.info("Cell %s evaluated to '%s' (%s)" % (
                    cell.address, value, type(value).__name__))
                cell.value = value

        return cell.value

    def evaluate(self, address):
        """ evaluate a cell or cells in the spreadsheet

        :param address: str, AddressRange, AddressCell or a tuple or list
            or iterable of these three
        :return: evaluted value/values
        """

        try:
            not_in_cell_map = address not in self.cell_map
        except TypeError:
            not_in_cell_map = True

        if not_in_cell_map:
            if (not isinstance(address, (str, AddressRange, AddressCell)) and
                    isinstance(address, collections.Iterable)):

                if not isinstance(address, (tuple, list)):
                    address = tuple(address)

                # process a tuple or list of addresses
                return type(address)(self.evaluate(c) for c in address)

            address = AddressRange.create(address).address
            if address not in self.cell_map:
                self._gen_graph(address)

        return self._evaluate(address)

    def _gen_graph(self, seed, recursed=False):
        """Given a starting point (e.g., A6, or A3:B7) on a particular sheet,
        generate a Spreadsheet instance that captures the logic and control
        flow of the equations.
        """
        if not isinstance(seed, (AddressCell, AddressRange)):
            if isinstance(seed, str):
                seed = AddressRange(seed)
            elif isinstance(seed, collections.Iterable):
                for s in seed:
                    self._gen_graph(s, recursed=True)
                self._process_gen_graph()
                return
            else:
                raise ValueError('Unknown seed: {}'.format(seed))

        # get/set the current sheet
        if not seed.has_sheet:
            seed = AddressRange(seed, sheet=self.excel.get_active_sheet_name())

        if seed.address in self.cell_map:
            # already did this cell/range
            return

        # process the seed
        self._make_cells(seed)

        if not recursed:
            # if not entered to process one cell / cellrange process other work
            self._process_gen_graph()

    def _process_gen_graph(self):

        while self.graph_todos:
            # connect the dependant cells in the graph
            dependant = self.graph_todos.pop()

            self.log.debug("Handling {}".format(dependant.address))

            for precedent_address in dependant.needed_addresses:
                if precedent_address.address not in self.cell_map:
                    self._gen_graph(precedent_address, recursed=True)

                self.dep_graph.add_edge(
                    self.cell_map[precedent_address.address], dependant)

        # calc the values for ranges
        for range_todo in reversed(self.range_todos):
            self._evaluate_range(range_todo)
        self.range_todos = []

        self.log.info(
            "Graph construction done, %s nodes, "
            "%s edges, %s self.cell_map entries" % (
                len(self.dep_graph.nodes()),
                len(self.dep_graph.edges()),
                len(self.cell_map))
        )


class _CellRange:
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

    def __getstate__(self):
        state = dict(self.__dict__)
        state['excel'] = None
        return state

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
                self._cells = [cell_map[addr.address]
                               for addr in self.addresses]
            else:
                self._cells = [[cell_map[addr.address] for addr in row]
                               for row in self.addresses]
        return self._cells

    @property
    def needed_addresses(self):
        return iter(self)

    @property
    def sheet(self):
        return self.address.sheet


class _Cell:
    ctr = 0

    @classmethod
    def next_id(cls):
        cls.ctr += 1
        return cls.ctr

    def __init__(self, address, value=None, formula=None, excel=None):
        self.address = address
        if isinstance(excel, _CompiledImporter):
            excel = None

        self.excel = excel
        self.formula = formula and ExcelFormula(
            formula, cell=self, formula_is_python_code=(excel is None))
        self.value = value

        # every cell has a unique id
        self.id = _Cell.next_id()

    def __getstate__(self):
        state = dict(self.__dict__)
        state['excel'] = None
        return state

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


class _CompiledImporter:
    def __init__(self, filename):
        self.filename = filename.rsplit('.', maxsplit=1)[0]
        self.value = None
        self.compiler = None

    CellValue = collections.namedtuple('CellValue', 'Formula Value')

    def get_range(self, address):
        if address.address in self.compiler.cell_map:
            cell_map = self.compiler.cell_map
            cell = cell_map[address.address]
            addrs = cell.addresses

            height, width = address.size
            if height == 1:
                addrs = [addrs]
            elif width == 1:  # pragma: no branch
                addrs = [[x] for x in addrs]

            cells = [[cell_map[addr.address] for addr in row] for row in addrs]
            formulas = [[c.formula for c in row] for row in cells]
            values = [[c.value for c in row] for row in cells]

            return self.CellValue(formulas, values)

        elif isinstance(self.value, str) and self.value.startswith('='):
            return self.CellValue(self.value, None)
        else:
            return self.CellValue('', self.value)
