# We will choose our wrapper with os compatibility
#       ExcelComWrapper : Must be run on Windows as it requires a COM link to an Excel instance.
#       ExcelOpxWrapper : Can be run anywhere but only with post 2010 Excel formats
import logging
from math import *
import pickle
import sys

import networkx as nx
import openpyxl.formula.tokenizer as tokenizer
from networkx.classes.digraph import DiGraph
from networkx.drawing.nx_pydot import write_dot
from networkx.readwrite.gexf import write_gexf

# ::TODO:: import *, or import in the eval function, or map all excel
# functions, or build a plugin system on failed look ups?
from pycel.excellib import (
    FUNCTION_MAP,
    linest,
    xsum,
)

from pycel.excelutil import (
    Cell,
    CellRange,
    flatten,
    get_linest_degree,
    is_range,
    resolve_range,
    split_address,
    split_range,
    uniqueify,
)

# ::TODO:: if keeping this move to __init__ or someplace so only needed once
ExcelWrapperImpl = None
if sys.platform in ('win32', 'cygwin'):
    try:
        import win32com.client
        import pythoncom
        from pycel.excelwrapper import ExcelComWrapper as ExcelWrapperImpl
    except ImportError:
        ExcelWrapperImpl = None

if ExcelWrapperImpl is None:
    from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl

__version__ = list(filter(str.isdigit, "$Revision: 2524 $"))
__date__ = list(filter(str.isdigit,
                       "$Date: 2011-09-06 17:05:00 +0100 (Tue, 06 Sep 2011) $"))
__author__ = list(filter(str.isdigit, "$Author: dg2d09 $"))


class CompilerError(Exception):
    """"Base class for Compiler errors"""


class Tokenizer(tokenizer.Tokenizer):
    """Amend openpyxwl tokenizer"""

    def __init__(self, formula):
        super(Tokenizer, self).__init__(formula)
        self.items = self._items()

    def _items(self):
        """ convert to use our Token"""
        t = [None] + [Token.from_token(t) for t in self.items] + [None]

        # convert or remove unneeded whitespace
        tokens = []
        for prev_token, token, next_token in zip(t, t[1:], t[2:]):
            if token.type != Token.WSPACE or not prev_token or not next_token:
                tokens.append(token)

            elif (prev_token.matches(type_=Token.FUNC, subtype=Token.CLOSE) or
                  prev_token.matches(type_=Token.PAREN, subtype=Token.CLOSE) or
                  prev_token.type == Token.OPERAND
            ) and (
                next_token.matches(type_=Token.FUNC, subtype=Token.OPEN) or
                next_token.matches(type_=Token.PAREN, subtype=Token.OPEN) or
                next_token.type == Token.OPERAND
            ):
                # this whitespace is an intersect operator
                tokens.append(Token(token.value, Token.OP_IN, Token.INTERSECT))
        return tokens


class Token(tokenizer.Token):
    """Amend openpyxwl token"""

    INTERSECT = "INTERSECT"
    ARRAYROW = "ARRAYROW"

    @classmethod
    def from_token(cls, token, value=None, type_=None, subtype=None):
        return cls(
            token.value if value is None else value,
            token.type if type_ is None else type_,
            token.subtype if subtype is None else subtype
        )

    @property
    def is_operator(self):
        return self.type in (Token.OP_PRE, Token.OP_IN, Token.OP_POST)

    @property
    def is_funcopen(self):
        return self.subtype == Token.OPEN and self.type in (
            Token.FUNC, Token.ARRAY, Token.ARRAYROW)

    def matches(self, type_=None, subtype=None, value=None):
        return ((type_ is None or self.type == type_) and
                (subtype is None or self.subtype == subtype) and
                (value is None or self.value == value))


class ASTNode(object):
    """A generic node in the AST used to compile a cell's formula"""

    class Context(object):
        """A small context object that nodes in the AST can use to emit code"""

        def __init__(self, curcell, excel):
            # the current cell for which we are generating code
            self.curcell = curcell
            # a handle to an excel instance
            self.excel = excel

    def __init__(self, token):
        super(ASTNode, self).__init__()
        self.token = token
        self._ast = None
        self._parent = None
        self._children = None
        self._descendents = None

    @classmethod
    def create(cls, token):
        """Simple factory function"""
        if token.type == token.OPERAND:
            if token.subtype == token.RANGE:
                return RangeNode(token)
            else:
                return OperandNode(token)
        elif token.is_funcopen:
            return FunctionNode(token)
        elif token.is_operator:
            return OperatorNode(token)
        else:
            return ASTNode(token)

    def __str__(self):
        return str(self.token.value.strip('('))

    def __repr__(self):
        return '{}<{}>'.format(type(self).__name__,
                               str(self.token.value.strip('(')))

    @property
    def ast(self):
        return self._ast

    @ast.setter
    def ast(self, value):
        self._ast = value

    @property
    def value(self):
        return self.token.value

    @property
    def type(self):
        return self.token.type

    @property
    def subtype(self):
        return self.token.subtype

    @property
    def children(self):
        if self._children is None:
            args = self.ast.predecessors(self)
            self._children = sorted(args, key=lambda x: self.ast.node[x]['pos'])
        # args.reverse()
        return self._children

    @property
    def descendents(self):
        if self._descendents is None:
            self._descendents = list(self.ast.nodes(self))
        return self._descendents

    @property
    def parent(self):
        if self._parent is None:
            self._parent = next(self.ast.successors(self), None)
        return self._parent

    def emit(self, context=None):
        """Emit code"""
        return self.value


class OperatorNode(ASTNode):
    # convert the operator to python equivalents
    op_map = {
        "^": "**",
        "=": "==",
        "&": "+",
        " ": "+"  # range intersection
    }

    def emit(self, context=None):
        xop = self.value

        # Get the arguments
        args = self.children

        op = self.op_map.get(xop, xop)

        if self.type == Token.OP_PRE:
            return "-" + args[0].emit(context)

        parent = self.parent
        # dont render the ^{1,2,..} part in a linest formula
        # TODO: bit of a hack
        if op == "**":
            if parent and parent.value.lower() == "linest(":
                return args[0].emit(context)

        # TODO silly hack to work around the fact that None < 0 is True
        #  (happens on blank cells)
        if op.startswith('<'):
            aa = args[0].emit(context)
            ss = "({} if {} is not None else 0) {} {}".format(
                aa, aa, op, args[1].emit(context))

        elif op.startswith('>'):
            aa = args[1].emit(context)
            ss = "{} {} ({} if {} is not None else 0)".format(
                args[0].emit(context), op, aa, aa, )
        else:
            if op != ',':
                op = ' ' + op
            ss = '{}{} {}'.format(
                args[0].emit(context), op, args[1].emit(context))

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "(" + ss + ")"

        return ss


class OperandNode(ASTNode):

    def emit(self, context=None):
        if self.subtype == self.token.LOGICAL:
            return str(self.value.lower() == "true")

        elif self.subtype in ("TEXT", "ERROR") and len(self.value) > 2:
            # if the string contains quotes, escape them
            return '"{}"'.format(self.value.replace('""', '\\"').strip('"'))

        else:
            return self.value


class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""

    def get_cells(self):
        return resolve_range(self.value)[0]

    def emit(self, context=None):
        # resolve the range into cells
        rng = self.value.replace('$', '')
        sheet = context.curcell.sheet + "!" if context else ""
        if is_range(rng):
            sh, start, end = split_range(rng)
            if sh:
                return 'eval_range("' + rng + '")'
            else:
                return 'eval_range("' + sheet + rng + '")'
        else:
            sh, col, row = split_address(rng)
            if sh:
                return 'eval_cell("' + rng + '")'
            else:
                return 'eval_cell("' + sheet + rng + '")'


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, *args):
        super(FunctionNode, self).__init__(*args)
        self.num_args = 0

        # map excel functions onto their python equivalents
        self.func_map = FUNCTION_MAP

    def comma_join_emit(self, context=None, fmt_str=None):
        if fmt_str is None:
            return ", ".join(n.emit(context) for n in self.children)
        else:
            return ", ".join(
                fmt_str.format(n.emit(context)) for n in self.children)

    def emit(self, context=None):
        func = self.value.lower().strip('(')

        # if a special handler is needed
        handler = getattr(self, 'func_{}'.format(func), None)
        if handler is not None:
            return handler(context)

        else:
            # map to the correct name
            return "{}({})".format(
                self.func_map.get(func, func), self.comma_join_emit(context))

    def func_atan2(self, context=None):
        # swap arguments
        a1, a2 = (a.emit(context) for a in self.children)
        return "atan2({}, {})".format(a2, a1)

    @staticmethod
    def func_pi(context=None):
        # constant, no parens
        # ::TODO:: need test case
        return "pi"

    def func_if(self, context=None):
        # inline the if
        args = [c.emit(context) for c in self.children] + [0]
        if len(args) in (3, 4):
            return "({} if {} else {})".format(args[1], args[0], args[2])

        raise CompilerError(
            "IF with %s arguments not supported".format(len(args) - 1))

    def func_array(self, context=None):
        if len(self.children) == 1:
            return '[{}]'.format(self.children[0].emit(context))
        else:
            # multiple rows
            return '[{}]'.format(self.comma_join_emit(context, '[{}]'))

    def func_arrayrow(self, context=None):
        # simply create a list
        return self.comma_join_emit(context)

    def func_linest(self, context=None):
        func = self.value.lower().strip('(')
        code = '{}({}'.format(func, self.comma_join_emit(context))

        if not context:
            degree, coef = -1, -1
        else:
            # linests are often used as part of an array formula spanning
            # multiple cells, one cell for each coefficient.  We have to
            # figure out where we currently are in that range.
            degree, coef = get_linest_degree(context.excel, context.curcell)

        # if we are the only linest (degree is one) and linest is nested
        # return vector, else return the coef.
        if func == "linest":
            code += ", degree=%s)" % degree
        else:
            code += ")"

        if not (degree == 1 and self.parent):
            code += "[%s]" % (coef - 1)

        return code

    def func_and(self, context=None):
        return "all([{}])".format(self.comma_join_emit(context))

    def func_or(self, context=None):
        return "any([{}])".format(self.comma_join_emit(context))


class Operator:
    """Small wrapper class to manage operators during shunting yard"""

    def __init__(self, value, precedence, associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity

    def __lt__(self, other):
        return (self.precedence < other.precedence or
                self.associativity == "left" and
                self.precedence == other.precedence
                )


class ExcelCompiler(object):
    """Class responsible for taking an Excel spreadsheet and compiling it
    to a Spreadsheet instance that can be serialized to disk, and executed
    independently of excel.
    """

    operators = {
        # http://office.microsoft.com/en-us/excel-help/
        #   calculation-operators-and-precedence-HP010078886.aspx
        ':': Operator(':', 8, 'left'),
        ' ': Operator(' ', 8, 'left'),  # range intersection
        ',': Operator(',', 8, 'left'),
        'u-': Operator('u-', 7, 'left'),  # unary negation
        '%': Operator('%', 6, 'left'),
        '^': Operator('^', 5, 'left'),
        '*': Operator('*', 4, 'left'),
        '/': Operator('/', 4, 'left'),
        '+': Operator('+', 3, 'left'),
        '-': Operator('-', 3, 'left'),
        '&': Operator('&', 2, 'left'),
        '=': Operator('=', 1, 'left'),
        '<': Operator('<', 1, 'left'),
        '>': Operator('>', 1, 'left'),
        '<=': Operator('<=', 1, 'left'),
        '>=': Operator('>=', 1, 'left'),
        '<>': Operator('<>', 1, 'left'),
    }

    def __init__(self, filename=None, excel=None):

        self.filename = filename

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
        f = open(fname, 'wb')
        pickle.dump(self, f, protocol=2)
        f.close()

    def export_to_dot(self, fname):
        write_dot(self.dep_graph, fname)

    def export_to_gexf(self, fname):
        write_gexf(self.dep_graph, fname)

    def plot_graph(self):
        import matplotlib.pyplot as plt

        pos = nx.spring_layout(self.dep_graph, iterations=2000)
        # pos=nx.spectral_layout(G)
        # pos = nx.random_layout(G)
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

        # no formula, fixed value
        if not cell.formula or cell.value is not None:
            # print "  returning constant or cached value for ", cell.address()
            return cell.value

        # recalculate formula
        # the compiled expression calls this function
        def eval_cell(address):
            return self.evaluate(address)

        def eval_range(rng):
            return self.evaluate_range(rng)

        try:
            print("Evaluating: %s, %s" % (cell.address(), cell.python_expression))
            value = eval(cell.compiled_expression)
            print("Cell %s evaluated to %s" % (cell.address(), value))
            if value is None:
                print("WARNING %s is None" % (cell.address()))
            cell.value = value
        except Exception as exc:
            if str(exc).startswith("Problem evaluating"):
                raise
            raise CompilerError("Problem evaluating: %s for %s, %s" % (
                exc, cell.address(), cell.python_expression))

        return cell.value

    @classmethod
    def parse_to_rpn(cls, expression):
        """
        Parse an excel formula expression into reverse polish notation

        Core algorithm taken from wikipedia with varargs extensions from
        http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-
            algorithm-to-allow-variable-numbers-of-arguments-to-functions/
        """

        lexer = Tokenizer(expression)

        # amend token stream to ease code production
        tokens = []
        for token, next_token in zip(lexer.items, lexer.items[1:] + [None]):

            if token.matches(Token.FUNC, Token.OPEN):
                tokens.append(token)
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.FUNC, Token.CLOSE):
                token = Token(')', Token.PAREN, Token.CLOSE)

            elif token.matches(Token.ARRAY, Token.OPEN):
                tokens.append(token)
                tokens.append(Token('(', Token.PAREN, Token.OPEN))
                tokens.append(Token('', Token.ARRAYROW, Token.OPEN))
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.ARRAY, Token.CLOSE):
                tokens.append(token)
                token = Token(')', Token.PAREN, Token.CLOSE)

            elif token.matches(Token.SEP, Token.ROW):
                tokens.append(Token(')', Token.PAREN, Token.CLOSE))
                tokens.append(Token(',', Token.SEP, Token.ARG))
                tokens.append(Token('', Token.ARRAYROW, Token.OPEN))
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.PAREN, Token.OPEN):
                token.value = '('

            elif token.matches(Token.PAREN, Token.CLOSE):
                token.value = ')'

            if token:
                tokens.append(token)

        output = []
        stack = []
        were_values = []
        arg_count = []

        for token in tokens:
            if token.type == token.OPERAND:

                output.append(ASTNode.create(token))
                if were_values:
                    were_values.pop()
                    were_values.append(True)

            elif token.type != token.PAREN and token.subtype == token.OPEN:

                if token.type in (token.ARRAY, Token.ARRAYROW):
                    token = Token(token.type, token.type, token.subtype)

                stack.append(token)
                arg_count.append(0)
                if were_values:
                    were_values.pop()
                    were_values.append(True)
                were_values.append(False)

            elif token.type == token.SEP:

                while stack and (stack[-1].subtype != token.OPEN):
                    output.append(ASTNode.create(stack.pop()))

                if were_values.pop():
                    arg_count[-1] += 1
                were_values.append(False)

                if not len(stack):
                    raise CompilerError("Mismatched or misplaced parentheses")

            elif token.is_operator:

                if token.type == token.OP_PRE and token.value == "-":
                    assert token.type == token.OP_PRE
                    o1 = cls.operators['u-']
                else:
                    o1 = cls.operators[token.value]

                while stack and stack[-1].is_operator:

                    if stack[-1].matches(type_=Token.OP_PRE, value='-'):
                        o2 = cls.operators['u-']
                    else:
                        o2 = cls.operators[stack[-1].value]

                    if o1 < o2:
                        output.append(ASTNode.create(stack.pop()))
                    else:
                        break

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            elif token.subtype == token.CLOSE:

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(ASTNode.create(stack.pop()))

                if not stack:
                    raise CompilerError("Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = ASTNode.create(stack.pop())
                    a = arg_count.pop()
                    w = were_values.pop()
                    if w:
                        a += 1
                    f.num_args = a
                    output.append(f)

        while stack:
            if stack[-1].subtype in (Token.OPEN, Token.CLOSE):
                raise CompilerError("Mismatched or misplaced parentheses")

            output.append(ASTNode.create(stack.pop()))

        return output

    @classmethod
    def build_ast(cls, rpn_expression):
        """build an AST from an Excel formula in reverse polish notation"""

        # use a directed graph to store the syntax tree
        ast = DiGraph()

        stack = []

        for node in rpn_expression:
            # The graph does not maintain the order of adding nodes/edges, so
            # add an attribute 'pos' so we can always sort to the correct order

            node.ast = ast
            if isinstance(node, OperatorNode):
                if node.token.type == node.token.OP_IN:
                    arg2 = stack.pop()
                    arg1 = stack.pop()
                    ast.add_node(arg1, pos=0)
                    ast.add_node(arg2, pos=1)
                    ast.add_edge(arg1, node)
                    ast.add_edge(arg2, node)
                else:
                    arg1 = stack.pop()
                    ast.add_node(arg1, pos=1)
                    ast.add_edge(arg1, node)

            elif isinstance(node, FunctionNode):
                if node.num_args:
                    args = stack[-node.num_args:]
                    del stack[-node.num_args:]
                    for i, a in enumerate(args):
                        ast.add_node(a, pos=i)
                        ast.add_edge(a, node)
            else:
                ast.add_node(node, pos=0)

            stack.append(node)

        assert 1 == len(stack)
        return stack[0]

    def cell2code(self, cell):
        """Generate python code for the given cell"""
        if cell.formula:
            ast_root = self.build_ast(
                self.parse_to_rpn(cell.formula or str(cell.value)))
            code = ast_root.emit(ASTNode.Context(cell, self.excel))
        else:
            ast_root = None
            if isinstance(cell.value, str):
                code = '"{}"'.format(cell.value)
            else:
                code = str(cell.value)

        # set the code & compile (will flag problems sooner rather than later)
        cell.python_expression = code
        cell.compile()

        return ast_root

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
        seeds = Cell.make_cells(self.excel, seed, sheet=cursheet)[0]

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

            # parse the formula into code
            cell_ast = self.cell2code(c1)

            # get all the cells/ranges this formula refers to, and remove dupes
            dependants = uniqueify(
                x.value.replace('$', '') for x, *_ in cell_ast.descendents
                if isinstance(x, RangeNode)
            )

            for dep in dependants:

                # if the dependency is a multi-cell range, create a range object
                if is_range(dep):
                    # this will make sure we always have an absolute address
                    rng = CellRange(dep, sheet=cursheet)

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
                        cells, nrows, ncols = Cell.make_cells(
                            self.excel, dep, sheet=cursheet)

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
                    cells = [Cell.resolve_cell(self.excel, dep, sheet=cursheet)]
                    target = self.cellmap[c1.address()]

                # process each cell
                for c2 in flatten(cells):
                    # if we haven't treated this cell already
                    if c2.address() not in self.cellmap:
                        if c2.formula:
                            # cell with a formula, add to the todo list
                            todo.append(c2)
                        else:
                            # constant cell, no need for further processing
                            self.cell2code(c2)

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
