import ast
import re

import openpyxl.formula.tokenizer as tokenizer
from networkx.classes.digraph import DiGraph
from pycel.excelutil import (
    AddressRange,
    DIV0,
    get_linest_degree,
    uniqueify,
)

EVAL_REGEX = re.compile(r'(_C_|_R_)(\([^)]*\))')


class FormulaParserError(Exception):
    """Base class for Parser errors"""


class CompilerError(Exception):
    """Base class for Compiler errors"""


class Tokenizer(tokenizer.Tokenizer):
    """Amend openpyxl tokenizer"""

    def __init__(self, formula):
        super(Tokenizer, self).__init__(formula)
        self.items = self._items()

    def _items(self):
        """Convert to use our Token"""
        t = [None] + [Token.from_token(t) for t in self.items] + [None]

        # convert or remove unneeded whitespace
        tokens = []
        for prev_token, token, next_token in zip(t, t[1:], t[2:]):
            if token.type != Token.WSPACE or not prev_token or not next_token:
                tokens.append(token)

            elif (
                prev_token.matches(type_=Token.FUNC, subtype=Token.CLOSE) or
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
    """Amend openpyxl token"""

    INTERSECT = "INTERSECT"
    ARRAYROW = "ARRAYROW"

    class Precedence:
        """Small wrapper class to manage operator precedence during parsing"""

        def __init__(self, precedence, associativity):
            self.precedence = precedence
            self.associativity = associativity

        def __lt__(self, other):
            return (self.precedence < other.precedence or
                    self.associativity == "left" and
                    self.precedence == other.precedence
                    )

    precedences = {
        # http://office.microsoft.com/en-us/excel-help/
        #   calculation-operators-and-precedence-HP010078886.aspx
        ':': Precedence(8, 'left'),
        ' ': Precedence(8, 'left'),  # range intersection
        ',': Precedence(8, 'left'),
        'u': Precedence(7, 'left'),  # unary operator
        '%': Precedence(6, 'left'),
        '^': Precedence(5, 'left'),
        '*': Precedence(4, 'left'),
        '/': Precedence(4, 'left'),
        '+': Precedence(3, 'left'),
        '-': Precedence(3, 'left'),
        '&': Precedence(2, 'left'),
        '=': Precedence(1, 'left'),
        '<': Precedence(1, 'left'),
        '>': Precedence(1, 'left'),
        '<=': Precedence(1, 'left'),
        '>=': Precedence(1, 'left'),
        '<>': Precedence(1, 'left'),
    }

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

    @property
    def precedence(self):
        assert self.is_operator
        return self.precedences[
            'u' if self.type == Token.OP_PRE else self.value]


class ASTNode(object):
    """A generic node in the AST used to compile a cell's formula"""

    def __init__(self, token, cell=None):
        super(ASTNode, self).__init__()
        self.token = token
        self.cell = cell
        self._ast = None
        self._parent = None
        self._children = None
        self._descendants = None

    @classmethod
    def create(cls, token, cell=None):
        """Simple factory function"""
        if token.type == Token.OPERAND:
            if token.subtype == Token.RANGE:
                return RangeNode(token, cell)
            else:
                return OperandNode(token, cell)

        elif token.is_funcopen:
            return FunctionNode(token, cell)

        elif token.is_operator:
            return OperatorNode(token, cell)

        raise FormulaParserError('Unknown token type: {}'.format(repr(token)))

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
    def descendants(self):
        if self._descendants is None:
            self._descendants = list(
                n for n in self.ast.nodes(self) if n[0] != self)
        return self._descendants

    @property
    def parent(self):
        if self._parent is None:
            self._parent = next(self.ast.successors(self), None)
        return self._parent

    def emit(self):
        """Emit code"""
        return self.value


class OperatorNode(ASTNode):
    op_map = {
        # convert the operator to python equivalents
        "^": "**",
        "=": "==",
        "<>": "!=",
        " ": "+"  # range intersection
    }

    def emit(self):
        xop = self.value

        # Get the arguments
        args = self.children

        op = self.op_map.get(xop, xop)

        if self.type == Token.OP_PRE:
            return "-" + args[0].emit()

        parent = self.parent
        # don't render the ^{1,2,..} part in a linest formula
        # TODO: bit of a hack
        if op == "**":
            if parent and parent.value.lower() == "linest(":
                return args[0].emit()

        if op != ',':
            op = ' ' + op
        ss = '{}{} {}'.format(
            args[0].emit(), op, args[1].emit())

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "(" + ss + ")"

        return ss


class OperandNode(ASTNode):

    def emit(self):
        if self.subtype == self.token.LOGICAL:
            return str(self.value.lower() == "true")

        elif self.subtype in ("TEXT", "ERROR") and len(self.value) > 2:
            # if the string contains quotes, escape them
            return '"{}"'.format(self.value.replace('""', '\\"').strip('"'))

        else:
            return self.value


class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""

    def emit(self):
        # resolve the range into cells
        sheet = self.cell and self.cell.sheet or ''
        if '!' in self.value:
            sheet = ''
        address = AddressRange.create(
            self.value.replace('$', ''), sheet=sheet, cell=self.cell)
        template = '_R_("{}")' if address.is_range else '_C_("{}")'
        return template.format(address)


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    """
    A dictionary that maps excel function names onto python equivalents. You
    should only add an entry to this map if the python name is different from
    the excel name (which it may need to be to prevent conflicts with
    existing python functions with that name, e.g., max).

    So if excel defines a function foobar(), all you have to do is add a
    function called foobar to this module.  You only need to add it to the
    function map, if you want to use a different name in the python code.

    Note: some functions (if, pi, atan2, and, or, array, ...) are already taken
    care of in the FunctionNode code, so adding them here will have no effect.
    """

    # dict of excel equivalent functions
    func_map = {
        "ln": "xlog",
        "min": "xmin",
        "max": "xmax",
        "sum": "xsum",
        "gammaln": "lgamma",
        "round": "xround"
    }

    def __init__(self, *args):
        super(FunctionNode, self).__init__(*args)
        self.num_args = 0

    def comma_join_emit(self, fmt_str=None):
        if fmt_str is None:
            return ", ".join(n.emit() for n in self.children)
        else:
            return ", ".join(
                fmt_str.format(n.emit()) for n in self.children)

    def emit(self):
        func = self.value.lower().strip('(')

        # if a special handler is needed
        handler = getattr(self, 'func_{}'.format(func), None)
        if handler is not None:
            return handler()

        else:
            # map to the correct name
            return "{}({})".format(
                self.func_map.get(func, func), self.comma_join_emit())

    def func_atan2(self):
        # swap arguments
        a1, a2 = (a.emit() for a in self.children)
        return "atan2({}, {})".format(a2, a1)

    @staticmethod
    def func_pi():
        # constant, no parens
        return "pi"

    def func_if(self):
        # inline the if
        args = [c.emit() for c in self.children] + [0]
        if len(args) in (3, 4):
            return "({} if {} else {})".format(args[1], args[0], args[2])

        raise FormulaParserError(
            "IF with %s arguments not supported".format(len(args) - 1))

    def func_array(self):
        if len(self.children) == 1:
            return '[{}]'.format(self.children[0].emit())
        else:
            # multiple rows
            return '[{}]'.format(self.comma_join_emit('[{}]'))

    def func_arrayrow(self):
        # simply create a list
        return self.comma_join_emit()

    def func_linest(self):
        func = self.value.lower().strip('(')
        code = '{}({}'.format(func, self.comma_join_emit())

        if not self.cell or not self.cell.excel:
            degree, coef = -1, -1
        else:
            # linests are often used as part of an array formula spanning
            # multiple cells, one cell for each coefficient.  We have to
            # figure out where we currently are in that range.
            degree, coef = get_linest_degree(self.cell)

        # if we are the only linest (degree is one) and linest is nested
        # return vector, else return the coef.
        if func == "linest":
            code += ", degree=%s)" % degree
        else:
            code += ")"

        if not (degree == 1 and self.parent):
            code += "[%s]" % (coef - 1)

        return code

    func_linestmario = func_linest

    def func_and(self):
        return "all(({},))".format(self.comma_join_emit())

    def func_or(self):
        return "any(({},))".format(self.comma_join_emit())


class ExcelFormula(object):
    """Take an Excel formula and compile it to Python code."""

    def __init__(self, formula, cell=None, formula_is_python_code=False):
        if formula_is_python_code:
            self.base_formula = None
            self._python_code = formula[1:]
        else:
            self.base_formula = formula
            self._python_code = None

        self.cell = cell
        self.lineno = 1
        self._rpn = None
        self._ast = None
        self._needed_addresses = None
        self._compiled_python = None

    def __str__(self):
        return self.base_formula or self.python_code

    def __getstate__(self):
        # code objects are not serializable
        state = dict(self.__dict__)
        for to_remove in '_compiled_python _ast _rpn _dep_graph'.split():
            if to_remove in state:
                state[to_remove] = None
        return state

    @property
    def rpn(self):
        if self._rpn is None:
            self._rpn = self.parse_to_rpn(self.base_formula)
        return self._rpn

    @property
    def ast(self):
        if self._ast is None and self.rpn:
            self._ast = self.build_ast(self.rpn)
        return self._ast

    @property
    def needed_addresses(self):
        """Return the addresses and address ranges this formula needs"""
        if self._needed_addresses is None:
            # get all the cells/ranges this formula refers to, and remove dupes
            if self.python_code:
                self._needed_addresses = uniqueify(
                    AddressRange(eval_call[1][2:-2])
                    for eval_call in EVAL_REGEX.findall(self.python_code)
                )
            else:
                self._needed_addresses = ()

        return self._needed_addresses

    @property
    def python_code(self):
        """Use the ast to generate python code"""
        if self._python_code is None:
            if self.ast is None:
                self._python_code = ''
            else:
                self._python_code = self.ast.emit()
        return self._python_code

    @property
    def compiled_python(self):
        """ Using the Python code, generate compiled python code"""
        if self._compiled_python is None and self.python_code:
            try:
                self._compile_python_ast()
            except Exception as exc:
                raise FormulaParserError(
                    "Failed to compile expression {}: {}".format(
                        self.python_code, exc))

        return self._compiled_python

    def ast_node(self, token):
        return ASTNode.create(token, self.cell)

    def parse_to_rpn(self, expression):
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

            tokens.append(token)

        output = []
        stack = []
        were_values = []
        arg_count = []

        for token in tokens:
            if token.type == token.OPERAND:

                output.append(self.ast_node(token))
                if were_values:
                    were_values[-1] = True

            elif token.type != token.PAREN and token.subtype == token.OPEN:

                if token.type in (token.ARRAY, Token.ARRAYROW):
                    token = Token(token.type, token.type, token.subtype)

                stack.append(token)
                arg_count.append(0)
                if were_values:
                    were_values[-1] = True
                were_values.append(False)

            elif token.type == token.SEP:

                while stack and (stack[-1].subtype != token.OPEN):
                    output.append(self.ast_node(stack.pop()))

                if not len(were_values):
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                were_values.pop()
                arg_count[-1] += 1
                were_values.append(False)

            elif token.is_operator:

                while stack and stack[-1].is_operator:
                    if token.precedence < stack[-1].precedence:
                        output.append(self.ast_node(stack.pop()))
                    else:
                        break

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            else:
                assert token.subtype == token.CLOSE

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(self.ast_node(stack.pop()))

                if not stack:
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = self.ast_node(stack.pop())
                    f.num_args = arg_count.pop() + int(were_values.pop())
                    output.append(f)

        while stack:
            if stack[-1].subtype in (Token.OPEN, Token.CLOSE):
                raise FormulaParserError("Mismatched or misplaced parentheses")

            output.append(self.ast_node(stack.pop()))

        return output

    @classmethod
    def build_ast(cls, rpn_expression):
        """build an AST from an Excel formula

        :param rpn_expression: a string formula or the result of parse_to_rpn()
        :return: AST which can be used to generate code
        """

        # use a directed graph to store the syntax tree
        tree = DiGraph()

        # production stack
        stack = []

        for node in rpn_expression:
            # The graph does not maintain the order of adding nodes/edges, so
            # add an attribute 'pos' so we can always sort to the correct order

            node.ast = tree
            if isinstance(node, OperatorNode):
                if node.token.type == node.token.OP_IN:
                    try:
                        arg2 = stack.pop()
                        arg1 = stack.pop()
                    except IndexError:
                        raise FormulaParserError(
                            "'{}' operator missing operand".format(
                                node.token.value))
                    tree.add_node(arg1, pos=0)
                    tree.add_node(arg2, pos=1)
                    tree.add_edge(arg1, node)
                    tree.add_edge(arg2, node)
                else:
                    try:
                        arg1 = stack.pop()
                    except IndexError:
                        raise FormulaParserError(
                            "'{}' operator missing operand".format(
                                node.token.value))
                    tree.add_node(arg1, pos=1)
                    tree.add_edge(arg1, node)

            elif isinstance(node, FunctionNode):
                if node.num_args:
                    args = stack[-node.num_args:]
                    del stack[-node.num_args:]
                    for i, a in enumerate(args):
                        tree.add_node(a, pos=i)
                        tree.add_edge(a, node)
            else:
                tree.add_node(node, pos=0)

            stack.append(node)

        assert 1 == len(stack)
        return stack[0]

    @classmethod
    def build_eval_context(cls, evaluate, evaluate_range):
        """eval with namespace management.  Will auto import needed functions

        Used like:

            build_eval(...)(expression returned from build_python)

        :param evaluate: a function to evaluate a cell address
        :param evaluate_range: a function to evaluate a range address
        :return: a function to evaluate a compiled expression from build_ast
        """

        import importlib

        modules = (
            importlib.import_module('pycel.excellib'),
            importlib.import_module('math'),
        )

        def eval_func(excel_formula):

            # the compiled expressions can call these functions if
            # referencing other cells or a range of cells

            def report_error():
                import traceback
                raise CompilerError("{}    {}".format(
                    traceback.format_exc(), excel_formula.python_code))

            def _C_(address):
                return evaluate(address)

            def _R_(rng):
                return evaluate_range(rng)

            def load_func(func_name):
                # if a function is not already loaded, load it
                funcs = (getattr(module, func_name, None) for module in modules)
                return next((f for f in funcs if f is not None), None)

            while True:
                try:
                    return eval(excel_formula.compiled_python)
                except NameError as exc:
                    name = str(exc).split("'")[1]
                    func = load_func(name)
                    if func is None:
                        report_error()
                    locals()[name] = func

                except ZeroDivisionError:
                    return DIV0

                except Exception:
                    report_error()

        return eval_func

    def _compile_python_ast(self):
        kwargs = dict(
            mode='eval',
            filename=str(self.cell.address) if self.cell else self.python_code,
        )
        tree = ast.parse(self.python_code, **kwargs)
        ast.increment_lineno(tree, self.lineno - 1)

        # edit the ast with a few changes to be more excel like

        class OperatorWrapper(ast.NodeTransformer):
            """Apply excel consistent type conversions"""

            def visit_Compare(self, node):
                """ change the compare node to a function node """
                node = ast.NodeTransformer.generic_visit(self, node)

                left = node.left
                op = ast.Str(s=type(node.ops[0]).__name__)
                right = node.comparators[0]

                return ast.Call(
                    func=ast.Name(id='excel_operator_operand_fixup',
                                  ctx=ast.Load()),
                    args=[left, op, right],
                    keywords=[],
                    lineno=node.lineno,
                    col_offset=node.col_offset,
                )

            def visit_BinOp(self, node):
                """ change the compare node to a function node """
                node = ast.NodeTransformer.generic_visit(self, node)

                left = node.left
                op = ast.Str(s=type(node.op).__name__)
                right = node.right

                return ast.Call(
                    func=ast.Name(id='excel_operator_operand_fixup',
                                  ctx=ast.Load()),
                    args=[left, op, right],
                    keywords=[],
                    lineno=node.lineno,
                    col_offset=node.col_offset,
                )

        # modify the ast tree to convert Compare and BinOp to Call
        tree = ast.fix_missing_locations(OperatorWrapper().visit(tree))

        # compile the tree
        self._compiled_python = compile(tree, **kwargs)
