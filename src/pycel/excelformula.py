import ast
import logging
import marshal
import math
import re
import sys

import openpyxl.formula.tokenizer as tokenizer
from networkx.classes.digraph import DiGraph
from networkx.exception import NetworkXError
from pycel.excelutil import (
    AddressRange,
    build_operator_operand_fixup,
    EMPTY,
    ERROR_CODES,
    get_linest_degree,
    math_wrap,
    PyCelException,
    uniqueify,
)

EVAL_REGEX = re.compile(r'(_C_|_R_)(\([^)]*\))')


class FormulaParserError(PyCelException):
    """Error during parsing"""


class FormulaEvalError(PyCelException):
    """Error during eval"""


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
                # drop unary +
                if not token.matches(type_=Token.OP_PRE, value='+'):
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

        # ::TODO:: remove after openpyxl updated
        # ::HACK:: to workaround openpyxl issue fixed in PR #301
        token_stream = iter(tokens)
        tokens = []
        for token in token_stream:
            if token.type == Token.OPERAND and (
                token.value.count('[') != token.value.count(']')
            ):
                new_value = (token.value +
                             next(token_stream).value +
                             next(token_stream).value)
                token = Token(new_value, token.type, token.subtype)
            tokens.append(token)

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


class ASTNode:
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
            try:
                args = self.ast.predecessors(self)
            except NetworkXError:
                args = []
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

    @property
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

    @property
    def emit(self):
        xop = self.value

        # Get the arguments
        args = self.children

        op = self.op_map.get(xop, xop)

        if self.type == Token.OP_PRE:
            return self.value + args[0].emit

        parent = self.parent
        # don't render the ^{1,2,..} part in a linest formula
        # TODO: bit of a hack
        if op == "**":
            if parent and parent.value.lower() == "linest(":
                return args[0].emit

        if op == '%':
            ss = '{} / 100'.format(args[0].emit)
        else:
            if op != ',':
                op = ' ' + op

            ss = '{}{} {}'.format(args[0].emit, op, args[1].emit)

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "(" + ss + ")"

        return ss


class OperandNode(ASTNode):

    @property
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

    @property
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

    Note: some functions (if, pi, and, or, array, ...) are already taken
    care of in the FunctionNode code, so adding them here will have no effect.
    """

    # dict of excel equivalent functions
    func_map = {
        "atan2": "xatan2",
        "gammaln": "lgamma",
        "if": "xif",
        "len": "xlen",
        "ln": "xlog",
        "max": "xmax",
        "min": "xmin",
        "round": "xround",
        "sum": "xsum",
    }

    def __init__(self, *args):
        super(FunctionNode, self).__init__(*args)
        self.num_args = 0

    def comma_join_emit(self, fmt_str=None):
        if fmt_str is None:
            return ", ".join(n.emit for n in self.children)
        else:
            return ", ".join(
                fmt_str.format(n.emit) for n in self.children)

    @property
    def emit(self):
        func = self.value.lower().strip('(')

        if func.startswith('_xlfn.'):
            func = func[6:]
        func = func.replace('.', '_')

        # if a special handler is needed
        handler = getattr(self, 'func_{}'.format(func), None)
        if handler is not None:
            return handler()

        else:
            # map to the correct name
            return "{}({})".format(
                self.func_map.get(func, func), self.comma_join_emit())

    @staticmethod
    def func_pi():
        # constant, no parens
        return "pi"

    def func_array(self):
        if len(self.children) == 1:
            return '[{}]'.format(self.children[0].emit)
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

        if not (degree == 1 and self.parent):  # pragma: no branch
            code += "[%s]" % (coef - 1)

        return code

    func_linestmario = func_linest

    def func_and(self):
        return "all(({},))".format(self.comma_join_emit())

    def func_or(self):
        return "any(({},))".format(self.comma_join_emit())

    def func_row(self):
        assert len(self.children) <= 1
        if len(self.children) == 0:
            address = '_REF_("{}")'.format(self.cell.address)
        else:
            address = self.children[0].emit
            address = address.replace('_R_', '_REF_').replace('_C_', '_REF_')
        return 'row({})'.format(address)

    def func_column(self):
        assert len(self.children) <= 1
        if len(self.children) == 0:
            address = '_REF_("{}")'.format(self.cell.address)
        else:
            address = self.children[0].emit
            address = address.replace('_R_', '_REF_').replace('_C_', '_REF_')
        return 'column({})'.format(address)


class ExcelFormula:
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
        self.filename = ''

        self._rpn = None
        self._ast = None
        self._needed_addresses = None
        self._compiled_python = None
        self._marshalled_python = None
        self.compiled_lambda = None

    def __str__(self):
        return self.base_formula or self.python_code

    def __getstate__(self):
        # build the python code
        self.python_code

        # Throw everything away except the python code
        state = dict(self.__dict__)
        remove_names = 'compiled_lambda _compiled_python _ast _rpn ' \
                       'base_formula _needed_addresses'
        for to_remove in remove_names.split():
            if to_remove in state:  # pragma: no branch
                state[to_remove] = None
        return state

    @property
    def rpn(self):
        if self._rpn is None:
            self._rpn = self._parse_to_rpn(self.base_formula)
        return self._rpn

    @property
    def ast(self):
        if self._ast is None and self.rpn:
            self._ast = self._build_ast(self.rpn)
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
                self._python_code = self.ast.emit
        return self._python_code

    @property
    def compiled_python(self):
        """ Using the Python code, generate compiled python code"""
        if self._compiled_python is None and self.python_code:
            if self._marshalled_python is not None:
                try:
                    marshalled, names = self._marshalled_python
                    self._compiled_python = marshal.loads(marshalled), names
                except Exception:
                    self._marshalled_python = None
                    return self.compiled_python
            else:
                try:
                    self._compile_python_ast()
                except Exception as exc:
                    raise FormulaParserError(
                        "Failed to compile expression {}: {}".format(
                            self.python_code, exc))

        return self._compiled_python

    def _ast_node(self, token):
        return ASTNode.create(token, self.cell)

    def _parse_to_rpn(self, expression):
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

                output.append(self._ast_node(token))
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
                    output.append(self._ast_node(stack.pop()))

                if not len(were_values):
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                were_values.pop()
                arg_count[-1] += 1
                were_values.append(False)

            elif token.is_operator:

                while stack and stack[-1].is_operator:
                    if token.precedence < stack[-1].precedence:
                        output.append(self._ast_node(stack.pop()))
                    else:
                        break

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            else:
                assert token.subtype == token.CLOSE

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(self._ast_node(stack.pop()))

                if not stack:
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = self._ast_node(stack.pop())
                    f.num_args = arg_count.pop() + int(were_values.pop())
                    output.append(f)

        while stack:
            if stack[-1].subtype in (Token.OPEN, Token.CLOSE):
                raise FormulaParserError("Mismatched or misplaced parentheses")

            output.append(self._ast_node(stack.pop()))

        return output

    @classmethod
    def _build_ast(cls, rpn_expression):
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
    def build_eval_context(cls, evaluate, evaluate_range, logger=None):
        """eval with namespace management.  Will auto import needed functions

        Used like:

            build_eval(...)(expression returned from build_python)

        :param evaluate: a function to evaluate a cell address
        :param evaluate_range: a function to evaluate a range address
        :param logger: a looger to use (defaults to pycel)
        :return: a function to evaluate a compiled expression from build_ast
        """

        import importlib

        modules = (
            importlib.import_module('pycel.excellib'),
            importlib.import_module('pycel.lib.binary'),
            importlib.import_module('math'),
        )

        logger = logger or logging.getLogger('pycel')
        error_messages = []

        def capture_error_state(is_exception, msg):
            if is_exception:
                import traceback
                try:
                    trace = traceback.format_exc()
                except AttributeError:  # pragma: no cover
                    # this is a ::HACK:: to work around PY34
                    trace = ''
            else:
                trace = ''  # pragma: no cover
            error_messages.append((trace, msg))

        def error_logger(level, python_code, msg=None, exc=None):
            """ Log a traceback, a msg, and reraise if asked

            :param level: level for the logger "error", "warning", "debug"...
            :param python_code: Code which caused the error
            :param msg: Additional information for logging
            :param exc: An exception to reraise, if desired
            :return: the constructed error message if not reraising
            """
            if exc:
                capture_error_state(exc, msg)
                assert 1 == len(error_messages)
            trace, msg = error_messages.pop()
            fmt_str = "{0}Eval: {1}" if msg is None else "{0}Eval: {1}\n{2}"
            error_msg = fmt_str.format(trace, python_code, msg)
            getattr(logger, level)(error_msg)
            if exc is not None:
                raise exc(error_msg)
            return error_msg

        def load_function(excel_formula, name_space):
            """exec the code into our address space"""

            # the compiled expressions can call these functions if
            # referencing other cells or a range of cells
            name_space['_C_'] = evaluate
            name_space['_R_'] = evaluate_range
            name_space['_REF_'] = AddressRange.create
            name_space['pi'] = math.pi

            for name in ('int', 'abs', 'round'):
                name_space[name] = math_wrap(
                    globals()['__builtins__'][name])

            # function to fixup the operands
            name_space['excel_operator_operand_fixup'] = \
                build_operator_operand_fixup(capture_error_state)

            # hook for the execed code to save the resulting lambda
            name_space['lambdas'] = lambdas = []

            # get the compiled code and needed names
            compiled, names = excel_formula.compiled_python

            # load the needed names
            for name in names:
                if name not in name_space:
                    funcs = ((getattr(module, name, None), module)
                             for module in modules)
                    func, module = next(
                        (f for f in funcs if f[0] is not None), (None, None))
                    if func is not None:
                        if module.__name__ == 'math':
                            name_space[name] = math_wrap(func)
                        else:
                            name_space[name] = func

            # exec the code to define the lambda
            exec(compiled, name_space, name_space)
            excel_formula.compiled_lambda = lambdas[0]
            del name_space['lambdas']

        def eval_func(excel_formula):
            """ Call the compiled lambda to evaluate the cell """

            if excel_formula.compiled_lambda is None:
                load_function(excel_formula, locals())

            try:
                ret_val = excel_formula.compiled_lambda()

            except Exception:
                error_logger('error', excel_formula.python_code,
                             exc=FormulaEvalError)

            if error_messages:
                level = 'warning' if ret_val in ERROR_CODES else 'info'
                error_logger(level, excel_formula.python_code)

            return ret_val if ret_val not in (None, EMPTY) else 0

        return eval_func

    def _compile_python_ast(self):
        """ Compile the python code into a lambda for execution

        ### Traceback will show this line if not loaded from a text file

        If the compiler has been loaded from (json, yaml, etc) then python
        expression will be shown in any tracebacks instead of the above
        """
        local_line = sys._getframe().f_lineno - 6

        source_code = "lambdas.append(lambda: {})".format(self.python_code)
        kwargs = dict(mode='exec', filename=self.filename or __file__)
        tree = ast.parse(source_code, **kwargs)
        ast.increment_lineno(tree, (self.lineno - 1) or local_line)

        names = set()

        # edit the ast with a few changes to be more excel like

        class OperatorWrapper(ast.NodeTransformer):
            """Apply excel consistent type conversions, fetch dependant names"""

            def visit_Name(self, node):
                """ Gather up all names needed """
                node = ast.NodeTransformer.generic_visit(self, node)
                names.add(node.id)
                return node

            def visit_Compare(self, node):
                """ change the compare node to a function node """
                node = ast.NodeTransformer.generic_visit(self, node)
                return self.replace_op(
                    node, node.left, node.ops[0], node.comparators[0])

            def visit_BinOp(self, node):
                """ change the BinOP node to a function node """
                node = ast.NodeTransformer.generic_visit(self, node)
                return self.replace_op(node, node.left, node.op, node.right)

            def visit_UnaryOp(self, node):
                """ change the UnaryOp node to a function node """
                node = ast.NodeTransformer.generic_visit(self, node)
                left = ast.Str(EMPTY)
                return self.replace_op(node, left, node.op, node.operand)

            def replace_op(self, node, left, node_op, right):
                """ change the compare node to a function node """

                op = ast.Str(s=type(node_op).__name__)
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
        self._compiled_python = compile(tree, **kwargs), names
        self._marshalled_python = marshal.dumps(self._compiled_python[0]), names
