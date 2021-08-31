# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import ast
import importlib
import logging
import marshal
import math
import sys
import tokenize as tk

import openpyxl.formula.tokenizer as tokenizer
from networkx.classes.digraph import DiGraph
from networkx.exception import NetworkXError

from pycel.excelutil import (
    AddressMultiAreaRange,
    AddressRange,
    build_operator_operand_fixup,
    coerce_to_number,
    EMPTY,
    ERROR_CODES,
    in_array_formula_context,
    NAME_ERROR,
    PyCelException,
    uniqueify,
)
from pycel.lib.function_helpers import load_functions
from pycel.lib.function_info import func_status_msg


ADDR_FUNCS_NAMES = '_R_', '_C_', '_REF_'


class FormulaParserError(PyCelException):
    """Error during parsing"""


class UnknownFunction(PyCelException):
    """Functions unknown to PyCel"""


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
                # ::HACK:: this is code to make the tokenizer behave like
                # this change to the openpyxl tokenizer.
                # https://bitbucket.org/openpyxl/openpyxl/pull-requests/345
                # If the pull request gets merged, we can the update our
                # openpyxl requirements and remove this code.
                if (token.matches(type_=Token.FUNC, subtype=Token.OPEN) and
                        ':' in token.value):

                    # split the address on the ':'
                    addr, func = token.value.rsplit(':', maxsplit=1)
                    tokens.append(Token(addr, Token.OPERAND, Token.RANGE))
                    tokens.append(Token(':', Token.OP_IN, ''))
                    token.value = func
                    tokens.append(token)

                elif (token.matches(type_=Token.OPERAND,
                                    subtype=Token.RANGE) and
                      token.value.startswith(':')):
                    # split the address on the ':'
                    tokens.append(Token(':', Token.OP_IN, ''))
                    token.value = token.value[1:]
                    tokens.append(token)

                # drop unary +
                elif not token.matches(type_=Token.OP_PRE, value='+'):
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
    EMPTY = "EMPTY"

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
        'u': Precedence(7, 'right'),  # unary operator
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

        raise FormulaParserError(f'Unknown token type: {repr(token)}')

    def __str__(self):
        return str(self.token.value.strip('('))

    def __repr__(self):
        return f"{type(self).__name__}<{self.token.value.strip('(')}>"

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
            self._children = sorted(
                args, key=lambda x: self.ast.nodes[x]['pos'])
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
        if op == '%':
            ss = f'{args[0].emit} / 100'
        elif op == ' ':
            # range intersection
            ss = '_R_' + (f'(str({args[0].emit} & {args[1].emit}))'
                          .replace('_R_', '_REF_')
                          .replace('_C_', '_REF_')
                          )
        elif op == ':':
            # range union
            ss = '_R_' + (f'(str({args[0].emit} ** {args[1].emit}))'
                          .replace('_R_', '_REF_')
                          .replace('_C_', '_REF_')
                          )
        else:
            if op != ',':
                op = ' ' + op
            ss = f'{args[0].emit}{op} {args[1].emit}'

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "(" + ss + ")"

        return ss


class OperandNode(ASTNode):

    @property
    def emit(self):
        if self.subtype == self.token.LOGICAL:
            return str(self.value.lower() == "true")

        elif self.subtype == self.token.EMPTY:
            return 'None'

        elif self.subtype in ("TEXT", "ERROR") and len(self.value) > 2:
            # if the string contains quotes, escape them
            value = self.value
            if value.startswith('"') and value.endswith('"'):
                value = value[1:-1]
            value = value.replace('""', r'\"')
            return f'"{value}"'

        else:
            return self.value


class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""

    @property
    def emit(self):
        return self._emit()

    def _emit(self, value=None):
        # resolve the range into cells
        sheet = self.cell and self.cell.sheet or ''
        value = value is not None and value or self.value
        if '!' in value:
            sheet = ''
        try:
            addr_str = value.replace('$', '')
            address = AddressRange.create(addr_str, sheet=sheet, cell=self.cell)
        except ValueError:
            # check for table relative address
            table_name = None
            if self.cell:
                excel = self.cell.excel
                if excel and '[' in addr_str:
                    table_name = excel.table_name_containing(self.cell.address)

            if not table_name:
                logging.getLogger('pycel').warning(f'Table Name not found: {addr_str}')
                return f'"{NAME_ERROR}"'

            addr_str = f'{table_name}{addr_str}'
            address = AddressRange.create(
                addr_str, sheet=self.cell.address.sheet, cell=self.cell)

        if isinstance(address, AddressMultiAreaRange):
            return ', '.join(self._emit(value=str(addr)) for addr in address)
        else:
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
        "abs": "abs_",
        "and": "and_",
        "atan2": "atan2_",
        "if": "if_",
        "int": "int_",
        "len": "len_",
        "max": "max_",
        "not": "not_",
        "or": "or_",
        "min": "min_",
        "round": "round_",
        "sum": "sum_",
        "xor": "xor_",
    }

    def __init__(self, *args):
        super(FunctionNode, self).__init__(*args)
        self.num_args = 0

    def comma_join_emit(self, fmt_str=None, to_emit=None):
        if to_emit is None:
            to_emit = self.children
        if fmt_str is None:
            return ", ".join(n.emit for n in to_emit)
        else:
            return ", ".join(fmt_str.format(n.emit) for n in to_emit)

    @property
    def emit(self):
        func = self.value.lower().strip('(')

        if func and func[0] == func[-1] == '_':
            func = func.upper()
        if func.startswith('_xlfn.'):
            func = func[6:]
        func = func.replace('.', '_')

        # if a special handler is needed
        handler = getattr(self, f'func_{func}', None)
        if handler is not None:
            return handler()
        else:
            # map to the correct name
            return f"{self.func_map.get(func, func)}({self.comma_join_emit()})"

    @staticmethod
    def func_pi():
        # constant, no parens
        return "pi"

    @staticmethod
    def func_true():
        # constant, no parens
        return "True"

    @staticmethod
    def func_false():
        # constant, no parens
        return "False"

    def func_array(self):
        return f"({self.comma_join_emit(fmt_str='({},)')},)"

    def func_arrayrow(self):
        # simply create a list
        return self.comma_join_emit()

    @property
    def _build_reference(self):
        if len(self.children) == 0:
            address = f'_REF_("{self.cell.address}")'
        else:
            address = self.children[0].emit
            address = address.replace('_R_', '_REF_').replace('_C_', '_REF_')
            if address.startswith('_REF_(str('):
                address = address[10:-2]
        return address

    def func_row(self):
        return f'row({self._build_reference})'

    def func_column(self):
        return f'column({self._build_reference})'

    def func_offset(self):
        to_emit = self.comma_join_emit().split(')', 1)[1]
        return f'offset({self._build_reference}{to_emit})'

    def func_indirect(self):
        to_emit = list(c.emit for c in self.children)
        if len(to_emit) == 1:
            to_emit.append('True')
        to_emit.append(f'"{self.cell.sheet}"')
        return f'indirect({", ".join(to_emit)})'

    SUBTOTAL_FUNCS = {
        1: 'average',
        2: 'count',
        3: 'counta',
        4: 'max_',
        5: 'min_',
        6: 'product',
        7: 'stdev',
        8: 'stdevp',
        9: 'sum_',
        10: 'var',
        11: 'varp',
    }

    def func_subtotal(self):
        # Excel reference: https://support.microsoft.com/en-us/office/
        #   SUBTOTAL-function-7B027003-F060-4ADE-9040-E478765B9939

        # Note: This does not implement skipping hidden rows.

        func_num = coerce_to_number(self.children[0].emit)
        if func_num not in self.SUBTOTAL_FUNCS:
            if func_num - 100 in self.SUBTOTAL_FUNCS:
                func_num -= 100
            else:
                raise ValueError(f"Unknown SUBTOTAL function number: {func_num}")

        func = self.SUBTOTAL_FUNCS[func_num]

        to_emit = self.comma_join_emit(fmt_str="{}", to_emit=self.children[1:])
        return f'{func}({to_emit})'


class ExcelFormula:
    """Take an Excel formula and compile it to Python code."""

    default_modules = (
        'pycel.excellib',
        'pycel.lib.date_time',
        'pycel.lib.engineering',
        'pycel.lib.information',
        'pycel.lib.logical',
        'pycel.lib.lookup',
        'pycel.lib.stats',
        'pycel.lib.text',
        'math',
    )

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
        self.msg = None

    def __str__(self):
        return self.base_formula or self.python_code

    def __repr__(self):
        return f'ExcelFormula({self.base_formula or self.python_code})'

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
                code = iter((self.python_code.encode(),))
                tokens = tuple(tk.tokenize(lambda: next(code)))
                addrs = []
                for i, t in enumerate(tokens):
                    if t.type == 1 and t.string in ADDR_FUNCS_NAMES and (
                            tokens[i + 1].string == '(' and
                            tokens[i + 3].string == ')'):
                        addrs.append(AddressRange(tokens[i + 2].string[1:-1]))
                self._needed_addresses = uniqueify(addrs)
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
                        f"Failed to compile expression {self.python_code}: {exc}")

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
                if next_token.matches(Token.SEP, Token.ARG):
                    tokens.append(token)
                    token = Token('', Token.OPERAND, Token.EMPTY)

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

            elif token.matches(Token.SEP, Token.ARG):
                if next_token.matches(Token.SEP, Token.ARG) or \
                        next_token.matches(Token.FUNC, Token.CLOSE):
                    tokens.append(token)
                    token = Token('', Token.OPERAND, Token.EMPTY)

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
                    raise FormulaParserError("Mismatched or misplaced parentheses")

                were_values.pop()
                arg_count[-1] += 1
                were_values.append(False)

            elif token.is_operator:

                while stack and stack[-1].is_operator and (
                        token.precedence < stack[-1].precedence):
                    output.append(self._ast_node(stack.pop()))

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            elif token.subtype == token.CLOSE:

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(self._ast_node(stack.pop()))

                if not stack:
                    raise FormulaParserError("Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = self._ast_node(stack.pop())
                    f.num_args = arg_count.pop() + int(were_values.pop())
                    output.append(f)

            else:
                assert token.type == token.WSPACE, f'Unexpected token: {token}'

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
                            f"'{node.token.value}' operator missing operand")
                    tree.add_node(arg1, pos=0)
                    tree.add_node(arg2, pos=1)
                    tree.add_edge(arg1, node)
                    tree.add_edge(arg2, node)
                else:
                    try:
                        arg1 = stack.pop()
                    except IndexError:
                        raise FormulaParserError(
                            f"'{node.token.value}' operator missing operand")
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
    def build_eval_context(cls, evaluate, evaluate_range,
                           logger=None, plugins=None):
        """eval with namespace management.  Will auto import needed functions

        Used like:

            build_eval(...)(expression returned from build_python)

        :param evaluate: a function to evaluate a cell address
        :param evaluate_range: a function to evaluate a range address
        :param logger: a logger to use (defaults to pycel)
        :param plugins: module paths for plugin lib functions
        :return: a function to evaluate a compiled expression from build_ast
        """

        if plugins is None:
            modules = ()
        elif isinstance(plugins, str):
            modules = (plugins, )
        else:
            modules = tuple(plugins)
        modules = tuple(importlib.import_module(m)
                        for m in modules + cls.default_modules)

        logger = logger or logging.getLogger('pycel')
        error_messages = []

        def capture_error_state(is_exception, msg):
            if is_exception:
                import traceback
                trace = traceback.format_exc()
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

            # function to fixup the operands
            name_space['excel_operator_operand_fixup'] = \
                build_operator_operand_fixup(capture_error_state)

            # hook for the execed code to save the resulting lambda
            name_space['lambdas'] = lambdas = []

            # get the compiled code and needed names
            compiled, names = excel_formula.compiled_python

            # load the needed names
            not_found = load_functions(names, name_space, modules)

            # exec the code to define the lambda
            exec(compiled, name_space, name_space)
            excel_formula.compiled_lambda = lambdas[0]
            del name_space['lambdas']
            return not_found

        def eval_func(excel_formula, cse_array_address=None):
            """ Call the compiled lambda to evaluate the cell """

            if excel_formula.compiled_lambda is None:
                missing = load_function(excel_formula, locals())
                if missing:
                    msg_fmt = 'Function {} is not implemented. '
                    excel_formula.msg = '\n'.join(
                        msg_fmt.format(f.upper()) +
                        func_status_msg(f)[1] for f in sorted(missing))

            try:
                with in_array_formula_context(cse_array_address):
                    ret_val = in_array_formula_context.fit_to_range(
                        excel_formula.compiled_lambda())

            except NameError:
                error_logger('error', excel_formula.python_code,
                             msg=excel_formula.msg, exc=UnknownFunction)

            except RecursionError as exc:
                raise RecursionError('Do you need to use cycles=True ?') from exc

            except Exception:
                address = f"{excel_formula.cell.address}: " if excel_formula.cell else ""
                error_logger('error', f"{address}{excel_formula.python_code}",
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

        source_code = f"lambdas.append(lambda: {self.python_code})"
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
                if isinstance(node.op, ast.BitAnd) and self.is_addr_and(node):
                    return node
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

            def is_addr_and(self, node):
                # reference intersection does not get fixup
                return (isinstance(node.left, ast.Call) and
                        node.left.func.id == '_REF_' and
                        isinstance(node.right, ast.Call) and
                        node.right.func.id == '_REF_'
                        )

        # modify the ast tree to convert Compare and BinOp to Call
        tree = ast.fix_missing_locations(OperatorWrapper().visit(tree))

        # compile the tree
        self._compiled_python = compile(tree, **kwargs), names
        self._marshalled_python = marshal.dumps(self._compiled_python[0]), names
