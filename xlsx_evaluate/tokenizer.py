"""Tokenise an Excel formula using an implementation of E. W. Bachtal's algorithm.

========================================================================
    found here:

                  http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html

              Tested with Python v2.5 (win32)
        Author: Robin Macharg
        Copyright: Algorithm (c) E. W. Bachtal, this implementation (c) R. Macharg

    Modification History

    Date         Author Comment
    =======================================================================
    2006/11/29 - RMM  - Made strictly class-based.
                    Added parse, render and pretty print methods
    2006/11    - RMM  - RMM = Robin Macharg
                          Created
    2011/10    - Dirk Gorissen - Patch to support scientific notation
========================================================================
"""

import re
from dataclasses import dataclass, field
import uuid

from .utils import col2num, init_uuid, num2col


class ExcelParserTokens:
    """Inheritable container for token definitions.

    ========================================================================
    Class: ExcelParserTokens
        Description: Inheritable container for token definitions

        Attributes: Self explanatory

        Methods: None
    ========================================================================
    """

    TOK_TYPE_NOOP = 'noop'
    TOK_TYPE_OPERAND = 'operand'
    TOK_TYPE_FUNCTION = 'function'
    TOK_TYPE_SUBEXPR = 'subexpression'
    TOK_TYPE_ARGUMENT = 'argument'
    TOK_TYPE_OP_PRE = 'operator-prefix'
    TOK_TYPE_OP_IN = 'operator-infix'
    TOK_TYPE_OP_POST = 'operator-postfix'
    TOK_TYPE_WSPACE = 'white-space'
    TOK_TYPE_UNKNOWN = 'unknown'

    TOK_SUBTYPE_START = 'start'
    TOK_SUBTYPE_STOP = 'stop'
    TOK_SUBTYPE_TEXT = 'text'
    TOK_SUBTYPE_NUMBER = 'number'
    TOK_SUBTYPE_LOGICAL = 'logical'
    TOK_SUBTYPE_ERROR = 'error'
    TOK_SUBTYPE_RANGE = 'range'
    TOK_SUBTYPE_MATH = 'math'
    TOK_SUBTYPE_CONCAT = 'concatenate'
    TOK_SUBTYPE_INTERSECT = 'intersect'
    TOK_SUBTYPE_UNION = 'union'
    TOK_SUBTYPE_NONE = 'none'


@dataclass
class f_token:  # noqa
    """Class f_token.

    ========================================================================
    Class: f_token

    Attributes:
        tvalue - See token definitions, above, for values
        ttype - See token definitions, above, for values
        tsubtype - See token definitions, above, for values

    Methods: f_token  - __init__()
    ========================================================================
    """

    tvalue: str
    ttype: str
    tsubtype: str
    unique_identifier: uuid = field(init=False, default_factory=init_uuid, compare=True, hash=True, repr=True)

    def __repr__(self):
        return f'<{self.__class__.__name__} tvalue: {self.tvalue} ttype: {self.ttype} tsubtype: {self.tsubtype} >'

    def __str__(self):
        return self.__repr__()


class f_tokens:  # noqa
    """An ordered list of tokens.

    ========================================================================
    Class: f_tokens
    Attributes:
        items - Ordered list
        index - Current position in the list

    Methods: f_tokens     - __init__()
        f_token      - add()      - Add a token to the end of the list
        None         - addRef()   - Add a token to the end of the list
        None         - reset()    - reset the index to -1
        Boolean      - BOF()      - End of list?
        Boolean      - EOF()      - Beginning of list?
        Boolean      - moveNext() - Move the index along one
        f_token/None - current()  - Return the current token
        f_token/None - next()     - Return the next token (leave the
                                 index unchanged)
        f_token/None - previous() - Return the previous token (leave
                                 the index unchanged)
    ========================================================================
    """

    def __init__(self):
        self.items = []
        self.index = -1

    def add(self, value, ttype, subtype=''):
        if not subtype:
            subtype = ''
        token = f_token(value, ttype, subtype)
        self.addRef(token)
        return token

    def addRef(self, token):
        self.items.append(token)

    def reset(self):
        self.index = -1

    def BOF(self):
        return self.index <= 0

    def EOF(self):
        return self.index >= len(self.items) - 1

    def moveNext(self):
        if self.EOF():
            return False
        self.index += 1
        return True

    def current(self):
        if self.index == -1:
            return None
        return self.items[self.index]

    def __next__(self):
        if self.EOF():
            return None
        return self.items[self.index + 1]

    def previous(self):
        if self.index < 1:
            return None
        return self.items[self.index - 1]

    def __iter__(self):
        """Make this object pass as an iterator."""
        self.reset()
        return self

    def next(self):
        return self.__next__()


class f_tokenStack(ExcelParserTokens):  # noqa
    """A LIFO stack of tokens.

    ========================================================================
    Class: f_tokenStack
    Inherits: ExcelParserTokens - a list of token values
    Description: A LIFO stack of tokens

    Attributes:
        items - Ordered list

    Methods: f_tokenStack - __init__()
        None         - push(token) - Push a token onto the stack
        f_token/None - pop()       - Pop a token off the stack
        f_token/None - token()     - Non-destructively return the top
                                   item on the stack
        String       - type()      - Return the top token's type
        String       - subtype()   - Return the top token's subtype
        String       - value()     - Return the top token's value
    ========================================================================
    """

    def __init__(self):
        self.items = []

    def push(self, token):
        self.items.append(token)

    def pop(self):
        token = self.items.pop()
        return f_token('', token.ttype, self.TOK_SUBTYPE_STOP)

    def token(self):
        """Note: this uses Pythons and/or "hack" to emulate C's ternary operator (i.e. cond ? exp1 : exp2)."""
        return (
            (
                (len(self.items) > 0)
                and [self.items[len(self.items) - 1]]
                or [None]
            )[0]
        )

    def value(self):
        return (self.token() and [self.token().tvalue] or [''])[0]

    def type(self):
        return (self.token() and [self.token().ttype] or [''])[0]

    def subtype(self):
        return (self.token() and [self.token().tsubtype] or [''])[0]


class ExcelParser(ExcelParserTokens):
    """Parse an excel formula into a stream of tokens. # noqa

    ========================================================================
    Class: ExcelParser
    Description:

    Attributes:

    Methods: f_tokens - getTokens(formula) - return a token stream (list)
    ========================================================================
    """

    def __init__(self, tokenize_range: bool = False):
        self.tokens = None
        self.OPERATORS = '+-*/^&=><:' if tokenize_range else '+-*/^&=><'

    def getTokens(self, formula: str):
        """Build tokens from stream."""

        def currentChar():
            return formula[offset]

        def doubleChar():
            return formula[offset:offset + 2]

        def nextChar():
            """Javascript returns an empty string if the index is out of bounds.

            Python throws an IndexError. We mimic this behaviour here.
            """
            try:
                formula[offset + 1]
            except IndexError:
                return ''
            else:
                return formula[offset + 1]

        def EOF():
            return offset >= len(formula)

        tokens = f_tokens()
        token_stack = f_tokenStack()
        offset = 0
        token = ''
        in_string = False
        in_path = False
        in_range = False
        in_error = False

        formula = formula.lstrip('\n =')

        # state-dependent character evaluation (order is important)
        while not EOF():
            # double-quoted strings
            # embeds are doubled
            # end marks token
            if in_string:
                if currentChar() == '"':
                    if nextChar() == '"':
                        token += '"'
                        offset += 1
                    else:
                        in_string = False
                        tokens.add(
                            token, self.TOK_TYPE_OPERAND,
                            self.TOK_SUBTYPE_TEXT)
                        token = ''

                else:
                    token += currentChar()
                offset += 1
                continue

            # single-quoted strings (links)
            # embeds are double
            # end does not mark a token
            if in_path:
                if currentChar() == "'":
                    if nextChar() == "'":
                        token += "'"
                        offset += 1
                    else:
                        in_path = False
                else:
                    token += currentChar()
                offset += 1
                continue

            # bracketed strings (range offset or linked workbook name)
            # no embeds (changed to "()" by Excel)
            # end does not mark a token
            if in_range:
                if currentChar() == ']':
                    in_range = False
                token += currentChar()
                offset += 1
                continue

            # error values
            # end marks a token, determined from absolute list of values
            if in_error:
                token += currentChar()
                offset += 1
                if ',#NULL!,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#N/A,'.find(f',{token},') != -1:
                    in_error = False
                    tokens.add(
                        token, self.TOK_TYPE_OPERAND, self.TOK_SUBTYPE_ERROR)
                    token = ''
                continue

            # scientific notation check
            regex_sn = r'^[1-9]{1}(\.[0-9]+)?[eE]{1}$'
            if '+-'.find(currentChar()) != -1 and len(token) > 1 and re.match(regex_sn, token):
                token += currentChar()
                offset += 1
                continue

            # independent character evaluation (order not important)
            # establish state-dependent character evaluations
            if currentChar() == '"':
                if len(token) > 0:
                    # not expected
                    tokens.add(token, self.TOK_TYPE_UNKNOWN)
                    token = ''
                in_string = True
                offset += 1
                continue

            if currentChar() == "'":
                if len(token) > 0:
                    # not expected
                    tokens.add(token, self.TOK_TYPE_UNKNOWN)
                    token = ''
                in_path = True
                offset += 1
                continue

            if currentChar() == '[':
                in_range = True
                token += currentChar()
                offset += 1
                continue

            if currentChar() == '#':
                if len(token) > 0:
                    # not expected
                    tokens.add(token, self.TOK_TYPE_UNKNOWN)
                    token = ''
                in_error = True
                token += currentChar()
                offset += 1
                continue

            # mark start and end of arrays and array rows
            if currentChar() == '{':
                if len(token) > 0:
                    # not expected
                    tokens.add(token, self.TOK_TYPE_UNKNOWN)
                    token = ''
                token_stack.push(tokens.add(
                    'ARRAY',
                    self.TOK_TYPE_FUNCTION, self.TOK_SUBTYPE_START)
                )
                token_stack.push(tokens.add(
                    'ARRAYROW',
                    self.TOK_TYPE_FUNCTION, self.TOK_SUBTYPE_START)
                )
                offset += 1
                continue

            if currentChar() == ';':
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                tokens.addRef(token_stack.pop())
                tokens.add(',', self.TOK_TYPE_ARGUMENT)
                token_stack.push(tokens.add(
                    'ARRAYROW',
                    self.TOK_TYPE_FUNCTION, self.TOK_SUBTYPE_START)
                )
                offset += 1
                continue

            if currentChar() == '}':
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                tokens.addRef(token_stack.pop())
                tokens.addRef(token_stack.pop())
                offset += 1
                continue

            # trim white-space
            if currentChar() in (' ', '\n'):
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                tokens.add('', self.TOK_TYPE_WSPACE)
                offset += 1
                while (currentChar() in (' ', '\n')) and (not EOF()):
                    offset += 1
                continue

            # multi-character comparators
            if ',>=,<=,<>,'.find(',' + doubleChar() + ',') != -1:
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                tokens.add(
                    doubleChar(),
                    self.TOK_TYPE_OP_IN, self.TOK_SUBTYPE_LOGICAL)
                offset += 2
                continue

            # standard infix operators
            if self.OPERATORS.find(currentChar()) != -1:
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                tokens.add(currentChar(), self.TOK_TYPE_OP_IN)
                offset += 1
                continue

            # standard postfix operators
            if '%'.find(currentChar()) != -1:
                if len(token) > 0:
                    tokens.add(float(token) / 100, self.TOK_TYPE_OPERAND)
                    token = ''
                else:
                    tokens.add('*', self.TOK_TYPE_OP_IN)
                    tokens.add(0.01, self.TOK_TYPE_OPERAND)
                # tokens.add(currentChar(), self.TOK_TYPE_OP_POST) # noqa
                offset += 1
                continue

            # start subexpression or function
            if currentChar() == '(':
                if len(token) > 0:
                    token_stack.push(tokens.add(
                        token, self.TOK_TYPE_FUNCTION, self.TOK_SUBTYPE_START)
                    )
                    token = ''
                else:
                    token_stack.push(tokens.add(
                        '', self.TOK_TYPE_SUBEXPR, self.TOK_SUBTYPE_START)
                    )
                offset += 1
                continue

            # function, subexpression, array parameters
            if currentChar() == ',':
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                if token_stack.type() != self.TOK_TYPE_FUNCTION:
                    tokens.add(
                        currentChar(),
                        self.TOK_TYPE_OP_IN, self.TOK_SUBTYPE_UNION)
                else:
                    tokens.add(currentChar(), self.TOK_TYPE_ARGUMENT)
                offset += 1
                if currentChar() == ',':
                    tokens.add(
                        'None',
                        self.TOK_TYPE_OPERAND, self.TOK_SUBTYPE_NONE)
                    token = ''
                continue

            # stop subexpression
            if currentChar() == ')':
                if len(token) > 0:
                    tokens.add(token, self.TOK_TYPE_OPERAND)
                    token = ''
                tokens.addRef(token_stack.pop())
                offset += 1
                continue

            # token accumulation
            token += currentChar()
            offset += 1

        # dump remaining accumulation
        if len(token) > 0:
            tokens.add(token, self.TOK_TYPE_OPERAND)

        # move all tokens to a new collection, excluding all unnecessary
        # white-space tokens
        tokens2 = f_tokens()

        while tokens.moveNext():
            token = tokens.current()

            if token.ttype == self.TOK_TYPE_WSPACE:
                if (
                    tokens.BOF()
                    or tokens.EOF()
                    or not (
                        (tokens.previous().ttype == self.TOK_TYPE_FUNCTION
                         and tokens.previous().tsubtype == self.TOK_SUBTYPE_STOP)
                        or (tokens.previous().ttype == self.TOK_TYPE_SUBEXPR
                            and tokens.previous().tsubtype == self.TOK_SUBTYPE_STOP)
                        or tokens.previous().ttype == self.TOK_TYPE_OPERAND
                    )
                    or not (
                        (tokens.next().ttype == self.TOK_TYPE_FUNCTION
                         and tokens.next().tsubtype == self.TOK_SUBTYPE_START)
                        or (tokens.next().ttype == self.TOK_TYPE_SUBEXPR
                            and tokens.next().tsubtype == self.TOK_SUBTYPE_START)
                        or tokens.next().ttype == self.TOK_TYPE_OPERAND
                    )
                ):
                    pass
                else:
                    tokens2.add(token.tvalue, self.TOK_TYPE_OP_IN, self.TOK_SUBTYPE_INTERSECT)
                continue

            tokens2.addRef(token)

        # switch infix "-" operator to prefix when appropriate, switch infix
        # "+" operator to noop when appropriate, identify operand and
        # infix-operator subtypes, pull "@" from in front of function names
        while tokens2.moveNext():
            token = tokens2.current()
            if token.ttype == self.TOK_TYPE_OP_IN and token.tvalue in ('-', '+'):
                token_sign = {
                    '-': self.TOK_TYPE_OP_PRE,
                    '+': self.TOK_TYPE_NOOP
                }
                if tokens2.BOF():
                    token.ttype = token_sign[token.tvalue]
                elif (
                        tokens2.previous().ttype == self.TOK_TYPE_FUNCTION
                        and tokens2.previous().tsubtype == self.TOK_SUBTYPE_STOP
                        or tokens2.previous().ttype == self.TOK_TYPE_SUBEXPR
                        and tokens2.previous().tsubtype == self.TOK_SUBTYPE_STOP
                        or tokens2.previous().ttype == self.TOK_TYPE_OP_POST
                        or tokens2.previous().ttype == self.TOK_TYPE_OPERAND
                ):
                    token.tsubtype = self.TOK_SUBTYPE_MATH
                else:
                    token.ttype = token_sign[token.tvalue]
                continue

            if token.ttype == self.TOK_TYPE_OP_IN and len(token.tsubtype) == 0:
                if '<>='.find(token.tvalue[0:1]) != -1:
                    token.tsubtype = self.TOK_SUBTYPE_LOGICAL
                elif token.tvalue == '&':
                    token.tsubtype = self.TOK_SUBTYPE_CONCAT
                else:
                    token.tsubtype = self.TOK_SUBTYPE_MATH
                continue

            if token.ttype == self.TOK_TYPE_OPERAND and len(token.tsubtype) == 0:
                try:
                    float(token.tvalue)
                except ValueError:
                    token.tsubtype = self.TOK_SUBTYPE_LOGICAL \
                        if token.tvalue in ('TRUE', 'FALSE') \
                        else self.TOK_SUBTYPE_RANGE
                else:
                    token.tsubtype = self.TOK_SUBTYPE_NUMBER
                continue

            if token.ttype == self.TOK_TYPE_FUNCTION:
                if token.tvalue[0:1] == '@':
                    token.tvalue = token.tvalue[1:]
                continue
        tokens2.reset()
        # move all tokens to a new collection, excluding all noops
        tokens = f_tokens()
        while tokens2.moveNext():
            if tokens2.current().ttype != self.TOK_TYPE_NOOP:
                tokens.addRef(tokens2.current())

        tokens.reset()
        return tokens

    def parse(self, formula):
        """Parse Excel formula."""
        self.tokens = self.getTokens(formula)
        return self.tokens
