"""Define criteria operators in Excel."""

import re
from typing import Union

from . import func_xltypes, operator, xlerrors

CRITERIA_REGEX = r'(\W*)(.*)'

CRITERIA_OPERATORS = {
    '<': operator.OP_LT,
    '<=': operator.OP_LE,
    '=': operator.OP_EQ,
    '<>': operator.OP_NE,
    '>=': operator.OP_GE,
    '>': operator.OP_GT,
}


def parse_criteria(criteria: Union[str, func_xltypes.Text, func_xltypes.Array]):
    """Parse criteria."""
    # Not support for arrays right now
    if isinstance(criteria, (str, func_xltypes.Text)):
        search = re.search(CRITERIA_REGEX, str(criteria)).group
        str_operator, str_value = search(1), search(2)
        operator = CRITERIA_OPERATORS.get(str_operator)
        if operator is None:
            operator = CRITERIA_OPERATORS['=']
            str_value = criteria
        value = str_value
        for xl_type in (
            func_xltypes.Number,
            func_xltypes.DateTime,
            func_xltypes.Boolean
        ):
            try:
                value = xl_type.cast(str_value)
            except xlerrors.ValueExcelError:
                pass
            else:
                break

        def check(probe):
            """Check a criteria via operator."""
            return operator(probe, value)

        return check

    criteria = func_xltypes.ExcelType.cast_from_native(criteria)
    if isinstance(criteria, func_xltypes.Array):
        raise xlerrors.ValueExcelError('Array criteria not support.')

    def check(probe):
        """Check a native criteria."""
        return probe == criteria

    return check
