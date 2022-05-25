"""Define criteria operators in Excel."""

from typing import Any, Union
import re

from . import operator, xlerrors, func_xltypes

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
            str_operator = CRITERIA_OPERATORS['=']
            str_value = criteria
        value = str_value
        for XlType in (
            func_xltypes.Number,
            func_xltypes.DateTime,
            func_xltypes.Boolean
        ):
            try:
                value = XlType.cast(str_value)
            except xlerrors.ValueExcelError:
                pass
            else:
                break

        def check(probe):
            """Function for check criteria via operator."""
            return operator(probe, value)

        return check

    criteria = func_xltypes.ExcelType.cast_from_native(criteria)
    if isinstance(criteria, func_xltypes.Array):
        raise xlerrors.ValueExcelError('Array criteria not support.')

    def check(probe: Any):
        """Function for check native criteria."""
        return probe == criteria

    return check
