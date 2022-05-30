"""Define errors in Excel."""

# https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/3.0/openpyxl/cell/cell.py #noqa
ERROR_CODE_NULL = '#NULL!'
ERROR_CODE_DIV_ZERO = '#DIV/0!'
ERROR_CODE_VALUE = '#VALUE!'
ERROR_CODE_REF = '#REF!'
ERROR_CODE_NAME = '#NAME?'
ERROR_CODE_NUM = '#NUM!'
ERROR_CODE_NA = '#N/A'

ERROR_CODES = (
    ERROR_CODE_NULL,
    ERROR_CODE_DIV_ZERO,
    ERROR_CODE_VALUE,
    ERROR_CODE_REF,
    ERROR_CODE_NAME,
    ERROR_CODE_NUM,
    ERROR_CODE_NA,
)

ERRORS_BY_CODE = {}


def register(cls):
    """Decorator for register errors."""
    ERRORS_BY_CODE[cls.value] = cls
    return cls


class ExcelError(Exception):
    """Excel exception."""

    def __init__(self, value, info=None):
        super().__init__(info)
        self.value = value
        self.info = info

    @classmethod
    def is_error(cls, value):
        return isinstance(value, cls)

    def __str__(self) -> str:
        return str(self.value)

    def __eq__(self, other) -> bool:
        if isinstance(other, str):
            return str(self) == other
        return id(self) == id(other)


class SpecificExcelError(ExcelError):
    """Specific Excel error."""

    value = None

    def __init__(self, info=None):
        super().__init__(self.value, info)


@register
class NullExcelError(SpecificExcelError):
    """Error NULL Excel code."""

    value = ERROR_CODE_NULL


@register
class DivZeroExcelError(SpecificExcelError):
    """Error DIV_BY_ZERO Excel code."""

    value = ERROR_CODE_DIV_ZERO


@register
class ValueExcelError(SpecificExcelError):
    """Error VALUE Excel code."""

    value = ERROR_CODE_VALUE


@register
class RefExcelError(SpecificExcelError):
    """Error REF Excel code."""

    value = ERROR_CODE_REF


@register
class NameExcelError(SpecificExcelError):
    """Error NAME Excel code."""

    value = ERROR_CODE_NAME


@register
class NumExcelError(SpecificExcelError):
    """Error NUM Excel code."""

    value = ERROR_CODE_NUM


@register
class NaExcelError(SpecificExcelError):
    """Error NA Excel code."""

    value = ERROR_CODE_NA
