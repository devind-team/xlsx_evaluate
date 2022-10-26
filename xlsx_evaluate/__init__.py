"""Library for evaluate xlsx formulas."""

from .model import ModelCompiler, Model  # noqa
from .evaluator import Evaluator  # noqa


from .functions.xl import FUNCTIONS, register  # noqa: F401
from .functions.xlerrors import *  # noqa: F401, F403
from .functions.func_xltypes import *  # noqa: F401, F403

# Make sure to register all functions
from .functions import (  # noqa: F401
    date,
    financial,
    information,
    logical,
    lookup,
    math,
    operator,
    statistics,
    text
)

__version__ = '0.4.7'
