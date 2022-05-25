"""Running model for Excel formulas."""

import copy
import gzip
import jsonpickle
import logging
import os
from dataclasses import dataclass, field

from . import xltypes, reader, parser, tokenizer

@dataclass
class Model:
    ...


class ModelCompiler:
    ...