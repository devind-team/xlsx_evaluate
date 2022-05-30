# Calculate XLSX formulas

[![CI](https://github.com/devind-team/xlsx_evaluate/workflows/Release/badge.svg)](https://github.com/devind-team/devind-django-dictionaries/actions)
[![Coverage Status](https://coveralls.io/repos/github/devind-team/xlsx_evaluate/badge.svg?branch=main)](https://coveralls.io/github/devind-team/devind-django-dictionaries?branch=main)
[![PyPI version](https://badge.fury.io/py/xlsx-evaluate.svg)](https://badge.fury.io/py/xlsx_evaluate)
[![License: MIT](https://img.shields.io/badge/License-MIT-success.svg)](https://opensource.org/licenses/MIT)

**xlsx_evaluate** - python library to convert excel functions in python code without the need for Excel itself within the scope of supported features.

This library is fork [xlcalculator](https://github.com/bradbase/xlcalculator). Use this library.

# Summary

- [Currently supports](docs/support.rst)
- [Supported Functions](docs/support_functions.rst)
- [Adding/Registering Excel Functions](docs/support_functions.rst)
- [Excel number precision](docs/number_precision.rst)
- [Test](docs/test.rst)

# Installation

```shell
# pip
pip install xlsx-evaluate
# poetry
poetry add xlsx-evaluate
```

# TODO

- Do not treat ranges as a granular AST node it instead as an operation ":" of
  two cell references to create the range. That will make implementing
  features like ``A1:OFFSET(...)`` easy to implement.

- Support for alternative range evaluation: by ref (pointer), by expr (lazy
  eval) and current eval mode.

    * Pointers would allow easy implementations of functions like OFFSET().

    * Lazy evals will allow efficient implementation of IF() since execution
      of true and false expressions can be delayed until it is decided which
      expression is needed.

- Implement array functions. It is really not that hard once a proper
  RangeData class has been implemented on which one can easily act with scalar
  functions.

- Improve testing

- Refactor model and evaluator to use pass-by-object-reference for values of
  cells which then get "used"/referenced by ranges, defined names and formulas

- Handle multi-file addresses

- Improve integration with pyopenxl for reading and writing files [example of
  problem space](https://stackoverflow.com/questions/40248564/pre-calculate-excel-formulas-when-exporting-data-with-python)