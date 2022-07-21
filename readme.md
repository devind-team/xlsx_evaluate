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


# Example

```python
input_dict = {
    'B4': 0.95,
    'B2': 1000,
    "B19": 0.001,
    'B20': 4,
    'B22': 1,
    'B23': 2,
    'B24': 3,
    'B25': '=B2*B4',
    'B26': 5,
    'B27': 6,
    'B28': '=B19 * B20 * B22',
    'C22': '=SUM(B22:B28)',
    "D1": "abc",
    "D2": "bca",
    "D3": "=CONCATENATE(D1, D2)",
  }

from xlsx_evaluate import ModelCompiler
from xlsx_evaluate import Evaluator

compiler = ModelCompiler()
my_model = compiler.read_and_parse_dict(input_dict)
evaluator = Evaluator(my_model)

for formula in my_model.formulae:
    print(f'Formula {formula} evaluates to {evaluator.evaluate(formula)}')

# cells need a sheet and Sheet1 is default.
evaluator.set_cell_value('Sheet1!B22', 100)
print('Formula B28 now evaluates to', evaluator.evaluate('Sheet1!B28'))
print('Formula C22 now evaluates to', evaluator.evaluate('Sheet1!C22'))
print('Formula D3 now evaluates to', evaluator.evaluate("Sheet1!D3"))
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
