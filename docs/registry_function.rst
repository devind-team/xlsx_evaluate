Adding/Registering Excel Functions
----------------------------------

Excel function support can be easily added.

Fundamental function support is found in the xlfunctions directory. The
functions are thematically organised in modules.

Excel functions can be added by any code using the
``functions.xl.register()`` decorator. Here is a simple example:

```python
from xlsx_evaluate.functions import xl

@xl.register()
@xl.validate_args
def ADDONE(num: xl.Number):
  return num + 1
```

The `@xl.validate_args` decorator will ensure that the annotated arguments are
converted and validated. For example, even if you pass in a string, it is
converted to a number (in typical Excel fashion):

```
>>> ADDONE(1):
2
>>> ADDONE('1'):
2
```

If you would like to contribute functions, please create a pull request. All
new functions should be accompanied by sufficient tests to cover the
functionality. Tests need to be written for both the Python implementation of
the function (tests/xlfunctions) and a comparison with Excel
(tests/xlfunctions_vs_excel).

