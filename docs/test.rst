Run tests
---------

Setup your environment:

```shell
virtualenv -p 3.7 ve
ve/bin/pip install -e .[test]
```

From the root xlcalculator directory
```shell
ve/bin/py.test -rw -s --tb=native
```

Or simply use ``tox``
```
tox
```