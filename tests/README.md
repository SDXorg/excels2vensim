Tests use [pytest](https://docs.pytest.org), [pytest-cov](https://pytest-cov.readthedocs.io), and [pytest-xdist](https://pytest-xdist.readthedocs.io).

To run tests:
```
make tests
```

To have coverage information:
```
make cover
```

To have information of visit lines in HTML files in `htmlcov` folder:
```
make coverhtml
```

To run in multiple CPUs (e.g: for 4 CPUs):
```
make command NUM_PROC=4
```
where `command` is any of above.

