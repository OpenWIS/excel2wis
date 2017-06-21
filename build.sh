#!/bin/bash
pip uninstall excel2wisxml
rm dist/*
python setup.py sdist
pip install dist/excel2wisxml-*.tar.gz --user
