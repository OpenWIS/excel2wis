#!/bin/bash
pip uninstall excel2wisxml
rm dist/*
# Create python package in dist directory
python setup.py sdist
pip install dist/excel2wisxml-*.tar.gz --user
