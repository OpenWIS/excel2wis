#!/bin/bash
# Uninstall excel2wisxml python package
pip uninstall excel2wisxml
rm dist/*
# Install setuptools
pip install setuptools
# Create excel2wisxml python package in dist directory
python setup.py sdist
# Install excel2wisxml package
pip install dist/excel2wisxml-*.tar.gz --user
