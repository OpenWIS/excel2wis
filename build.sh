#!/bin/bash
# Uninstall excel2wisxml python package
pip2 uninstall excel2wisxml
rm dist/*
# Install setuptools
pip2 install setuptools
# Create excel2wisxml python package in dist directory
python2 setup.py sdist
# Install excel2wisxml package
pip2 install dist/excel2wisxml-*.tar.gz --user
