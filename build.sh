#!/bin/bash
pip2 uninstall excel2wisxml
rm dist/*
python2 setup.py sdist
pip2 install dist/excel2wisxml-*.tar.gz --user
