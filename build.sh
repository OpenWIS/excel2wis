#!/bin/bash
pip2 uninstall excel2wisxml
rm dist/*
<<<<<<< HEAD
python2 setup.py sdist
pip2 install dist/excel2wisxml-*.tar.gz --user
=======
# Create python package in dist directory
pip install setuptools
python setup.py sdist
pip install dist/excel2wisxml-*.tar.gz --user
>>>>>>> c3a4af4dd98696d0dd6aa2f10a6aaed2947fe0f2
