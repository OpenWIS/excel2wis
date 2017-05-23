#!/bin/bash
pip uninstall excel2wisxml
python setup.py sdist
pip install dist/excel2wisxml-3.3.tar.gz --user
