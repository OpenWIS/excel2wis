# -*- coding: utf-8 -*-
from setuptools import setup
from setuptools import find_packages

setup(
    name='excel2wisxml',
    version='3.3',
    include_package_data = True,
    packages=["excel2wisxml"],
    zip_safe = False,
    entry_points={
        'console_scripts': [
            'excel2wisxml = excel2wisxml.excel2wisxml:main']
    }
)
