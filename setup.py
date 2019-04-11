# -*- coding: utf-8 -*-
from setuptools import setup
from setuptools import find_packages

with open("README.md", "r") as fh:
        long_description = fh.read()

setup(
    name='excel2wisxml',
    version='3.5.4',
    author_email="gisc_support@meteo.fr",
    description="Generation of WMO Core 1.3 profile metadata",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="http://openwis.github.io/openwis-documentation/projects/excel2wis/",
    include_package_data=True,
    packages=["excel2wisxml"],
    install_requires=[
        'xlrd',
        'lxml',
        'argparse',
    ],
    zip_safe=False,
    entry_points={
        'console_scripts': [
            'excel2wisxml = excel2wisxml.excel2wisxml:main',
            'createExcel2wisxml = excel2wisxml.excel2wisxml:createExcel']
    }
)
