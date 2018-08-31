# excel2wisxml

## Version of pyhton
The package has been tested with python 2.6.6

## Installation
### Configuration of python user
In ~/.bashrc file add
> PATH=$HOME/local/bin
> export PATH
> PYTHONUSERBASE=$HOME/local
> export PYTHONUSERBASE
### Installation of python package
pip install excel2wisxml.tar.gz --user

## Use

### Generate Metadata-guide-record.xls excel template in your current repertory
createExcel2wisxml

### Generate the WMO Core 1.3 metadata from the excel document
excel2wisxml Metadata-guide-record.xls
The generation xml files will be stored in the same directory, with filename format:
MD\_uniqueid\_YYYYMMDDHHMMSS.xml

### Options
[--MFopenwis]
The script creates a CSV file containing metadata urn and data file name (a row for each metadata).
