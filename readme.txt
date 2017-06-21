##############################
# Metadata generation script #
##############################

# Version of pyhton
The script has been tested with python 2.6.6

# Install the script
## Configuration of python user
In ~/.bashrc file add
> PATH=$HOME/local/bin
> export PATH
> PYTHONUSERBASE=$HOME/local
> export PYTHONUSERBASE
## Installation of the script
pip install excel2wisxml.tar.gz --user

# Run the script
excel2wisxml Metadata-guide-record.xls
where Metadata-guide-record.xls is the excel file containing metadata information

# Options
[--MFopenwis]
The script creates a CSV file containing metadata urn and data file name (a row for each metadata).

# Note
If you see the error “: No such file or directory” on script execution you need to remove the "Windows line ending"
by running the following command on unix system : 
dos2unix excel2wisxml.py
