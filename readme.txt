##############################
# Metadata generation script #
##############################

# Version of pyhton
The script has been tested with python 2.6.6

# Librairies
2 python libraries must be installed to run the script:
- xlrd 0.9.4
- xmltodict 0.9.2

# Run the script
./excel2xml.py Metadata-guide-record.xls
where Metadata-guide-record.xls is the excel file containing metadata information
The script needs the file template_WMO.xml to run (in the same directory).

# Options
[-openwis]
The script creates a file (suffixed _datalink.csv) containing urn - filename link
