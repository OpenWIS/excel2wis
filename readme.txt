##############################
# Metadata generation script #
##############################

# Version of pyhton
The script has been tested with python 2.6.6

# Librairies
4 python libraries must be installed to run the script:
- xlrd 0.9.4
- xmltodict 0.9.2
- xlwt 1.0.0 
- argparse 7.1.2

# Run the script
./excel2xml.py Metadata-guide-record.xls
where Metadata-guide-record.xls is the excel file containing metadata information
The script needs the file template_WMO.xml to run (in the same directory).

# Options
[--openwis tempate_OpenWIS.xml]
The script creates an XML file necessary to insert metadata in OpenWIS

# Note
If you see the error “: No such file or directory” on script execution you need to remove the "Windows line ending"
by running the following command on unix system : 
dos2unix excel2xml.py
