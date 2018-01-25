# Versioning

The script version and excel file compatible version are referenced at the beginning of the script.

Excel file version is referenced in MD generic sheet. During generic metadata processing, the script checks if versions match. Otherwise it stops.

# Functions module

Imported module _excel2wisxmltils.py_ contains functions used by the script to add metadata elements in an xml tree.

# Excel file shape configuration

Shape of the excel file is set for the 4 sheets of the excel file :
- delta linking _MD Fields_ columns to associated _Help_ rows
- _MD Fields_ :
    - column and row start
    - mandatory row
    - section row
- _Help_ : associate columns with their header names
- _MD generic_ :
    - row start
    - associate columns with their names
- _MD Thesaurus_ :
    - column start
    - associate columns with their names

Be careful always to keep row numbers, column numbers, header name values and delta linking Help rows to MD Fields columns consistent with excel file.

The addition of a column in _MD Field_ or a row in _Help_ does not impact shape configuration.

# Read and parse excel file with xlrd
Get sheets and check excel file shape :
- Do _MD Fields_ ID match with _Help_ ID ?
- Are mandatory fields filled ?

If an error is identified the script stops and an appropriate error message is displayed.

# Parse the template as an XML tree with etree

# Add metadata elements

An element is added only when its value is not null.

## 1. Generic metadata (_MD generic_ sheet)

Each row of _MD generic_ sheet stands for a metadata element. The element's value is added at the location specified by the xpath. A codelist, attributes and a translation can also be added.

If there is no metadata date value, the computer time stamp is added by default.    

If there is a field value but no associated xpath, a warn message is displayed, except for **ExcelVersion** element. This element is not added in the metadata file. It is used to check if the excel file version is compatible with the script version. If not, the script stops. 

For some elements - identified thanks to their tag name - some other tags are added :
- **Resource locator url**  
addition of online resource protocol _WWW:LINK-1.0-http--link_

### DCPC Metadata

Metadata are identified as DCPC metadata when one of _OpenWIS only:_ tag is filled with a value. For such metadata some additionnal elements (request and subscribe URL) are added during specific metadata processing.

### Translations

A metadata is identified as translated when _Second Language_ is filled with a value. For such metadata some further elements are added (translation language, value and encoding).

## 2. Non generic metadata (_MD Fields_ and _Help_ sheets)

An XML metadata file is generated for each row of _MD Field_ sheet.

For each row :
- _MD Fields_ columns are parsed one by one to get the values ;
- linked _Help_ rows are parsed to get :
    - xpath (location of metadata element in xml tree)
    - attribute name, value and location (attribute name is set to _No_ if there is no attribute to add)
    - thesaurus name
    - multivalue boolean (indicates if the tag must be added only once or several times)
    - codelist value
    - type

Regular case is to add the metadata tag value at the location specified by the xpath.

Other elements can be added according to which _Help_ cells are filled :
- codeList and codeListValue attributes (added at the location specified by the xpath, except for Keywords)
- other comma-separated attributes (added at the location specified in _Attribute Location_ or at xpath location if _Attribute Location_ is empty) ;
- thesaurus name, date, dateType and codeList attributes (information available in _MD Thesaurus_ sheet) ;
- type tag with its associated value and codelist attributes (_KeywordType_ and _DateType_ for instance).

Some values are kept in memory such as :
- URN ;
- resource Title ;
- file name pattern.

Some elements need a specific processing :

- **URN** [identified with its xpath]  
    - _MD Field_ Unique identifier value is replaced by the URN - the concatenation of _MD generic_ UID and _MD Fields_ UID  
    - 2 _MD generic_ tags are replaced by the concatenation of their value and the URN : 
        - location (address) for on-line access
        - permanent link
    - If one of _MD generic_ field _OpenWIS only:_ is not null, two online resources are added (_Suscribe on DCPC_ and _Request on DCPC_)

- **GTS Priority** [identified with its xpath]  
Value is replaced by _"GTSPriorityN"_ with _N_ the 9th character of _MD Field_ cell value

- **File name** [identified with its xpath]  
One or several GFNC file name(s) and associated information are added :
    - file description (dynamic value : Resource title)
    
- **Temporal Extent** [identified with its xpath]
For indeterminate temporal extent value, the value is not added as a tag value but as an attribute value. If the value is "before" or "after" a potential additionnal time element can be added as the tag value.

- **Free links** [identified with its xpath]  
Several links can be added. 3 tags are added for each link :
    - name (optional, URL value if not specified)
    - protocol (static value)
    - URL (dynamic value)

- **Resource Format** [identified with its xpath]  
Several resource formats can be added. 4 tags are added for each resource format (dynamic values) :
    - name
    - version
    - specification (optional)
    - mime type  
If file names are specified, format name, version and mime type information is also added in the GFNC section for each file name.

- **Tag or group of tags added several times** [Multivalue set to Yes]  
Add a group of tags as many times as the number of comma-separated values. First multivalued tag to add is suffixed by _"[]"_.

- **Attribute addition exception** [Attribute Value set to MD_Fields]
The value read in _MD Fields_ is added as an attribute value at xpath location and not as a tag value.

### Translations

When a metadata is identified as translated some further elements are added (translation language and value) for each element translated in _MD Fields Translate_ sheet. For multivalue tags, these further elements are added for each occurence.

# Empty descriptiveKeywords tags removal
There are several descriptive keywords tags in the template. But not all of them are mandatory. Empty descriptiveKeywords are removed top down.

# Warn messages
Warn message can be displayed for each metadata in case where :
- there is a field value but no xpath ;
- a xpath is incorrect.
