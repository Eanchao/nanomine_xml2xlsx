# nanomine_xml2xlsx
A reverse code of xlsx2xml that generates Excel data spreadsheets from nanomine xml data.

By Bingyin Hu

### 1. System preparations

Required packages:
import xlsxwriter
from lxml import etree
import os
import collections
- os
  - Python default package

- collections
  - Python default package

- lxml
  - Used to parse xml files.
  - http://lxml.de/

- xlsxwriter
  - Used to write xlsx files.
  - https://xlsxwriter.readthedocs.io/
    
Open the command or terminal and run
```
pip install -r requirements.txt
```
### 2. How to run

In python, create the xml2xlsx object
```
from xml2xlsx import xml2xlsx

xmlDir = '''{The path of your xml file}'''
test = xml2xlsx(xmlDir)
```

Then run the conversion by
```
test.nm_run() # Nanomine only
```
or
```
test.run() # General case (not pretty)
```

Save the workbook by
```
test.save()
```