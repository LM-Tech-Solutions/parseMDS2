# parseMDS2
Convert MDS2 HN-2019 format to XML


## Features
* Accepts .xls and .xlsx as input


## Requirements
Python 3.8 or higher

Python libraries:
* openpyxl


## User Guide
```sh
$ python parseMDS2.py example-mds2-worksheet.xlsx
```

```sh
$ python parseMDS2.py -h
usage: parseMDS2.py [-h] *.xlsx

Parse MDS2 HN-2019 XLSX into XML

positional arguments:
  *.xlsx      MDS2 form in .xls or .xlsx format

optional arguments:
  -h, --help  show this help message and exit
```


## Screenshots
![XML Output](img/screenshot_xml_output.png | width=100)

![XML_Sections](img/screenshot_xml_sections.png | width=100)


## Test Environments
Ubuntu 20.04 LTS, 64-bit

Windows 10 20H2, 64-bit


## Sources
Example-mds2-worksheet [XLSX]. (2019). National Electrical Manufacturers Association (NEMA). ANSI/NEMA HN 1-2019. https://www.nema.org/Standards/view/Manufacturer-Disclosure-Statement-for-Medical-Device-Security.
