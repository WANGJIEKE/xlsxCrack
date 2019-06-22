# xlsxCrack
A simple program for removing password from `.xlsx` or `.xlsm` file.

## Usage
```
xlsxcrack.py [-h] file_name
```

Notice that I wrote this script in Python 3.7.2 so please use a newer version of Python.

## How It Works
`.xlsx` (or `.xlsm`) files are actually a zipped set of `.xml` files and other resources files. Therefore, just delete the 
`<workbookProtection />` tag inside `xl/workbook.xml` and `<sheetProtection />` tag inside `xl/worksheets/sheet*.xml` 
(the asterisk represents a non-negative integer), the password will disappear.
