[![CodeQL](https://github.com/sorzkode/eksel/actions/workflows/codeql.yml/badge.svg)](https://github.com/sorzkode/eksel/actions/workflows/codeql.yml)
[[MIT Licence](https://en.wikipedia.org/wiki/MIT_License)]


![alt text](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/ekselgit.png)

# Eksel Splitter

Quickly copy and save worksheets from Excel to a folder of your choosing.

## Example

![alt text](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/example.png)

## Installation

Download from Github, changedir (cd) to the script directory and run the following:
```
pip install -e .
```
*This will install the Eksel Splitter package locally 

Installation isn't required to run the script but you will need to ensure the requirements below are met.

## Requirements

The install above should take care of requirments.

Alternatively you can run: pip install -r requirements.txt

  [[Python 3](https://www.python.org/downloads/)]

  [[PySimpleGUI module](https://pypi.org/project/PySimpleGUI/)]

  [[xlwings module](https://pypi.org/project/xlwings/)]

  [[tkinter](https://docs.python.org/3/library/tkinter.html)] :: Linux Users

## Usage

If Eksel Splitter is installed you can use the following command syntax:
```
python -m eksel
```
Otherwise you can run the script directly by changing directory (cd) in a terminal of your choice to the Eksel Splitter directory and using the following syntax:
```
python eksel.py
```
Once the script is initiated: 
```
  1. Click the "Select File" button.
  2. Select your Excel file and click "ok".
  3. Manipulate worksheets (if desired).
  4. Save worksheets using the corresponding button.
```

Things to note:
```
  * SHEETBOX #1 will populate a list of all worksheet names when a file is selected.
  * If you want to copy and save all worksheets, click the "Save Box #1" button. Otherwise, click any worksheet names you want to move to SHEETBOX #2.
  * If names are moved to SHEETBOX #2, the "Save Box #2" button will be enabled.
  * Determine which listbox you want to save and use the corresponding button.
  * Names can be switched between the listboxes by clicking on them.
  * Use the "Clear All" button if you made a mistake or need to start over.
  * Acceptable filetypes: *.xlsx, *.xlsm, and *.xls.
  * Worksheets are saved as *.xlsx by default.
  * Any conflicting filenames will be overwritten by default. 
```





