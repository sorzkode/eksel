[![CodeQL](https://github.com/sorzkode/eksel/actions/workflows/codeql.yml/badge.svg)](https://github.com/sorzkode/eksel/actions/workflows/codeql.yml)
[[MIT Licence](https://en.wikipedia.org/wiki/MIT_License)]


![alt text](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/ekselgit.png)

# Eksel Splitter

Eksel Splitter is a Python script that simplifies the process of copying and saving worksheets from Excel to a folder of your choice. With Eksel Splitter, you can easily manipulate and organize your Excel data by selecting specific worksheets and saving them separately. This script provides a user-friendly interface and offers features such as selecting multiple worksheets, saving in different file formats, and handling conflicting filenames. Whether you need to split large Excel files or extract specific data, Eksel Splitter is a convenient tool to streamline your workflow.

## Example

![alt text](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/example.png)

## Installation

To install the Eksel Splitter script, follow these steps:

1. Download the Eksel Splitter script from the [GitHub repository](https://github.com/sorzkode/eksel).

2. Open a terminal or command prompt and navigate to the directory where you downloaded the script.

3. Run the following command to install the Eksel Splitter package locally:
  ```
  pip install -e .
  ```
  This will install the Eksel Splitter package and its dependencies.

  Note: Installation isn't required to run the script, but it's recommended to ensure the requirements are met.

## Requirements

The installation command above should take care of the requirements automatically. 

However, if you need to install them manually, you can run:
```
pip install -requirements.txt
```

  [[Python 3](https://www.python.org/downloads/)]

  [[PySimpleGUI module](https://pypi.org/project/PySimpleGUI/)]

  [[xlwings module](https://pypi.org/project/xlwings/)]

  [[tkinter](https://docs.python.org/3/library/tkinter.html)] :: Linux Users

## Usage

To use Eksel Splitter, follow these steps:

1. If you have installed Eksel Splitter, open a terminal or command prompt and run the following command:
  ```
  python -m eksel
  ```
  If you haven't installed the package, navigate to the Eksel Splitter directory in the terminal using the `cd` command, and then run:
  ```
  python eksel.py
  ```

2. Once the script is initiated, click the "Select File" button.

3. Select your Excel file and click "OK".

4. If desired, manipulate the worksheets by moving them between SHEETBOX #1 and SHEETBOX #2. 
  - SHEETBOX #1 displays all worksheet names when a file is selected.
  - To copy and save all worksheets, click the "Save Box #1" button.
  - To select specific worksheets, click on their names to move them to SHEETBOX #2.
  - The "Save Box #2" button will be enabled if names are moved to SHEETBOX #2.

5. To save the worksheets, use the corresponding button based on the listbox you want to save.

6. If you need to switch names between the listboxes, simply click on them.

7. If you made a mistake or want to start over, use the "Clear All" button.

8. Eksel Splitter supports the following filetypes: *.xlsx, *.xlsm, and *.xls. Worksheets are saved as *.xlsx by default.

9. By default, any conflicting filenames will be overwritten. 

Note: Make sure you have the necessary requirements installed as mentioned in the "Requirements" section of this README file.






