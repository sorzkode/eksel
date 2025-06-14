[![CodeQL](https://github.com/sorzkode/eksel/actions/workflows/codeql.yml/badge.svg)](https://github.com/sorzkode/eksel/actions/workflows/codeql.yml)
[[MIT Licence](https://en.wikipedia.org/wiki/MIT_License)]

![alt text](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/ekselgit.png)

# Eksel Splitter

Eksel Splitter is a Python application that simplifies the process of copying and saving worksheets from Excel to a folder of your choice. With Eksel Splitter, you can easily manipulate and organize your Excel data by selecting specific worksheets and saving them separately. This application provides a user-friendly graphical interface and offers features such as selecting multiple worksheets, saving in different file formats, and handling conflicting filenames. Whether you need to split large Excel files or extract specific data, Eksel Splitter is a convenient tool to streamline your workflow.

## Features

- **User-friendly GUI** built with tkinter
- **Drag-and-drop style interface** for moving worksheets between boxes
- **Error message copying** - All error dialogs include a "Copy Error" button for easy error reporting
- **Batch processing** - Save multiple worksheets at once
- **Smart file naming** - Automatically sanitizes worksheet names for safe file saving
- **Conflict resolution** - Prompts before overwriting existing files

## Example

![Screenshot](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/example.png)

## Demo Video

https://github.com/user-attachments/assets/b44c888b-a8db-4dde-a9cc-8e988d2bc18f



Or [click here to watch](https://raw.githubusercontent.com/sorzkode/eksel/master/assets/eksel-demo.mp4) if the video doesn't load.

## Installation

To install the Eksel Splitter application, follow these steps:

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
pip install -r requirements.txt
```

### Dependencies

- [[Python 3.7+](https://www.python.org/downloads/)]
- [[xlwings](https://pypi.org/project/xlwings/)] - For Excel automation
- [[Pillow](https://pypi.org/project/Pillow/)] - For logo display
- [[pyperclip](https://pypi.org/project/pyperclip/)] - For clipboard operations
- [[tkinter](https://docs.python.org/3/library/tkinter.html)] - Usually included with Python, but Linux users may need to install separately

### System Requirements

- Microsoft Excel must be installed on your system
- Windows, macOS, or Linux (with Excel running via Wine or similar)

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

2. Once the application starts, click the "Select File" button.

3. Select your Excel file and click "OK".

4. The application will load all worksheets from your Excel file into SHEETBOX #1.

5. Manipulate the worksheets as needed:
   - **To save all worksheets**: Click the "Save Box #1" button
   - **To save specific worksheets**: Click on worksheet names in SHEETBOX #1 to move them to SHEETBOX #2
   - **To move worksheets back**: Click on names in SHEETBOX #2 to return them to SHEETBOX #1

6. Click the appropriate "Save Box" button to save the worksheets in that box.

7. Select a destination folder when prompted.

8. The application will save each worksheet as a separate .xlsx file.

### Additional Features

- **Clear All**: Reset the application and start with a new file
- **Help Menu**: Access usage instructions and about information
- **Error Copying**: If any error occurs, click the "Copy Error" button in the error dialog to copy the full error message to your clipboard

### Supported File Types

- *.xlsx (Excel Workbook)
- *.xlsm (Excel Macro-Enabled Workbook)
- *.xls (Excel 97-2003 Workbook)

All worksheets are saved as *.xlsx format by default.

### Notes

- The application sanitizes worksheet names to create valid filenames
- You'll be prompted before overwriting any existing files
- Excel runs in the background (hidden) for better performance

## Troubleshooting

### Common Issues

1. **"Failed to initialize Excel" error**
   - Ensure Microsoft Excel is installed on your system
   - Check that Excel is not running with elevated permissions

2. **Logo not displaying**
   - Ensure the `assets/eklogo.png` file exists in the correct location
   - Verify Pillow is installed: `pip install Pillow`

3. **Copy Error button not working**
   - Install pyperclip: `pip install pyperclip`
   - The application will fall back to tkinter's clipboard if pyperclip fails

4. **"Cannot save workbook with same name" error**
   - This should only happen if you are trying to replace a file with itself - IE you open a workbook named "sheet1" and the only worksheet is named "sheet1" and then you try to save to the same location.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

- **Mister Riley** - [sorzkode](https://github.com/sorzkode)
