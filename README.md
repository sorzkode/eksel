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

### Option 1: Using the Standalone Executable (Recommended)

**No Python installation required!**

1. Download the latest release from the [Releases page](https://github.com/sorzkode/eksel/releases)
2. Extract `eksel.exe` from the archive
3. Double-click `eksel.exe` to run

**System Requirements:**
- Windows operating system
- Microsoft Excel installed on your system

### Option 2: Running from Source

**For developers or users who prefer running from source:**

1. Clone or download the repository from [GitHub](https://github.com/sorzkode/eksel)

2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   source .venv/bin/activate  # Linux/Mac
   ```

3. Install the package and dependencies:
   ```bash
   pip install -e .
   ```

**Dependencies:**
- Python 3.8+
- xlwings - For Excel automation
- Pillow - For logo display
- pyperclip - For clipboard operations
- tkinter - Usually included with Python

**System Requirements:**
- Microsoft Excel must be installed on your system
- Windows (primary), macOS, or Linux (with Excel via Wine)

## Usage

### If using the executable:
Simply double-click `eksel.exe` to launch the application.

### If running from source:
```bash
python eksel.py
```

### Using the application:

1. Once the application starts, click the "Select File" button.

2. Select your Excel file and click "OK".

3. The application will load all worksheets from your Excel file into SHEETBOX #1.

4. Manipulate the worksheets as needed:
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
