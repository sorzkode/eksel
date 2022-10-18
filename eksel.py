#!/usr/bin/env python3

'''
███████ ██   ██ ███████ ███████ ██      
██      ██  ██  ██      ██      ██      
█████   █████   ███████ █████   ██      
██      ██  ██       ██ ██      ██      
███████ ██   ██ ███████ ███████ ███████ 
                                        
                                        
███████ ██████  ██      ██ ████████ ████████ ███████ ██████  
██      ██   ██ ██      ██    ██       ██    ██      ██   ██ 
███████ ██████  ██      ██    ██       ██    █████   ██████  
     ██ ██      ██      ██    ██       ██    ██      ██   ██ 
███████ ██      ███████ ██    ██       ██    ███████ ██   ██ 
                                                             
                                                                                                             
Split and save worksheets from Excel.
-
Author:
sorzkode
sorzkode@proton.me
https://github.com/sorzkode

MIT License
Copyright (c) 2022 sorzkode
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''
# Dependencies
from tkinter.font import ITALIC, BOLD
import PySimpleGUI as sg 
import xlwings as xw  

# PySimpleGUI version info
psversion = sg.version

# Initialize variables
excel_file = ""
save_path = ""

# File selection function
def select_file():
    excel_file = sg.popup_get_file('Select your Excel File', file_types=(("Excel Files", "*.xlsx;*.xlsm;*.xls"),), title='File Selection', grab_anywhere=True, keep_on_top=True)
    
    if not excel_file: 
        sg.popup_cancel('No file selected', grab_anywhere=True, keep_on_top=True)
        return
    return excel_file

# Save location selection function
def select_folder():
    save_path = sg.popup_get_folder('Select your Excel File', title='Folder Selection', grab_anywhere=True, keep_on_top=True)
    
    if not save_path:
        sg.popup_cancel('No folder selected', grab_anywhere=True, keep_on_top=True)
        return
    return save_path

# GUI window theme
sg.theme('Default1')

# Hides active Excel application
xlapp = xw.App(visible=False)

# Default listbox values
lbox1 = ['Worksheet names will appear here', 'Names can be moved between the boxes', 'Click any names you want to move']
lbox2 = ['Determine which list of names to save', '', 'Save using the corresponding button below']

# Listox values to be used after file selection
sheetbox1 = []
sheetbox2 = []

# Application menu
app_menu = [['&Help', ['&Usage', '&About']],]

# All GUI elements (logo, listboxes, buttons etc)
layout = [[sg.Menu(app_menu, tearoff=False, key='-MENU-')],
          [sg.Image(filename='assets\eklogo.png', key='-LOGO-')],
          [sg.Button('Select File', font=('Lucida', 12, BOLD), pad=(5,15)),
          sg.In('Select your Excel file...', size=70, font=('Lucida', 11, ITALIC), text_color='Gray', readonly=True, enable_events=True, key='-XLFILE-')],
          [sg.Text('SHEETBOX #1', font=('Lucida', 14, BOLD), size=(32,1)),
          sg.Text('SHEETBOX #2', font=('Lucida', 14, BOLD), size=(32,1))], 
          [sg.Listbox(lbox1, font=('Lucida', 11, ITALIC), size=(30,5), enable_events=True, expand_x=True, expand_y=True, key='-SHEETBOX1-', disabled=True),
          sg.Listbox(lbox2, font=('Lucida', 11, ITALIC), size=(30,5), enable_events=True, expand_x=True, expand_y=True, key='-SHEETBOX2-', disabled=True)],
          [sg.Button('Save Box #1', font=('Lucida', 12, BOLD), pad=(5,15), disabled=True), 
          sg.Button('Save Box #2', font=('Lucida', 12, BOLD), pad=(5,15), disabled=True),
          sg.Button('Clear All', font=('Lucida', 12, BOLD), pad=(5,15), disabled=True),  
          sg.Button('Exit', font=('Lucida', 12, BOLD), pad=(5,15)),
          sg.Text('Acceptable filetypes: *.xlsx, *.xlsm, and *.xls', font=('Lucida', 10, ITALIC), justification='right', size=(40,1), pad=(5,15))]]

# Calls the main window / application
window = sg.Window('eksel splitter', layout, resizable=True, icon='assets\split.ico', grab_anywhere=True, keep_on_top=True)

# Event loops when buttons are pressed / actions are taken in the app
while True:
    event, values = window.read()

# Window closed event
    if event == sg.WIN_CLOSED or event == 'Exit':
        xlapp.kill() 
        break        

# File selection event
    if event == 'Select File':
        try:
            sheetbox1.clear()
            sheetbox2.clear()
            excel_file = select_file()                                             
            window['-XLFILE-'].update(excel_file, text_color='Black') 
            wb = xlapp.books.open(excel_file)
            box1_count = 0                         
            for sh in wb.sheets:                                      
                sheetbox1.append(sh.name)
                box1_count += 1                                          
            window['Clear All'].update(disabled=False)                    
            window['Save Box #1'].update(disabled=False)
            if box1_count > 1:
                window['-SHEETBOX1-'].update(sheetbox1, disabled=False)
            elif box1_count == 1:
                window['-SHEETBOX1-'].update(disabled=False) 
                window['-SHEETBOX1-'].update(sheetbox1)
                window['-SHEETBOX1-'].update(disabled=True)
        except: sg.popup('Try again', keep_on_top=True)                   

# Listbox item selection event (moves items to second listbox)
    if event == '-SHEETBOX1-':
        box1_selection = values[event]
        if box1_selection:
            box1_item = box1_selection[0]
            window['-SHEETBOX2-'].update(disabled=False)
            sheetbox2.append(box1_item)
            sheetbox1.remove(box1_item)
            window['-SHEETBOX2-'].update(sheetbox2)
            window['-SHEETBOX1-'].update(sheetbox1)
            window['Save Box #2'].update(disabled=False)

# Listbox item selection event (moves items back to the first listbox)
    if event == '-SHEETBOX2-':
        box2_selection = values[event]
        if box2_selection:
            box2_item = box2_selection[0]
            sheetbox2.remove(box2_item)
            sheetbox1.append(box2_item)
            window['-SHEETBOX2-'].update(sheetbox2)
            window['-SHEETBOX1-'].update(sheetbox1)

# Save listbox #1
    if event == 'Save Box #1':
        try:
            save_path = select_folder()
            for sheet in sheetbox1:
                wb.sheets[sheet].api.Copy()
                newbook = xw.books.active
                newbook.save(f'{save_path}/{sheet}.xlsx')
                newbook.close()
            window['Clear All'].update(disabled=True)
            window['Save Box #1'].update(disabled=True)
            window['Save Box #2'].update(disabled=True)
            window['-XLFILE-'].update('Success! Try another file.')
            window['-SHEETBOX1-'].update(lbox1)
            window['-SHEETBOX1-'].update(disabled=True)
            window['-SHEETBOX2-'].update(lbox2)
            window['-SHEETBOX2-'].update(disabled=True)
            sheetbox1.clear()
            sheetbox2.clear()
            sg.popup('Success!', keep_on_top=True)
        except: sg.popup('Try again', keep_on_top=True)

# Save listbox #2
    if event == 'Save Box #2':
        try:
            save_path = select_folder()
            for sheet in sheetbox2:
                wb.sheets[sheet].api.Copy()
                newbook = xw.books.active
                newbook.save(f'{save_path}/{sheet}.xlsx')
                newbook.close()
            window['Clear All'].update(disabled=True)
            window['Save Box #1'].update(disabled=True)
            window['Save Box #2'].update(disabled=True)
            window['-XLFILE-'].update('Success! Try another file.')
            window['-SHEETBOX1-'].update(lbox1)
            window['-SHEETBOX1-'].update(disabled=True)
            window['-SHEETBOX2-'].update(lbox2)
            window['-SHEETBOX2-'].update(disabled=True)
            sheetbox1.clear()
            sheetbox2.clear()
            sg.popup('Success!', keep_on_top=True)
        except: sg.popup('Try again', keep_on_top=True)

# Clear all button
    if event == 'Clear All':
        sheetbox1.clear()
        sheetbox2.clear()
        window['-XLFILE-'].update('Cleared...Select a new Excel file.')
        window['-SHEETBOX1-'].update(disabled=False)
        window['-SHEETBOX1-'].update(lbox1)
        window['-SHEETBOX1-'].update(disabled=True)
        window['-SHEETBOX2-'].update(disabled=False)
        window['-SHEETBOX2-'].update(lbox2)
        window['-SHEETBOX2-'].update(disabled=True)
        window['Clear All'].update(disabled=True)
        window['Save Box #1'].update(disabled=True)
        window['Save Box #2'].update(disabled=True)

# About menu selection
    if event == 'About':
        sg.popup( 
        'Quickly copy and save worksheets as new files',
        '',
        'Author: sorzkode',
        'Website: https://github.com/sorzkode',
        'License: MIT',
        '',
        'Acceptable filetypes: *.xlsx, *.xlsm, and *.xls',
        'Worksheets are saved as .xlsx by default',
        'Conflicting filenames will be automatically overwritten',
        '',
        f'PySimpleGUI Version: {psversion}',
        '', 
        grab_anywhere=True, keep_on_top=True, title='About')

# Usage menu selection
    if event == 'Usage':
        sg.popup( 
        'Follow these basic steps:',
        '',
        '1. Click the "Select File" button',
        '2. Select any Excel file and click "ok"',
        '3. Save',
        '',
        'Note:',
        'SHEETBOX #1 will populate a list of all worksheet names when a file is selected',
        'If you want to copy and save all worksheets, click the "Save Box #1" button',
        'Otherwise, click any worksheet names you want to move to SHEETBOX #2',
        'If names are moved to SHEETBOX #2, the "Save Box #2" button will be enabled',
        'Determine which listbox you want to save and use the corresponding button',
        'Names can be switched between the listboxes by clicking on them',
        'Use the "Clear All" button if you made a mistake or need to start over',
        '',
        grab_anywhere=True, keep_on_top=True, title='Usage')

window.close()