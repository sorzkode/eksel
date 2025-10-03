#!/usr/bin/env python3
"""
███████ ██   ██ ███████ ███████ ██      
██      ██  ██  ██      ██      ██      
█████   █████   ███████ █████   ██      
██      ██  ██       ██ ██      ██      
███████ ██   ██ ███████ ███████ ███████ 
                                        
███████ ██████  ██      ██ ████████ ████████ ███████ ██████  
██      ██   ██ ██      ██    ██       ██    ██      ██   ██ 
███████ ██████  ██      ██    ██       ██    █████   ██████  
     ██ ██      ██      ██    ██       ██    ██      ██   ██ 
███████ ██      ███████ ██    ██       ██    ███████ ██   ██ 

Split and save worksheets from Excel.

Author: sorzkode
GitHub: https://github.com/sorzkode

MIT License
Copyright (c) 2025 sorzkode
"""

import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw
from typing import Optional, List
from PIL import Image, ImageTk
import pyperclip


class ErrorDialog(tk.Toplevel):
    """Custom error dialog with copy functionality."""
    
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.message = message
        
        # Make dialog modal
        self.grab_set()
        
        # Configure window
        self.resizable(False, False)
        
        # Create widgets
        self.create_widgets()
        
        # Center the dialog
        self.center_window()
        
    def create_widgets(self):
        """Create dialog widgets."""
        # Main frame
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Error icon and message
        msg_frame = ttk.Frame(main_frame)
        msg_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Message text
        msg_text = tk.Text(msg_frame, wrap=tk.WORD, width=50, height=8, 
                          relief=tk.FLAT, bg=self['bg'])
        msg_text.insert('1.0', self.message)
        msg_text.config(state=tk.DISABLED)
        msg_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(msg_frame, command=msg_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        msg_text.config(yscrollcommand=scrollbar.set)
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        # Copy button
        copy_btn = ttk.Button(btn_frame, text="Copy Error", 
                             command=self.copy_error)
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # OK button
        ok_btn = ttk.Button(btn_frame, text="OK", command=self.destroy)
        ok_btn.pack(side=tk.LEFT)
        
        # Bind Enter to OK
        self.bind('<Return>', lambda e: self.destroy())
        
    def copy_error(self):
        """Copy error message to clipboard."""
        try:
            pyperclip.copy(self.message)
            # Show brief confirmation
            self.after(100, lambda: self.show_copy_confirmation())
        except Exception:
            # If pyperclip fails, try tkinter clipboard
            try:
                self.clipboard_clear()
                self.clipboard_append(self.message)
                self.after(100, lambda: self.show_copy_confirmation())
            except:
                pass
    
    def show_copy_confirmation(self):
        """Show brief confirmation that text was copied."""
        confirmation = tk.Toplevel(self)
        confirmation.overrideredirect(True)
        confirmation.configure(bg='#4CAF50')
        
        label = tk.Label(confirmation, text="Copied!", bg='#4CAF50', 
                        fg='white', font=('Arial', 10, 'bold'), padx=10, pady=5)
        label.pack()
        
        # Position near copy button
        confirmation.update_idletasks()
        x = self.winfo_x() + 50
        y = self.winfo_y() + self.winfo_height() - 70
        confirmation.geometry(f"+{x}+{y}")
        
        # Auto-close after 1 second
        confirmation.after(1000, confirmation.destroy)
        
    def center_window(self):
        """Center the dialog on parent."""
        self.update_idletasks()
        
        # Get parent position
        parent = self.master
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Calculate position
        width = self.winfo_width()
        height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        self.geometry(f"+{x}+{y}")


class EkselSplitter:
    def __init__(self):
        """Initialize the application."""
        self.root = tk.Tk()
        self.root.title("Eksel Splitter")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        try:
            icon_path = Path(__file__).parent / "assets" / "split.ico"
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            pass
        
        # Initialize variables
        self.excel_file = tk.StringVar(value="Select your Excel file...")
        self.wb = None
        self.xlapp = None
        self.sheets_box1 = []
        self.sheets_box2 = []
        
        # Configure styles
        self.setup_styles()
        
        # Create GUI
        self.create_widgets()
        
        # Create Excel app instance (hidden)
        self.setup_excel()
        
        # Bind cleanup on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def setup_styles(self):
        """Configure ttk styles for the application."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        bg_color = "#f0f0f0"
        self.root.configure(bg=bg_color)
        
        style.configure("Title.TLabel", font=("Arial", 14, "bold"))
        style.configure("Info.TLabel", font=("Arial", 10, "italic"))
        
    def setup_excel(self):
        """Initialize hidden Excel application."""
        try:
            self.xlapp = xw.App(visible=False, add_book=False)
        except Exception as e:
            self.show_error("Excel Error", 
                          f"Failed to initialize Excel: {str(e)}\n\n"
                          "Please ensure Excel is installed.")
            sys.exit(1)
    
    def show_error(self, title, message):
        """Show error dialog with copy capability."""
        ErrorDialog(self.root, title, message)
    
    def create_widgets(self):
        """Create all GUI widgets."""
        # Menu bar
        self.create_menu()
        
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights for responsiveness
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        for i in range(6):
            main_frame.rowconfigure(i, weight=1 if i == 3 else 0)
        
        # Logo/Title
        logo_frame = ttk.Frame(main_frame)
        logo_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Load and display logo
        logo_path = Path(__file__).parent / "assets" / "eklogo.png"
        pil_image = Image.open(logo_path)
        
        # Resize if too large
        max_width = 400
        if pil_image.width > max_width:
            ratio = max_width / pil_image.width
            new_height = int(pil_image.height * ratio)
            pil_image = pil_image.resize((max_width, new_height), Image.Resampling.LANCZOS)
        
        # Convert to PhotoImage
        self.logo_image = ImageTk.PhotoImage(pil_image)
        logo_label = ttk.Label(logo_frame, image=self.logo_image)
        logo_label.pack()
        
        # File selection
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                       pady=(0, 20))
        file_frame.columnconfigure(1, weight=1)
        
        self.select_btn = ttk.Button(file_frame, text="Select File", 
                                   command=self.select_file)
        self.select_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.excel_file, 
                                  state="readonly")
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Listbox headers
        ttk.Label(main_frame, text="SHEETBOX #1", 
                 style="Title.TLabel").grid(row=2, column=0, pady=(0, 5))
        ttk.Label(main_frame, text="SHEETBOX #2", 
                 style="Title.TLabel").grid(row=2, column=1, pady=(0, 5))
        
        # Listboxes with scrollbars
        self.create_listboxes(main_frame)
        
        # Buttons
        self.create_buttons(main_frame)
        
        # Info label
        info_label = ttk.Label(main_frame, 
                             text="Acceptable filetypes: *.xlsx, *.xlsm, *.xls", 
                             style="Info.TLabel")
        info_label.grid(row=5, column=0, columnspan=2, pady=(10, 0))
        
    def create_menu(self):
        """Create application menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Usage", command=self.show_usage)
        help_menu.add_command(label="About", command=self.show_about)
        
    def create_listboxes(self, parent):
        """Create the two listboxes with scrollbars."""
        # Listbox 1
        frame1 = ttk.Frame(parent)
        frame1.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), 
                   padx=(0, 10))
        
        scrollbar1 = ttk.Scrollbar(frame1)
        scrollbar1.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox1 = tk.Listbox(frame1, yscrollcommand=scrollbar1.set,
                                 selectmode=tk.SINGLE, font=("Arial", 10))
        self.listbox1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar1.config(command=self.listbox1.yview)
        
        # Initial values
        for item in ["All worksheet names will appear here", 
                    "Names can be moved between Box1 and 2..",
                    "by clicking on the name"]:
            self.listbox1.insert(tk.END, item)
        
        self.listbox1.bind('<<ListboxSelect>>', self.on_box1_select)
        self.listbox1.config(state=tk.DISABLED)
        
        # Listbox 2
        frame2 = ttk.Frame(parent)
        frame2.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar2 = ttk.Scrollbar(frame2)
        scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox2 = tk.Listbox(frame2, yscrollcommand=scrollbar2.set,
                                 selectmode=tk.SINGLE, font=("Arial", 10))
        self.listbox2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar2.config(command=self.listbox2.yview)
        
        # Initial values
        for item in ["Whichever box you save", "will save all worksheets",
                    "as separate Excel files"]:
            self.listbox2.insert(tk.END, item)
        
        self.listbox2.bind('<<ListboxSelect>>', self.on_box2_select)
        self.listbox2.config(state=tk.DISABLED)
        
    def create_buttons(self, parent):
        """Create action buttons."""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(20, 0))
        
        self.save_box1_btn = ttk.Button(button_frame, text="Save Box #1", 
                                      command=self.save_box1, state=tk.DISABLED)
        self.save_box1_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_box2_btn = ttk.Button(button_frame, text="Save Box #2", 
                                      command=self.save_box2, state=tk.DISABLED)
        self.save_box2_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(button_frame, text="Clear All", 
                                  command=self.clear_all, state=tk.DISABLED)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        self.exit_btn = ttk.Button(button_frame, text="Exit", 
                                 command=self.on_closing)
        self.exit_btn.pack(side=tk.LEFT, padx=5)
        
    def select_file(self):
        """Handle file selection."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), 
                      ("All files", "*.*")]
        )
        
        if not filename:
            return
        
        # Update UI to show loading
        self.select_btn.config(state=tk.DISABLED)
        self.excel_file.set("Loading...")
        self.root.update_idletasks()  # Force UI update
        
        # Load the file directly (no threading)
        self.load_excel_file(filename)
        
    def load_excel_file(self, filename):
        """Load Excel file and populate sheet names."""
        try:
            # Clear previous data
            self.sheets_box1.clear()
            self.sheets_box2.clear()
            
            # Close previous workbook if open
            if self.wb:
                try:
                    self.wb.close()
                except:
                    pass
            
            # Open new workbook
            self.wb = self.xlapp.books.open(filename)
            
            # Get sheet names
            for sheet in self.wb.sheets:
                self.sheets_box1.append(sheet.name)
            
            # Update UI
            self.update_ui_after_load(filename)
            
        except Exception as e:
            self.show_error("Error", f"Failed to load file: {str(e)}")
            self.excel_file.set("Select your Excel file...")
            self.select_btn.config(state=tk.NORMAL)
            
    def update_ui_after_load(self, filename):
        """Update UI after file is loaded."""
        self.excel_file.set(filename)
        self.select_btn.config(state=tk.NORMAL)
        
        # Enable/update listbox 1
        self.listbox1.config(state=tk.NORMAL)
        self.listbox1.delete(0, tk.END)
        for sheet in self.sheets_box1:
            self.listbox1.insert(tk.END, sheet)
        
        # Clear listbox 2
        self.listbox2.config(state=tk.NORMAL)
        self.listbox2.delete(0, tk.END)
        
        # Enable buttons
        self.clear_btn.config(state=tk.NORMAL)
        if self.sheets_box1:
            self.save_box1_btn.config(state=tk.NORMAL)
            
    def on_box1_select(self, event):
        """Handle selection in listbox 1."""
        selection = self.listbox1.curselection()
        if not selection:
            return
        
        index = selection[0]
        sheet_name = self.listbox1.get(index)
        
        # Move to box 2
        self.sheets_box1.remove(sheet_name)
        self.sheets_box2.append(sheet_name)
        
        # Update listboxes
        self.listbox1.delete(index)
        self.listbox2.insert(tk.END, sheet_name)
        
        # Enable save box 2 if it has items
        if self.sheets_box2:
            self.save_box2_btn.config(state=tk.NORMAL)
            
    def on_box2_select(self, event):
        """Handle selection in listbox 2."""
        selection = self.listbox2.curselection()
        if not selection:
            return
        
        index = selection[0]
        sheet_name = self.listbox2.get(index)
        
        # Move back to box 1
        self.sheets_box2.remove(sheet_name)
        self.sheets_box1.append(sheet_name)
        
        # Update listboxes
        self.listbox2.delete(index)
        self.listbox1.insert(tk.END, sheet_name)
        
        # Disable save box 2 if empty
        if not self.sheets_box2:
            self.save_box2_btn.config(state=tk.DISABLED)
            
    def save_sheets(self, sheet_list: List[str], box_name: str):
        """Save sheets from the specified list."""
        if not sheet_list:
            messagebox.showwarning("Warning", f"No sheets in {box_name}")
            return
        
        folder = filedialog.askdirectory(title="Select Save Location")
        if not folder:
            return
        
        try:
            saved_count = 0
            for sheet_name in sheet_list:
                # Copy sheet to new workbook
                self.wb.sheets[sheet_name].api.Copy()
                new_wb = self.xlapp.books.active
                
                # Generate safe filename
                safe_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_'))
                safe_name = safe_name.strip()
                if not safe_name:
                    safe_name = "Sheet"
                    
                filepath = os.path.join(folder, f"{safe_name}.xlsx")
                
                # Handle existing files
                if os.path.exists(filepath):
                    result = messagebox.askyesno("File Exists", 
                                               f"'{safe_name}.xlsx' already exists. Overwrite?")
                    if not result:
                        new_wb.close(False)
                        continue
                
                new_wb.save(filepath)
                new_wb.close()
                saved_count += 1
            
            messagebox.showinfo("Success", 
                              f"Successfully saved {saved_count} worksheet(s)!")
            self.reset_after_save()
            
        except Exception as e:
            self.show_error("Error", f"Failed to save sheets: {str(e)}")
            
    def save_box1(self):
        """Save sheets from box 1."""
        self.save_sheets(self.sheets_box1.copy(), "Box #1")
        
    def save_box2(self):
        """Save sheets from box 2."""
        self.save_sheets(self.sheets_box2.copy(), "Box #2")
        
    def clear_all(self):
        """Clear all data and reset UI."""
        # Clear data
        self.sheets_box1.clear()
        self.sheets_box2.clear()
        
        # Close workbook
        if self.wb:
            try:
                self.wb.close()
            except:
                pass
            self.wb = None
        
        # Reset UI
        self.excel_file.set("Cleared...Select a new Excel file.")
        
        # Reset listboxes
        self.listbox1.config(state=tk.NORMAL)
        self.listbox1.delete(0, tk.END)
        for item in ["Worksheet names will appear here", 
                    "Names can be moved between the boxes",
                    "Click any names you want to move"]:
            self.listbox1.insert(tk.END, item)
        self.listbox1.config(state=tk.DISABLED)
        
        self.listbox2.config(state=tk.NORMAL)
        self.listbox2.delete(0, tk.END)
        for item in ["Determine which list of names to save", "",
                    "Save using the corresponding button below"]:
            self.listbox2.insert(tk.END, item)
        self.listbox2.config(state=tk.DISABLED)
        
        # Disable buttons
        self.clear_btn.config(state=tk.DISABLED)
        self.save_box1_btn.config(state=tk.DISABLED)
        self.save_box2_btn.config(state=tk.DISABLED)
        
    def reset_after_save(self):
        """Reset UI after successful save."""
        self.excel_file.set("Success! Try another file.")
        self.clear_all()
        
    def show_usage(self):
        """Show usage information."""
        usage_text = """Follow these basic steps:

        1. Click the "Select File" button
        2. Select any Excel file and click "OK"
        3. Save

        Note:
        • SHEETBOX #1 will populate with all worksheet names
        • To save all worksheets, click "Save Box #1"
        • To save specific sheets, click names to move them to SHEETBOX #2
        • Click names to move them between boxes
        • Use "Clear All" to start over

        Supported formats: *.xlsx, *.xlsm, *.xls
        Worksheets are saved as .xlsx by default"""
        
        messagebox.showinfo("Usage", usage_text)
        
    def show_about(self):
        """Show about information."""
        about_text = """Eksel Splitter
        Version 1.1.0

        Quickly copy and save worksheets as new files

        Author: sorzkode
        Website: https://github.com/sorzkode
        License: MIT

        Supported formats: *.xlsx, *.xlsm, *.xls
        Conflicting filenames require confirmation"""
        
        messagebox.showinfo("About", about_text)
        
    def on_closing(self):
        """Handle window closing."""
        # Clean up Excel
        if self.wb:
            try:
                self.wb.close()
            except:
                pass
        
        if self.xlapp:
            try:
                self.xlapp.quit()
            except:
                try:
                    self.xlapp.kill()
                except:
                    pass
        
        self.root.destroy()
        
    def run(self):
        """Start the application."""
        self.root.mainloop()


def main():
    """Main entry point."""
    app = EkselSplitter()
    app.run()


if __name__ == "__main__":
    main()