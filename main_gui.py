"""
AB1 Sequence Renamer v. 1.0, 25.01.2025
This program is intended to rename .ab1 sequence chromatogram files
contained in a directory provided by user based on the Excel table
with mapping of old file name to new name. the same new name also
substitutes the old internal sequence name contained in .ab1 files.

MIT License

Copyright (c) 2025 Oleg Shchepin

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import utils
import os
import pandas as pd


class AB1RenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AB1 Sequence Renamer")
        self.headers = []
        self.sheet_names = []

        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_widgets()
        self.setup_bindings()

    def create_widgets(self):
        # Excel File Selection
        ttk.Label(self.main_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W)
        self.excel_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.excel_path, width=50).grid(row=0, column=1)
        ttk.Button(self.main_frame, text="Browse", command=self.select_excel).grid(row=0, column=2)

        # Sheet Selection
        ttk.Label(self.main_frame, text="Excel Sheet:").grid(row=1, column=0, sticky=tk.W)
        self.sheet_name = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(self.main_frame, textvariable=self.sheet_name, state="disabled")
        self.sheet_dropdown.grid(row=1, column=1, sticky=tk.W)

        # AB1 Files Directory
        ttk.Label(self.main_frame, text="AB1 Directory:").grid(row=2, column=0, sticky=tk.W)
        self.ab1_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.ab1_path, width=50).grid(row=2, column=1)
        ttk.Button(self.main_frame, text="Browse", command=self.select_ab1_dir).grid(row=2, column=2)

        # Output Directory
        ttk.Label(self.main_frame, text="Output Directory:").grid(row=3, column=0, sticky=tk.W)
        self.output_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.output_path, width=50).grid(row=3, column=1)
        ttk.Button(self.main_frame, text="Browse", command=self.select_output_dir).grid(row=3, column=2)

        # Header Selection
        ttk.Label(self.main_frame, text="Old Name Header:").grid(row=4, column=0, sticky=tk.W)
        self.old_name_header = tk.StringVar()
        self.old_name_dropdown = ttk.Combobox(self.main_frame, textvariable=self.old_name_header, state="disabled")
        self.old_name_dropdown.grid(row=4, column=1, sticky=tk.W)

        ttk.Label(self.main_frame, text="New Name Header:").grid(row=5, column=0, sticky=tk.W)
        self.new_name_header = tk.StringVar()
        self.new_name_dropdown = ttk.Combobox(self.main_frame, textvariable=self.new_name_header, state="disabled")
        self.new_name_dropdown.grid(row=5, column=1, sticky=tk.W)

        # Run Button
        ttk.Button(self.main_frame, text="Run Renaming", command=self.run_renaming).grid(row=6, column=1, pady=20)

        # About Button
        ttk.Button(self.main_frame, text="About", command=self.show_about).grid(row=7, column=1, pady=10)

    def setup_bindings(self):
        self.sheet_name.trace_add('write', self.on_sheet_selected)

    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path.set(file_path)
            self.load_sheet_names()

    def load_sheet_names(self):
        try:
            excel_file = pd.ExcelFile(self.excel_path.get())
            self.sheet_names = excel_file.sheet_names
            self.sheet_dropdown['values'] = self.sheet_names
            self.sheet_dropdown.config(state="readonly")

            if not self.sheet_names:
                messagebox.showerror("Error", "The Excel file contains no sheets.")
                self.reset_sheet_selection()
                return

            # Find first valid sheet silently
            valid_sheet = None
            for sheet in self.sheet_names:
                try:
                    # Directly check headers without triggering GUI updates
                    utils.find_header_line(self.excel_path.get(), sheet_name=sheet)
                    valid_sheet = sheet
                    break
                except Exception:
                    continue

            if valid_sheet:
                self.sheet_name.set(valid_sheet)
                self.load_headers()
            else:
                messagebox.showerror("Error",
                                     "No valid header line found in any sheet.\n"
                                     "Please ensure at least one sheet contains:\n"
                                     "1. A header row with non-empty first two columns\n"
                                     "2. Valid 'Macrogen' and 'Real name' columns")
                self.reset_sheet_selection()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file:\n{str(e)}")
            self.reset_sheet_selection()

    def on_sheet_selected(self, *args):
        if self.sheet_name.get():
            self.load_headers()

    def load_headers(self):
        try:
            sheet = self.sheet_name.get()
            _, self.headers = utils.find_header_line(self.excel_path.get(), sheet_name=sheet)
            clean_headers = [h for h in self.headers if h.strip()]

            self.old_name_dropdown['values'] = clean_headers
            self.new_name_dropdown['values'] = clean_headers
            self.old_name_dropdown.config(state="readonly")
            self.new_name_dropdown.config(state="readonly")

            # Set defaults without validation
            self.old_name_header.set(clean_headers[0] if clean_headers else '')
            self.new_name_header.set(clean_headers[1] if len(clean_headers) > 1 else '')

        except Exception as e:
            self.headers = []
            self.old_name_dropdown.config(state="disabled")
            self.new_name_dropdown.config(state="disabled")
            # Don't show error here - already handled in sheet selection

    def reset_sheet_selection(self):
        self.sheet_names = []
        self.sheet_dropdown['values'] = []
        self.sheet_dropdown.config(state="disabled")
        self.headers = []
        self.old_name_dropdown.config(state="disabled")
        self.new_name_dropdown.config(state="disabled")

    def select_ab1_dir(self):
        dir_path = filedialog.askdirectory(title="Select AB1 Files Directory")
        if dir_path:
            self.ab1_path.set(dir_path)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="Select Output Directory")
        if dir_path:
            self.output_path.set(dir_path)

    def show_about(self):
        about_text = (
            "AB1 Sequence Renamer v. 1.0, 25.01.2025\n"
            "Copyright (c) 2025 Oleg Shchepin\n\n"
            "MIT License\n\n"
            "This program is free software: you can redistribute it and/or modify "
            "it under the terms of the MIT License. https://opensource.org/license/MIT"
        )
        messagebox.showinfo("About", about_text)

    def run_renaming(self):
        try:
            # Validate inputs
            if not all([self.excel_path.get(), self.ab1_path.get(), self.output_path.get()]):
                raise ValueError("All paths must be specified")
            if not self.old_name_header.get() or not self.new_name_header.get():
                raise ValueError("Both header fields must be selected")

            # Create output directory
            os.makedirs(self.output_path.get(), exist_ok=True)

            # Process files
            sheet = self.sheet_name.get()
            header_pos, _ = utils.find_header_line(self.excel_path.get(), sheet_name=sheet)
            name_dict, df_clean = utils.create_mapping(
                self.excel_path.get(),
                header_pos,
                input_col=self.old_name_header.get(),
                output_col=self.new_name_header.get(),
                sheet_name=sheet
            )

            # Save mapping table
            csv_path = os.path.join(self.output_path.get(), "mapping_table.csv")
            df_clean.to_csv(csv_path, index=False)

            # Process AB1 files
            ab1_list = utils.get_ab1_file_list(self.ab1_path.get())
            missing_files = []

            count_successful = len(ab1_list)
            for ab1 in ab1_list:
                try:
                    original_name = ab1.rsplit('.', 1)[0]  # Remove .ab1 extension
                    new_name = name_dict[original_name]
                except KeyError:
                    missing_files.append(ab1)
                    count_successful -= 1
                    continue

                # Sanitize the new filename
                sanitized_name = utils.sanitize_filename(new_name)

                # Process file
                ab1_byte_list = utils.change_internal_name(
                    os.path.join(self.ab1_path.get(), ab1),
                    sanitized_name
                )
                utils.save_renamed_ab1(
                    os.path.join(self.output_path.get(), f"{sanitized_name}.ab1"),
                    ab1_byte_list
                )

            # Show success message with warnings if needed
            success_msg = f"{count_successful} files successfully processed!\n\n" \
                          f"Renamed AB1 files saved to: {self.output_path.get()}\n\n" \
                          f"Mapping table saved to: {csv_path}"

            if missing_files:
                success_msg += "\n\nWARNING: Missing mappings for:\n- " + \
                               "\n- ".join(missing_files)

            messagebox.showinfo("Success", success_msg)

        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AB1RenamerApp(root)
    root.mainloop()
