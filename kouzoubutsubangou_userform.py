import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import re

class StructureDataEntryApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Structure Data Entry System")
        self.root.geometry("450x300")
        self.root.minsize(400, 250)
        
        # Center the window on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.root.winfo_screenheight() // 2) - (300 // 2)
        self.root.geometry(f"450x300+{x}+{y}")
        
        self.workbook_path = None
        self.grouped_df = None
        self.structure_data_df = None
        
        self.create_main_gui()
    
    def create_main_gui(self):
        """Create main GUI for file selection"""
        main_frame = tk.Frame(self.root, padx=30, pady=30)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Structure Data Entry System", 
                              font=("Arial", 14, "bold"), fg="navy")
        title_label.pack(pady=(0, 15))
        
        # Simple instruction
        instruction_label = tk.Label(main_frame, text="Select workbook with 'グループ化点検履歴' sheet", 
                                   font=("Arial", 10))
        instruction_label.pack(pady=(0, 15))
        
        # Status label
        self.status_label = tk.Label(main_frame, text="Ready...", 
                                    font=("Arial", 9), fg="gray")
        self.status_label.pack(pady=(0, 10))
        
        # Select file button
        select_btn = tk.Button(main_frame, text="Browse & Select File", 
                             command=self.select_workbook_with_feedback, 
                             bg="#4CAF50", fg="white", 
                             width=20, height=1, font=("Arial", 10))
        select_btn.pack(pady=8)
        
        # Structure data button (initially disabled)
        self.structure_btn = tk.Button(main_frame, text="Enter Structure Data", 
                                     command=self.show_structure_data_form, 
                                     bg="#9C27B0", fg="white", 
                                     width=20, height=1, font=("Arial", 10),
                                     state="disabled")
        self.structure_btn.pack(pady=8)
        
        # Exit button
        exit_btn = tk.Button(main_frame, text="Exit", 
                           command=self.confirm_exit, bg="#f44336", fg="white", 
                           width=12, height=1, font=("Arial", 9))
        exit_btn.pack(pady=(15, 0))
    
    def select_workbook_with_feedback(self):
        """Select workbook with user feedback"""
        self.status_label.config(text="Opening browser...", fg="blue")
        self.root.update()
        
        # File selection
        self.workbook_path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=os.path.expanduser("~")
        )
        
        if not self.workbook_path:
            self.status_label.config(text="No file selected", fg="orange")
            return
        
        self.status_label.config(text="Loading...", fg="blue")
        self.root.update()
        
        # Validate in background
        self.root.after(100, self.validate_workbook)

    def validate_workbook(self):
        """Validate workbook with progress feedback"""
        try:
            if not os.path.exists(self.workbook_path):
                raise Exception("File not found")
            
            self.status_label.config(text="Checking sheets...", fg="blue")
            self.root.update()
            
            # Validate required sheets
            wb = load_workbook(self.workbook_path)
            required_sheet = 'グループ化点検履歴'
            
            if required_sheet not in wb.sheetnames:
                self.status_label.config(text="Sheet not found!", fg="red")
                messagebox.showerror("Error", f"Sheet '{required_sheet}' not found")
                self.status_label.config(text="Ready...", fg="gray")
                return
            
            self.status_label.config(text="Loading data...", fg="blue")
            self.root.update()
            
            # Load data
            self.grouped_df = pd.read_excel(self.workbook_path, sheet_name=required_sheet)
            
            if len(self.grouped_df) == 0:
                raise Exception("Sheet is empty")
            
            # Load structure data if exists
            self.load_structure_data()
            
            self.status_label.config(text="File ready!", fg="green")
            self.structure_btn.config(state="normal")
            
            # Show success message
            messagebox.showinfo("Success", f"File loaded\nRecords: {len(self.grouped_df):,}")
            
        except Exception as e:
            self.status_label.config(text="Error", fg="red")
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Ready...", fg="gray")

    def load_structure_data(self):
        """Load existing structure data sheet"""
        try:
            self.structure_data_df = pd.read_excel(self.workbook_path, sheet_name='構造物番号')
            # Ensure all required columns exist
            required_columns = [
                '路線名', '構造物名称', '駅間', '構造物番号', '長さ(m)', 
                '構造形式', '構造形式_重み', '角度', '角度_重み', 
                '供用年数', '供用年数_重み'
            ]
            for col in required_columns:
                if col not in self.structure_data_df.columns:
                    self.structure_data_df[col] = ''
        except:
            # Create empty structure data sheet with headers
            self.structure_data_df = pd.DataFrame(columns=[
                '路線名', '構造物名称', '駅間', '構造物番号', '長さ(m)', 
                '構造形式', '構造形式_重み', '角度', '角度_重み', 
                '供用年数', '供用年数_重み'
            ])
            self.save_structure_data()

    def get_missing_structure_entries(self):
        """Get ALL unique structure names and station intervals that need data entry"""
        missing_entries = []
        
        # Ensure structure_data_df has required columns
        if len(self.structure_data_df) == 0:
            self.structure_data_df = pd.DataFrame(columns=[
                '路線名', '構造物名称', '駅間', '構造物番号', '長さ(m)', 
                '構造形式', '構造形式_重み', '角度', '角度_重み', 
                '供用年数', '供用年数_重み'
            ])
        
        # Get ALL unique structure names and station intervals
        unique_kozo = set()
        unique_ekikan = set()
        
        for _, row in self.grouped_df.iterrows():
            rosen = str(row.get('路線名', '')).strip() if pd.notna(row.get('路線名', '')) else ''
            group_method = str(row.get('グループ化方法', '')).strip() if pd.notna(row.get('グループ化方法', '')) else ''
            
            if group_method == '構造物名称':
                # This row is grouped by structure name
                kozo = str(row.get('構造物名称', '')).strip() if pd.notna(row.get('構造物名称', '')) else ''
                if kozo and kozo not in ['', 'nan', 'NaN']:
                    unique_kozo.add((rosen, kozo))
            
            elif group_method == '駅間':
                # This row is grouped by station interval - extract from grouping key or create from 駅始→駅至
                ekikan_start = str(row.get('駅（始）', '')).strip() if pd.notna(row.get('駅（始）', '')) else ''
                ekikan_end = str(row.get('駅（至）', '')).strip() if pd.notna(row.get('駅（至）', '')) else ''
                
                if ekikan_start and ekikan_end and ekikan_start not in ['', 'nan', 'NaN'] and ekikan_end not in ['', 'nan', 'NaN']:
                    ekikan = f"{ekikan_start}→{ekikan_end}"
                    unique_ekikan.add((rosen, ekikan))
        
        print(f"Found {len(unique_kozo)} unique structure names")
        print(f"Found {len(unique_ekikan)} unique station intervals")
        
        # Check which structure names are missing
        for rosen, kozo in unique_kozo:
            if len(self.structure_data_df) == 0:
                exists = False
            else:
                exists = not self.structure_data_df[
                    (self.structure_data_df['構造物名称'].astype(str).str.strip() == kozo) & 
                    (self.structure_data_df['路線名'].astype(str).str.strip() == rosen)
                ].empty
            
            if not exists:
                missing_entries.append({
                    'type': '構造物名称',
                    'rosen': rosen,
                    'value': kozo,
                    'display_value': kozo
                })
        
        # Check which station intervals are missing
        for rosen, ekikan in unique_ekikan:
            if len(self.structure_data_df) == 0:
                exists = False
            else:
                exists = not self.structure_data_df[
                    (self.structure_data_df['駅間'].astype(str).str.strip() == ekikan) & 
                    (self.structure_data_df['路線名'].astype(str).str.strip() == rosen)
                ].empty
            
            if not exists:
                missing_entries.append({
                    'type': '駅間',
                    'rosen': rosen,
                    'value': ekikan,
                    'display_value': ekikan
                })
        
        print(f"Total missing entries: {len(missing_entries)}")
        
        # Sort entries: structure names first, then station intervals
        missing_entries.sort(key=lambda x: (x['type'] == '駅間', x['rosen'], x['value']))
        
        return missing_entries

    def show_structure_data_form(self):
        """Show structure data input form in Excel-like table format"""
        missing_entries = self.get_missing_structure_entries()
        
        if not missing_entries:
            messagebox.showinfo("Info", "All structure data is already entered!")
            return
        
        # Create form window
        form_window = tk.Toplevel(self.root)
        form_window.title("Structure Data Entry")
        form_window.geometry("1200x700")
        form_window.grab_set()
        form_window.resizable(True, True)
        form_window.transient(self.root)
        
        # Center window
        form_window.update_idletasks()
        x = (form_window.winfo_screenwidth() // 2) - (1200 // 2)
        y = (form_window.winfo_screenheight() // 2) - (700 // 2)
        form_window.geometry(f"1200x700+{x}+{y}")
        
        main_frame = tk.Frame(form_window, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Enter Structure Data", 
                              font=("Arial", 12, "bold"), fg="navy")
        title_label.pack(pady=(0, 10))
        
        # Count info
        kozo_count = len([e for e in missing_entries if e['type'] == '構造物名称'])
        ekikan_count = len([e for e in missing_entries if e['type'] == '駅間'])
        info_text = f"Found {kozo_count} structure names + {ekikan_count} station intervals = {len(missing_entries)} total entries"
        info_label = tk.Label(main_frame, text=info_text, font=("Arial", 10), fg="blue")
        info_label.pack(pady=(0, 10))
        
        # Create scrollable frame for table
        canvas = tk.Canvas(main_frame, height=500)
        scrollbar_v = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollbar_h = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        
        # Pack scrollbars and canvas
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_v.pack(side="right", fill="y")
        scrollbar_h.pack(side="bottom", fill="x")
        
        # Separate entries by type
        kozo_entries = [e for e in missing_entries if e['type'] == '構造物名称']
        ekikan_entries = [e for e in missing_entries if e['type'] == '駅間']
        
        # Store all entry widgets
        self.entry_widgets = {}
        
        current_row = 0
        
        # Section 1: 構造物名称 entries
        if kozo_entries:
            # Section header
            section_label = tk.Label(scrollable_frame, text=f"構造物名称 Section ({len(kozo_entries)} entries)", 
                                   font=("Arial", 10, "bold"), fg="white", bg="navy",
                                   relief="solid", borderwidth=1, height=2)
            section_label.grid(row=current_row, column=0, columnspan=10, sticky="ew", padx=1, pady=2)
            current_row += 1
            
            # Column headers
            headers = ['路線名', '構造物名称', '構造物番号', '長さ(m)', '構造形式', 
                      '構造形式_重み', '角度', '角度_重み', '供用年数', '供用年数_重み']
            
            for col, header in enumerate(headers):
                header_label = tk.Label(scrollable_frame, text=header, 
                                      font=("Arial", 8, "bold"), bg="lightgray",
                                      relief="solid", borderwidth=1, width=12)
                header_label.grid(row=current_row, column=col, sticky="ew", padx=1, pady=1)
            current_row += 1
            
            # Data rows for 構造物名称
            for entry in kozo_entries:
                item_key = f"kozo_{entry['value']}_{entry['rosen']}"
                self.entry_widgets[item_key] = {
                    'type': '構造物名称',
                    'rosen': entry['rosen'],
                    'main_value': entry['value'],
                    'widgets': {}
                }
                
                # 路線名 (display only)
                rosen_label = tk.Label(scrollable_frame, text=entry['rosen'], 
                                     font=("Arial", 8), bg="white",
                                     relief="solid", borderwidth=1, width=12)
                rosen_label.grid(row=current_row, column=0, sticky="ew", padx=1, pady=1)
                
                # 構造物名称 (display only)
                kozo_label = tk.Label(scrollable_frame, text=entry['value'], 
                                    font=("Arial", 8), bg="white",
                                    relief="solid", borderwidth=1, width=12)
                kozo_label.grid(row=current_row, column=1, sticky="ew", padx=1, pady=1)
                
                # Input fields
                input_fields = ['構造物番号', '長さ(m)', '構造形式', '構造形式_重み', 
                               '角度', '角度_重み', '供用年数', '供用年数_重み']
                
                for col, field in enumerate(input_fields, 2):
                    entry_widget = tk.Entry(scrollable_frame, width=12, font=("Arial", 8),
                                          relief="solid", borderwidth=1)
                    entry_widget.grid(row=current_row, column=col, sticky="ew", padx=1, pady=1)
                    self.entry_widgets[item_key]['widgets'][field] = entry_widget
                
                current_row += 1
        
        # Section 2: 駅間 entries
        if ekikan_entries:
            # Empty row separator
            tk.Label(scrollable_frame, text="", height=1).grid(row=current_row, column=0, columnspan=10)
            current_row += 1
            
            # Section header
            section_label = tk.Label(scrollable_frame, text=f"駅間 Section ({len(ekikan_entries)} entries)", 
                                   font=("Arial", 10, "bold"), fg="white", bg="darkgreen",
                                   relief="solid", borderwidth=1, height=2)
            section_label.grid(row=current_row, column=0, columnspan=10, sticky="ew", padx=1, pady=2)
            current_row += 1
            
            # Column headers
            headers = ['路線名', '駅間', '構造物番号', '長さ(m)', '構造形式', 
                      '構造形式_重み', '角度', '角度_重み', '供用年数', '供用年数_重み']
            
            for col, header in enumerate(headers):
                header_label = tk.Label(scrollable_frame, text=header, 
                                      font=("Arial", 8, "bold"), bg="lightgray",
                                      relief="solid", borderwidth=1, width=12)
                header_label.grid(row=current_row, column=col, sticky="ew", padx=1, pady=1)
            current_row += 1
            
            # Data rows for 駅間
            for entry in ekikan_entries:
                item_key = f"ekikan_{entry['value']}_{entry['rosen']}"
                self.entry_widgets[item_key] = {
                    'type': '駅間',
                    'rosen': entry['rosen'],
                    'main_value': entry['value'],
                    'widgets': {}
                }
                
                # 路線名 (display only)
                rosen_label = tk.Label(scrollable_frame, text=entry['rosen'], 
                                     font=("Arial", 8), bg="white",
                                     relief="solid", borderwidth=1, width=12)
                rosen_label.grid(row=current_row, column=0, sticky="ew", padx=1, pady=1)
                
                # 駅間 (display only)
                ekikan_label = tk.Label(scrollable_frame, text=entry['value'], 
                                      font=("Arial", 8), bg="white",
                                      relief="solid", borderwidth=1, width=12)
                ekikan_label.grid(row=current_row, column=1, sticky="ew", padx=1, pady=1)
                
                # Input fields
                input_fields = ['構造物番号', '長さ(m)', '構造形式', '構造形式_重み', 
                               '角度', '角度_重み', '供用年数', '供用年数_重み']
                
                for col, field in enumerate(input_fields, 2):
                    entry_widget = tk.Entry(scrollable_frame, width=12, font=("Arial", 8),
                                          relief="solid", borderwidth=1)
                    entry_widget.grid(row=current_row, column=col, sticky="ew", padx=1, pady=1)
                    self.entry_widgets[item_key]['widgets'][field] = entry_widget
                
                current_row += 1
        
        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Skip button
        skip_btn = tk.Button(button_frame, text="Skip (Use Defaults)", 
                           command=lambda: self.show_default_values_dialog(missing_entries, form_window), 
                           bg="#FF9800", fg="white", width=18, height=1, font=("Arial", 9))
        skip_btn.pack(side="left", padx=5)
        
        # Save button
        save_btn = tk.Button(button_frame, text="Save & Continue", 
                           command=lambda: self.save_table_data_and_close(form_window), 
                           bg="#4CAF50", fg="white", width=15, height=1, font=("Arial", 9))
        save_btn.pack(side="right", padx=5)
        
        # Cancel button
        cancel_btn = tk.Button(button_frame, text="Cancel", 
                             command=form_window.destroy, bg="#f44336", fg="white", 
                             width=10, height=1, font=("Arial", 9))
        cancel_btn.pack(side="right", padx=5)

    def show_default_values_dialog(self, missing_entries, form_window):
        """Show dialog to enter default values for all columns"""
        default_window = tk.Toplevel(form_window)
        default_window.title("Set Default Values")
        default_window.geometry("400x300")
        default_window.grab_set()
        default_window.resizable(False, False)
        default_window.transient(form_window)
        
        # Center window
        default_window.update_idletasks()
        x = (default_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (default_window.winfo_screenheight() // 2) - (300 // 2)
        default_window.geometry(f"400x300+{x}+{y}")
        
        main_frame = tk.Frame(default_window, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Title - This line was missing or incomplete
        title_label = tk.Label(main_frame, text="Enter Default Values", 
                      font=("Arial", 12, "bold"), fg="navy")
        title_label.pack(pady=(0, 15))
        
        # Info
        info_label = tk.Label(main_frame, text=f"These values will be applied to all {len(missing_entries)} entries", 
                            font=("Arial", 9))
        info_label.pack(pady=(0, 15))
            
        # Create entry fields for defaults
        self.default_entries = {}
        
        fields_frame = tk.Frame(main_frame)
        fields_frame.pack(fill="both", expand=True)
        
        default_values = [
            ('構造物番号', '1'),
            ('長さ(m)', '100'),
            ('構造形式', 'RC'),
            ('構造形式_重み', '1'),
            ('角度', '0'),
            ('角度_重み', '1'),
            ('供用年数', '50'),
            ('供用年数_重み', '10')
        ]
        
        for i, (field, default_val) in enumerate(default_values):
            row = i // 2
            col = (i % 2) * 2
            
            # Label
            label = tk.Label(fields_frame, text=f"{field}:", font=("Arial", 9))
            label.grid(row=row, column=col, sticky="w", padx=5, pady=5)
            
            # Entry
            entry = tk.Entry(fields_frame, width=10, font=("Arial", 9))
            entry.insert(0, default_val)
            entry.grid(row=row, column=col+1, padx=5, pady=5)
            self.default_entries[field] = entry
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        def apply_defaults():
            try:
                # Get default values
                defaults = {}
                for field, entry_widget in self.default_entries.items():
                    defaults[field] = entry_widget.get().strip()
                
                # Apply to all missing entries
                for entry in missing_entries:
                    new_row = {
                        '路線名': entry['rosen'],
                        '構造物名称': entry['value'] if entry['type'] == '構造物名称' else '',
                        '駅間': entry['value'] if entry['type'] == '駅間' else ''
                    }
                    
                    # Add default values
                    for field, value in defaults.items():
                        new_row[field] = value
                    
                    # Add to dataframe
                    self.structure_data_df = pd.concat([
                        self.structure_data_df, 
                        pd.DataFrame([new_row])
                    ], ignore_index=True)
                
                # Save to Excel
                self.save_structure_data()
                
                # Close windows
                default_window.destroy()
                form_window.destroy()
                
                # Show success
                messagebox.showinfo("Success", f"Default values applied to {len(missing_entries)} entries!")
                
            except Exception as e:
                messagebox.showerror("Error", str(e))
        
        apply_btn = tk.Button(button_frame, text="Apply Defaults", 
                            command=apply_defaults, bg="#4CAF50", fg="white", 
                            width=15, height=1, font=("Arial", 10))
        apply_btn.pack(side="right", padx=5)
        
        cancel_btn = tk.Button(button_frame, text="Cancel", 
                             command=default_window.destroy, bg="#f44336", fg="white", 
                             width=10, height=1, font=("Arial", 10))
        cancel_btn.pack(side="right", padx=5)

    def save_table_data_and_close(self, form_window):
        """Save data from table entries and close form"""
        try:
            saved_count = 0
            
            for item_key, entry_data in self.entry_widgets.items():
                widgets = entry_data['widgets']
                entry_type = entry_data['type']
                rosen = entry_data['rosen']
                main_value = entry_data['main_value']
                
                # Check if any field has data
                has_data = any(widget.get().strip() for widget in widgets.values())
                
                if has_data:
                    new_row = {
                        '路線名': rosen,
                        '構造物名称': main_value if entry_type == '構造物名称' else '',
                        '駅間': main_value if entry_type == '駅間' else ''
                    }
                    
                    # Add field values
                    for field_name, widget in widgets.items():
                        new_row[field_name] = widget.get().strip()
                    
                    # Add to dataframe
                    self.structure_data_df = pd.concat([
                        self.structure_data_df, 
                        pd.DataFrame([new_row])
                    ], ignore_index=True)
                    
                    saved_count += 1
            
            if saved_count > 0:
                # Save to Excel
                self.save_structure_data()
                messagebox.showinfo("Success", f"Saved {saved_count} entries successfully!")
            else:
                messagebox.showwarning("Warning", "No data entered")
            
            # Close form
            form_window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_structure_data(self):
        """Save structure data to Excel"""
        try:
            with pd.ExcelWriter(self.workbook_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Write structure data sheet
                self.structure_data_df.to_excel(writer, sheet_name='構造物番号', index=False)
                
                # Preserve other sheets
                try:
                    original_wb = load_workbook(self.workbook_path)
                    for sheet_name in original_wb.sheetnames:
                        if sheet_name != '構造物番号':
                            try:
                                df_temp = pd.read_excel(self.workbook_path, sheet_name=sheet_name)
                                df_temp.to_excel(writer, sheet_name=sheet_name, index=False)
                            except Exception as e:
                                continue
                except Exception as e:
                    pass
                    
        except Exception as e:
            raise Exception(f"Error saving structure data: {str(e)}")

    def confirm_exit(self):
        """Confirm before exiting"""
        if messagebox.askyesno("Exit", "Exit application?"):
            self.root.quit()

    def run(self):
        """Run the application"""
        self.root.mainloop()


# Main execution
if __name__ == "__main__":
    app = StructureDataEntryApp()
    app.run()