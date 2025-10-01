import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re

class EnhancedCombinedProcessorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Enhanced Combined Max Function & 補修考慮 Processor")
        self.root.geometry("500x350")
        self.root.minsize(450, 300)
        
        # Center the window on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.root.winfo_screenheight() // 2) - (350 // 2)
        self.root.geometry(f"500x350+{x}+{y}")
        
        self.workbook_path = None
        self.grouped_df = None
        self.structure_df = None
        
        self.create_main_gui()
    
    def abbreviate_sen_name(self, sen_name):
        """Convert route name to abbreviation"""
        if pd.isna(sen_name) or sen_name == '':
            return ''
        
        sen_name = str(sen_name).strip()
        
        abbreviation_map = {
            "東急多摩川線": "TM",
            "多摩川線": "TM", 
            "東横線": "TY",
            "大井町線": "OM",
            "池上線": "IK",
            "田園都市線": "DT",
            "目黒線": "MG",
            "こどもの国線": "KD",
            "世田谷線": "SG"
        }
        
        return abbreviation_map.get(sen_name, sen_name)
    
    def lookup_structure_number(self, structure_df, rosen_name, kozo_name, ekikan):
        """Lookup 構造物番号 from structure sheet"""
        try:
            if structure_df is None or len(structure_df) == 0:
                return ''
                
            rosen_name = str(rosen_name).strip() if pd.notna(rosen_name) else ''
            
            # First try to match by structure name
            if kozo_name and str(kozo_name).strip() not in ['', 'nan', 'NaN']:
                kozo_name = str(kozo_name).strip()
                matches = structure_df[
                    (structure_df['構造物名称'].astype(str).str.strip() == kozo_name) & 
                    (structure_df['路線名'].astype(str).str.strip() == rosen_name)
                ]
                
                if not matches.empty:
                    bangou = matches.iloc[0]['構造物番号']
                    if pd.notna(bangou) and str(bangou).strip() not in ['', 'nan']:
                        return str(bangou).strip()
            
            # If not found by structure name, try by station interval
            if ekikan and str(ekikan).strip() not in ['', 'nan', 'NaN']:
                ekikan = str(ekikan).strip()
                matches = structure_df[
                    (structure_df['駅間'].astype(str).str.strip() == ekikan) & 
                    (structure_df['路線名'].astype(str).str.strip() == rosen_name)
                ]
                
                if not matches.empty:
                    bangou = matches.iloc[0]['構造物番号']
                    if pd.notna(bangou) and str(bangou).strip() not in ['', 'nan']:
                        return str(bangou).strip()
            
            return ''
            
        except Exception as e:
            print(f"Error finding structure number: {e}")
            return ''
    
    def add_enhanced_columns(self, df, structure_df=None):
        """Add enhanced columns: 路線名略称 and 構造物番号"""
        enhanced_df = df.copy()
        
        # Add 路線名略称 column
        if '路線名' in enhanced_df.columns:
            enhanced_df['路線名略称'] = enhanced_df['路線名'].apply(self.abbreviate_sen_name)
        else:
            enhanced_df['路線名略称'] = ''
        
        # Add 構造物番号 column
        enhanced_df['構造物番号'] = ''
        
        if structure_df is not None:
            for index, row in enhanced_df.iterrows():
                rosen_name = row.get('路線名', '')
                kozo_name = row.get('構造物名称', '')
                
                # Create ekikan for lookup
                ekikan = ''
                if row.get('駅（始）', '') and row.get('駅（至）', ''):
                    ekikan = f"{row.get('駅（始）', '')}→{row.get('駅（至）', '')}"
                
                # Lookup structure number
                bangou = self.lookup_structure_number(structure_df, rosen_name, kozo_name, ekikan)
                enhanced_df.at[index, '構造物番号'] = bangou
        
        return enhanced_df
    
    def reorder_columns_enhanced(self, df):
        """Reorder columns: グループ化キー → グループ化方法 → 種別 → 構造物名称 → 駅（始） → 駅（至） → 点検区分1 → データ件数 → 路線名 → 路線名略称 → 構造物番号 → years"""
        
        # Define the correct enhanced column order
        priority_columns = [
            'グループ化キー',
            'グループ化方法', 
            '種別',
            '構造物名称',
            '駅（始）',
            '駅（至）',
            '点検区分1',
            'データ件数',
            '路線名',
            '路線名略称',
            '構造物番号'
        ]
        
        # Get year columns (ending with '結果' or containing years)
        year_columns = []
        
        for col in df.columns:
            if str(col).endswith('結果') or any(year in str(col) for year in ['2018', '2019', '2020', '2021', '2022', '2023', '2024']):
                year_columns.append(col)
        
        # Sort year columns chronologically
        def extract_year(col_name):
            try:
                for year in ['2018', '2019', '2020', '2021', '2022', '2023', '2024']:
                    if year in str(col_name):
                        return int(year)
                return 0
            except:
                return 0
        
        year_columns.sort(key=extract_year)
        
        # Create final column order
        final_columns = []
        
        # Add priority columns that exist
        for col in priority_columns:
            if col in df.columns:
                final_columns.append(col)
        
        # Add year columns
        final_columns.extend(year_columns)
        
        # Add any remaining columns that weren't in priority or year columns
        remaining_columns = [col for col in df.columns if col not in final_columns]
        final_columns.extend(remaining_columns)
        
        # Return reordered dataframe
        return df[final_columns]
    
    def create_main_gui(self):
        """Create main GUI for file selection"""
        main_frame = tk.Frame(self.root, padx=30, pady=30)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Enhanced Combined Processor", 
                              font=("Arial", 14, "bold"), fg="navy")
        title_label.pack(pady=(0, 15))
        
        # Enhanced features info
        features_text = ("Enhanced Features:\n"
                        "• Adds 路線名略称 and 構造物番号 columns\n"
                        "• Proper column ordering\n"
                        "• Processes both 補修無視 and 補修考慮 sheets")
        features_label = tk.Label(main_frame, text=features_text, 
                                font=("Arial", 10), justify="center")
        features_label.pack(pady=(0, 15))
        
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
        
        # Process button (initially disabled)
        self.process_btn = tk.Button(main_frame, text="Generate Enhanced Sheets", 
                                   command=self.process_both_functions, 
                                   bg="#FF9800", fg="white", 
                                   width=20, height=1, font=("Arial", 10),
                                   state="disabled")
        self.process_btn.pack(pady=8)
        
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
            
            # Try to load structure data for enhancements
            try:
                self.structure_df = pd.read_excel(self.workbook_path, sheet_name='構造物番号')
                print("Found 構造物番号 sheet - enhanced features enabled")
            except:
                self.structure_df = None
                print("No 構造物番号 sheet found - basic features only")
            
            # Find year result columns
            year_columns = self.find_year_result_columns()
            
            if len(year_columns) == 0:
                raise Exception("No year columns found")
            
            enhancement_status = "with enhancements" if self.structure_df is not None else "basic version"
            self.status_label.config(text="File ready!", fg="green")
            self.process_btn.config(state="normal")
            
            messagebox.showinfo("Success", f"File loaded ({enhancement_status})\nRecords: {len(self.grouped_df):,}")
            
        except Exception as e:
            self.status_label.config(text="Error", fg="red")
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Ready...", fg="gray")

    def find_year_result_columns(self):
        """Find columns that end with '結果' and contain years"""
        year_columns = []
        year_pattern = re.compile(r'(20\d{2})\s*結果')
        
        for col in self.grouped_df.columns:
            if str(col).endswith('結果'):
                match = year_pattern.search(str(col))
                if match:
                    year = int(match.group(1))
                    year_columns.append((year, col))
        
        # Sort by year (oldest to newest)
        year_columns.sort(key=lambda x: x[0])
        
        return [col for year, col in year_columns]

    def process_both_functions(self):
        """Process both max function and hoshuu kouryou logic with enhancements"""
        try:
            # Show small progress dialog
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Processing Enhanced Sheets")
            progress_window.geometry("300x100")
            progress_window.grab_set()
            progress_window.resizable(False, False)
            progress_window.transient(self.root)
            
            # Center the progress window
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (300 // 2)
            y = (progress_window.winfo_screenheight() // 2) - (100 // 2)
            progress_window.geometry(f"300x100+{x}+{y}")
            
            progress_frame = tk.Frame(progress_window, padx=15, pady=15)
            progress_frame.pack(fill="both", expand=True)
            
            # Status label
            status_label = tk.Label(progress_frame, text="Processing with enhancements...", font=("Arial", 9))
            status_label.pack(pady=3)
            
            # Progress bar
            progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
            progress_bar.pack(fill="x", pady=3)
            progress_bar.start()
            
            # Process data
            self.root.after(100, lambda: self.execute_both_functions(progress_window))
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def execute_both_functions(self, progress_window):
        """Execute both max function and hoshuu kouryou logic with enhancements"""
        try:
            # Find year result columns
            year_columns = self.find_year_result_columns()
            
            if len(year_columns) == 0:
                raise Exception("No year columns found")
            
            # Apply Max Function with enhancements
            max_result_df = self.apply_max_function_enhanced(year_columns)
            
            # Apply Hoshuu Kouryou with enhancements
            hoshuu_result_df = self.apply_hoshuu_kouryou_enhanced(year_columns)
            
            # Save both enhanced sheets
            self.save_both_enhanced_results(max_result_df, hoshuu_result_df)
            
            # Close progress window automatically
            progress_window.destroy()
            
            # Show completion dialog
            self.root.after(200, lambda: self.show_enhanced_completion_dialog(len(max_result_df), len(year_columns)))
            
        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("Error", str(e))

    def apply_max_function_enhanced(self, year_columns):
        """Apply max function logic with enhanced columns"""
        result_df = self.grouped_df.copy()
        
        # Apply max function logic row by row
        for index, row in result_df.iterrows():
            previous_value = None
            
            for i, year_col in enumerate(year_columns):
                current_value = row[year_col]
                
                if i == 0:
                    if pd.notna(current_value) and str(current_value).strip() != '':
                        try:
                            previous_value = float(current_value)
                            result_df.at[index, year_col] = previous_value
                        except (ValueError, TypeError):
                            result_df.at[index, year_col] = current_value
                            previous_value = None
                    else:
                        result_df.at[index, year_col] = current_value
                        previous_value = None
                else:
                    if pd.notna(current_value) and str(current_value).strip() != '':
                        try:
                            current_numeric = float(current_value)
                            
                            if previous_value is not None:
                                if current_numeric < previous_value:
                                    result_df.at[index, year_col] = previous_value
                                else:
                                    previous_value = current_numeric
                                    result_df.at[index, year_col] = current_numeric
                            else:
                                previous_value = current_numeric
                                result_df.at[index, year_col] = current_numeric
                                
                        except (ValueError, TypeError):
                            result_df.at[index, year_col] = current_value
                    else:
                        result_df.at[index, year_col] = current_value
        
        # Add enhanced columns
        enhanced_df = self.add_enhanced_columns(result_df, self.structure_df)
        
        # Reorder columns
        final_df = self.reorder_columns_enhanced(enhanced_df)
        
        return final_df

    def apply_hoshuu_kouryou_enhanced(self, year_columns):
        """Apply hoshuu kouryou logic with enhanced columns"""
        result_df = self.grouped_df.copy()
        
        # Apply hoshuu kouryou logic row by row
        for index, row in result_df.iterrows():
            previous_value = None
            
            for i, year_col in enumerate(year_columns):
                current_value = row[year_col]
                
                if i == 0:
                    if pd.notna(current_value) and str(current_value).strip() != '':
                        try:
                            previous_value = float(current_value)
                            result_df.at[index, year_col] = previous_value
                        except (ValueError, TypeError):
                            result_df.at[index, year_col] = current_value
                            previous_value = None
                    else:
                        result_df.at[index, year_col] = current_value
                        previous_value = None
                else:
                    if pd.notna(current_value) and str(current_value).strip() != '':
                        try:
                            current_numeric = float(current_value)
                            
                            if previous_value is not None:
                                if current_numeric < previous_value:
                                    # Set ALL previous values to 0.1
                                    for j in range(i):
                                        prev_year_col = year_columns[j]
                                        result_df.at[index, prev_year_col] = 0.1
                                    
                                    # Set current value
                                    result_df.at[index, year_col] = current_numeric
                                    previous_value = current_numeric
                                else:
                                    previous_value = current_numeric
                                    result_df.at[index, year_col] = current_numeric
                            else:
                                previous_value = current_numeric
                                result_df.at[index, year_col] = current_numeric
                                
                        except (ValueError, TypeError):
                            result_df.at[index, year_col] = current_value
                    else:
                        result_df.at[index, year_col] = current_value
        
        # Add enhanced columns
        enhanced_df = self.add_enhanced_columns(result_df, self.structure_df)
        
        # Reorder columns
        final_df = self.reorder_columns_enhanced(enhanced_df)
        
        return final_df

    def save_both_enhanced_results(self, max_result_df, hoshuu_result_df):
        """Save both enhanced results to Excel sheets"""
        try:
            max_sheet_name = "補修無視"
            hoshuu_sheet_name = "補修考慮"
            
            with pd.ExcelWriter(self.workbook_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                max_result_df.to_excel(writer, sheet_name=max_sheet_name, index=False)
                hoshuu_result_df.to_excel(writer, sheet_name=hoshuu_sheet_name, index=False)
                
                try:
                    original_wb = load_workbook(self.workbook_path)
                    for sheet_name in original_wb.sheetnames:
                        if sheet_name not in [max_sheet_name, hoshuu_sheet_name]:
                            try:
                                df_temp = pd.read_excel(self.workbook_path, sheet_name=sheet_name)
                                df_temp.to_excel(writer, sheet_name=sheet_name, index=False)
                            except Exception as e:
                                continue
                except Exception as e:
                    pass
                    
        except Exception as e:
            raise Exception(f"Error saving enhanced results: {str(e)}")

    def show_enhanced_completion_dialog(self, total_rows, year_columns_count):
        """Show enhanced completion dialog"""
        completion_window = tk.Toplevel(self.root)
        completion_window.title("Enhanced Processing Complete")
        completion_window.geometry("400x300")
        completion_window.grab_set()
        completion_window.resizable(False, False)
        completion_window.transient(self.root)
        
        # Center window
        completion_window.update_idletasks()
        x = (completion_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (completion_window.winfo_screenheight() // 2) - (300 // 2)
        completion_window.geometry(f"400x300+{x}+{y}")
        
        main_frame = tk.Frame(completion_window, padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Enhanced Processing Complete!", 
                              font=("Arial", 12, "bold"), fg="green")
        title_label.pack(pady=(0, 10))
        
        # Enhanced features info
        features_text = ("✅ Enhanced Features Applied:\n\n"
                        "• 路線名略称 column added\n"
                        "• 構造物番号 column added\n"
                        "• Proper column ordering\n"
                        "• データ件数 → 路線名 → 路線名略称 → 構造物番号")
        features_label = tk.Label(main_frame, text=features_text, font=("Arial", 10), 
                                 justify="left", fg="blue")
        features_label.pack(pady=(0, 10))
        
        # Processing info
        info_text = f"Records: {total_rows:,}\nSheets created: 2 (enhanced)\n補修無視 & 補修考慮"
        info_label = tk.Label(main_frame, text=info_text, font=("Arial", 10))
        info_label.pack(pady=(0, 15))
        
        # Enhancement status
        enhancement_status = "with 構造物番号 lookup" if self.structure_df is not None else "basic version (no 構造物番号 sheet)"
        status_label = tk.Label(main_frame, text=f"Enhancement: {enhancement_status}", 
                               font=("Arial", 9), fg="green" if self.structure_df is not None else "orange")
        status_label.pack(pady=(0, 15))
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack()
        
        def open_excel():
            try:
                import os
                os.startfile(self.workbook_path)
                messagebox.showinfo("Excel Opened", 
                                  "✅ Enhanced Excel file opened!\n\n"
                                  "Check the new enhanced columns:\n"
                                  "• 補修無視 sheet\n"
                                  "• 補修考慮 sheet")
                completion_window.after(1000, completion_window.destroy)
                self.root.after(2000, self.root.quit)
            except:
                messagebox.showinfo("Info", f"Please open file manually:\n{self.workbook_path}")
                completion_window.destroy()

        def close_only():
            completion_window.destroy()
            messagebox.showinfo("Complete", 
                              "✅ Enhanced processing completed successfully!\n\n"
                              "Both 補修無視 and 補修考慮 sheets created with:\n"
                              "• 路線名略称 columns\n"
                              "• 構造物番号 columns\n"
                              "• Proper column ordering")
            self.root.after(1000, self.root.quit)
        
        excel_btn = tk.Button(button_frame, text="Open Enhanced Excel", 
                            command=open_excel, bg="#4CAF50", fg="white", 
                            width=15, height=1, font=("Arial", 10))
        excel_btn.pack(side="left", padx=5)
        
        close_btn = tk.Button(button_frame, text="Complete", 
                            command=close_only, bg="#2196F3", fg="white", 
                            width=12, height=1, font=("Arial", 10))
        close_btn.pack(side="left", padx=5)

    def confirm_exit(self):
        """Confirm before exiting"""
        if messagebox.askyesno("Exit", "Exit enhanced application?"):
            self.root.quit()

    def run(self):
        """Run the enhanced application"""
        self.root.mainloop()


# Main execution
if __name__ == "__main__":
    print("Enhanced Combined Max Function & 補修考慮 Processor Starting...")
    print("=" * 60)
    print("🚀 Enhanced Features:")
    print("• 路線名略称 column (TM, TY, OM, IK, DT, MG, KD, SG)")
    print("• 構造物番号 column (auto-lookup from 構造物番号 sheet)")
    print("• Column order: データ件数 → 路線名 → 路線名略称 → 構造物番号 → others → years")
    print("• Processes both 補修無視 (max function) and 補修考慮 (repair consideration)")
    print("• Enhanced with proper column positioning")
    print("=" * 60)
    print("Required input:")
    print("• グループ化点検履歴 sheet (from enhanced grouping)")
    print("• 構造物番号 sheet (optional, for enhanced features)")
    print("=" * 60)
    print("Output:")
    print("• 補修無視 sheet (enhanced with new columns)")
    print("• 補修考慮 sheet (enhanced with new columns)")
    print("=" * 60)
    
    app = EnhancedCombinedProcessorApp()
    app.run()