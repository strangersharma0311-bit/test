import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import numpy as np
import re
import time

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.workbook = None
        self.selected_columns_for_weighting = []
        self.create_gui()

    def create_gui(self):
        self.root.title("Excel Processor App")
        self.root.geometry("500x300")

        select_file_button = ttk.Button(self.root, text="Select Workbook", command=self.select_workbook)
        select_file_button.pack(pady=20)

        extract_button = ttk.Button(self.root, text="Extract and Merge Data", command=self.extract_and_merge_data)
        extract_button.pack(pady=20)

        apply_weights_button = ttk.Button(self.root, text="Apply Weights", command=self.apply_weights)
        apply_weights_button.pack(pady=20)

        create_results_button = ttk.Button(self.root, text="Create 演算結果 Sheet", command=self.create_enzan_kekka_sheet)
        create_results_button.pack(pady=20)

    def select_workbook(self):
        self.workbook = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.workbook:
            print(f"Selected workbook: {self.workbook}")

    def extract_and_merge_data(self):
        if not self.workbook:
            print("No workbook selected!")
            return

        start_time = time.time()

        # Load the workbook and get the sheet names
        wb = load_workbook(self.workbook)
        sheet_names = [sheet for sheet in wb.sheetnames if sheet.isnumeric()]
        
        # Load the 抽出列 sheet
        chuushutsu_df = pd.read_excel(self.workbook, sheet_name="抽出列")
        
        # Create an empty DataFrame for the ketsugou sheet
        ketsugou_df = pd.DataFrame()
        
        # Dictionary to store the column headers for each year
        header_dict = {}
        
        # Loop through each year sheet and extract the data
        for year in sorted(sheet_names, reverse=True):
            year_col = None
            
            # Find the column for the current year in 抽出列 sheet
            for col in chuushutsu_df.columns:
                if str(year) in str(col):
                    year_col = col
                    break
            
            if year_col is None:
                print(f"Year {year} not found in 抽出列 sheet!")
                continue
            
            # Get the columns to extract for the current year
            columns_to_extract = chuushutsu_df[year_col].dropna().tolist()
            print(f"Year: {year}, Columns to extract: {columns_to_extract}")
            
            # Load the year sheet
            year_df = pd.read_excel(self.workbook, sheet_name=year)
            
            # Ensure the columns exist in the year sheet
            available_columns = [col for col in columns_to_extract if col in year_df.columns]
            print(f"Available columns in {year}: {available_columns}")
            
            # Extract the columns and rename them with the year prefix
            extracted_df = year_df[['調査番号'] + available_columns]
            extracted_df.columns = ['調査番号'] + [f"{year} {col}" for col in available_columns]
            
            # Merge with the ketsugou DataFrame
            if ketsugou_df.empty:
                ketsugou_df = extracted_df
            else:
                ketsugou_df = pd.merge(ketsugou_df, extracted_df, on='調査番号', how='outer', suffixes=('', '_duplicate'))
                # Remove duplicate columns
                ketsugou_df = ketsugou_df.loc[:, ~ketsugou_df.columns.str.endswith('_duplicate')]
        
        # Sort the ketsugou DataFrame by 調査番号
        ketsugou_df = ketsugou_df.sort_values(by='調査番号')
        
        # Save the ketsugou DataFrame to a new sheet
        with pd.ExcelWriter(self.workbook, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            ketsugou_df.to_excel(writer, sheet_name='結合シート', index=False)
        
        end_time = time.time()
        print(f"結合シート has been created successfully in {end_time - start_time:.2f} seconds.")
        messagebox.showinfo("Success", "結合シート has been created successfully!")

        # Proceed to create 抽出シート
        self.create_chuushutsu_sheet()

    def create_chuushutsu_sheet(self):
        if not self.workbook:
            print("No workbook selected!")
            return

        # Load the workbook and get the sheet names
        wb = load_workbook(self.workbook)
        ketsugou_df = pd.read_excel(self.workbook, sheet_name='結合シート')
        
        # Load the 点数化列 sheet
        tensuuka_df = pd.read_excel(self.workbook, sheet_name="点数化列")
        
        # Create an empty DataFrame for the 抽出シート
        chuushutsu_df = pd.DataFrame()

        # Debug information: Display columns in ketsugou_df
        print("Columns in 結合シート (ketsugou_df):")
        print(ketsugou_df.columns.tolist())
        
        # Extract columns based on 点数化列 sheet
        for col in tensuuka_df.columns:
            year = col
            columns_to_extract = tensuuka_df[year].dropna().tolist()
            print(f"Year: {year}, Columns to extract: {columns_to_extract}")
            for col_name in columns_to_extract:
                if f"{col_name}" in ketsugou_df.columns:
                    chuushutsu_df[f"{col_name}"] = ketsugou_df[f"{col_name}"]
                else:
                    print(f"Column {col_name} not found in 結合シート")
        
        # Add 路線名 and 構造物名称 columns
        chuushutsu_df['路線名'] = ketsugou_df[[col for col in ketsugou_df.columns if '路線名' in col]].bfill(axis=1).iloc[:, 0]
        chuushutsu_df['構造物名称'] = ketsugou_df[[col for col in ketsugou_df.columns if '構造物名称' in col]].bfill(axis=1).iloc[:, 0]

        # Add 種別 column
        shubetsu_cols = [col for col in ketsugou_df.columns if '種別' in col]
        if shubetsu_cols:
            chuushutsu_df['種別'] = ketsugou_df[shubetsu_cols].bfill(axis=1).iloc[:, 0]

        # Add new columns: 点検区分1, 駅（始）, 駅（至）
        tenken_cols = [col for col in ketsugou_df.columns if '点検区分1' in col]
        if tenken_cols:
            chuushutsu_df['点検区分1'] = ketsugou_df[tenken_cols].bfill(axis=1).iloc[:, 0]

        eki_hajimari_cols = [col for col in ketsugou_df.columns if '駅（始）' in col]
        if eki_hajimari_cols:
            chuushutsu_df['駅（始）'] = ketsugou_df[eki_hajimari_cols].bfill(axis=1).iloc[:, 0]

        eki_itaru_cols = [col for col in ketsugou_df.columns if '駅（至）' in col]
        if eki_itaru_cols:
            chuushutsu_df['駅（至）'] = ketsugou_df[eki_itaru_cols].bfill(axis=1).iloc[:, 0]

        # Reorder columns
        base_cols = ['路線名', '構造物名称', '種別', '点検区分1', '駅（始）', '駅（至）']
        existing_base_cols = [col for col in base_cols if col in chuushutsu_df.columns]
        other_cols = [col for col in chuushutsu_df.columns if col not in base_cols]
        chuushutsu_df = chuushutsu_df[existing_base_cols + other_cols]

        # Save the 抽出シート DataFrame to a new sheet
        with pd.ExcelWriter(self.workbook, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            chuushutsu_df.to_excel(writer, sheet_name='抽出シート', index=False)
        
        print("抽出シート has been created successfully.")
        messagebox.showinfo("Success", "抽出シート has been created successfully!")

    def apply_weights(self):
        if not self.workbook:
            print("No workbook selected!")
            return

        try:
            # Load the workbook and get the sheet names
            wb = load_workbook(self.workbook)
            chuushutsu_df = pd.read_excel(self.workbook, sheet_name='抽出シート')
            tensuuka_df = pd.read_excel(self.workbook, sheet_name='点数化列')
            lookup_df = pd.read_excel(self.workbook, sheet_name='重みテーブル')
            
            print("=== DEBUG: Loaded DataFrames ===")
            print(f"抽出シート shape: {chuushutsu_df.shape}")
            print(f"点数化列 shape: {tensuuka_df.shape}")
            print(f"重みテーブル shape: {lookup_df.shape}")
            print(f"重みテーブル columns: {lookup_df.columns.tolist()}")
            print("\n" + "="*50 + "\n")
            
            # Check if we need to expand lookup table for additional columns
            lookup_df = self.expand_lookup_table_if_needed(lookup_df, tensuuka_df)
            
            # Create lookup dictionaries from the 重みテーブル sheet
            lookup_dicts = self.create_lookup_dicts(lookup_df)

            # Track values not found in the lookup tables
            not_found_values = [[] for _ in range(len(lookup_dicts))]

            # Process each column from 点数化列 sheet
            for col_index, col in enumerate(tensuuka_df.columns):
                print(f"Processing 点数化列 column: {col}")
                
                for i, column_name in enumerate(tensuuka_df[col].dropna()):
                    weight_col_name = f"{column_name} 重み"
                    print(f"Processing column: {column_name}")
                    print(f"Weight column name: {weight_col_name}")
                    
                    if column_name in chuushutsu_df.columns:
                        dict_index = i % len(lookup_dicts)
                        lookup_dict = lookup_dicts[dict_index]
                        print(f"Using lookup dictionary {dict_index} for column {column_name}")
                        
                        # Create weight values list
                        weight_values = []
                        for val in chuushutsu_df[column_name]:
                            weight = self.lookup_weight(lookup_dict, val, not_found_values[dict_index])
                            weight_values.append(weight)
                        
                        # Check if weight column already exists
                        if weight_col_name in chuushutsu_df.columns:
                            print(f"Weight column '{weight_col_name}' already exists. Replacing values...")
                            # Replace existing column values
                            chuushutsu_df[weight_col_name] = weight_values
                        else:
                            # Add new weight column right after the original column
                            col_position = chuushutsu_df.columns.get_loc(column_name) + 1
                            chuushutsu_df.insert(col_position, weight_col_name, weight_values)
                            print(f"Added new weight column: {weight_col_name} at position {col_position}")
                    else:
                        print(f"Column {column_name} not found in 抽出シート")

            # Handle values not found in the lookup tables
            if any(not_found_values):
                print("Found missing values, showing choice dialog...")
                self.show_missing_values_choice(not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df)
            else:
                # If no missing values, directly write to Excel
                self.write_to_excel(chuushutsu_df, tensuuka_df, lookup_df)

        except Exception as e:
            print(f"Error: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def expand_lookup_table_if_needed(self, lookup_df, tensuuka_df):
        """Expand lookup table if there are more columns in 点数化列 than lookup tables"""
        # Calculate maximum number of columns needed per year
        max_cols_per_year = 0
        for col in tensuuka_df.columns:
            col_count = len(tensuuka_df[col].dropna())
            max_cols_per_year = max(max_cols_per_year, col_count)
        
        # Calculate current number of lookup table pairs
        current_table_pairs = len(lookup_df.columns) // 2
        
        print(f"Max columns per year: {max_cols_per_year}")
        print(f"Current lookup table pairs: {current_table_pairs}")
        
        # If we need more lookup tables
        if max_cols_per_year > current_table_pairs:
            additional_tables_needed = max_cols_per_year - current_table_pairs
            print(f"Need to add {additional_tables_needed} additional lookup tables")
            
            # Create new columns for additional lookup tables
            for i in range(additional_tables_needed):
                table_num = current_table_pairs + i + 1
                key_col_name = f"Table{table_num}_Key"
                value_col_name = f"Table{table_num}_Value"
                
                # Add empty columns
                lookup_df[key_col_name] = None
                lookup_df[value_col_name] = None
                
                print(f"Added lookup table {table_num}: {key_col_name}, {value_col_name}")
        
        return lookup_df

    def create_enzan_kekka_sheet(self):
        """Create calculation results sheet based on operations"""
        if not self.workbook:
            print("No workbook selected!")
            return
            
        try:
            # Load required sheets
            chuushutsu_df = pd.read_excel(self.workbook, sheet_name='抽出シート')
            tensuuka_df = pd.read_excel(self.workbook, sheet_name='点数化列')
            
            try:
                enzanshi_df = pd.read_excel(self.workbook, sheet_name='演算子')
            except:
                print("演算子 sheet not found, using default A*B*C operation")
                enzanshi_df = pd.DataFrame()
            
            print("Creating 演算結果 sheet...")
            
            # Create result dataframe
            result_df = pd.DataFrame()
            
            # Add basic columns (路線名, 構造物名称, 種別, 点検区分1, 駅（始）, 駅（至）)
            if '路線名' in chuushutsu_df.columns:
                result_df['路線名'] = chuushutsu_df['路線名']
            else:
                # If not found directly, look for columns containing '路線名'
                rosen_cols = [col for col in chuushutsu_df.columns if '路線名' in col]
                if rosen_cols:
                    result_df['路線名'] = chuushutsu_df[rosen_cols].bfill(axis=1).iloc[:, 0]

            if '構造物名称' in chuushutsu_df.columns:
                result_df['構造物名称'] = chuushutsu_df['構造物名称']
            else:
                # If not found directly, look for columns containing '構造物名称'
                kozo_cols = [col for col in chuushutsu_df.columns if '構造物名称' in col]
                if kozo_cols:
                    result_df['構造物名称'] = chuushutsu_df[kozo_cols].bfill(axis=1).iloc[:, 0]

            # Add 種別 column
            if '種別' in chuushutsu_df.columns:
                result_df['種別'] = chuushutsu_df['種別']
            else:
                # If not found directly, look for columns containing '種別'
                shubetsu_cols = [col for col in chuushutsu_df.columns if '種別' in col]
                if shubetsu_cols:
                    result_df['種別'] = chuushutsu_df[shubetsu_cols].bfill(axis=1).iloc[:, 0]
                else:
                    result_df['種別'] = ''

            # Add new columns: 点検区分1, 駅（始）, 駅（至）
            # 点検区分1
            tenken_cols = [col for col in chuushutsu_df.columns if '点検区分1' in col]
            if tenken_cols:
                result_df['点検区分1'] = chuushutsu_df[tenken_cols].bfill(axis=1).iloc[:, 0]
            else:
                result_df['点検区分1'] = ''

            # 駅（始）
            eki_hajimari_cols = [col for col in chuushutsu_df.columns if '駅（始）' in col]
            if eki_hajimari_cols:
                result_df['駅（始）'] = chuushutsu_df[eki_hajimari_cols].bfill(axis=1).iloc[:, 0]
            else:
                result_df['駅（始）'] = ''

            # 駅（至）
            eki_itaru_cols = [col for col in chuushutsu_df.columns if '駅（至）' in col]
            if eki_itaru_cols:
                result_df['駅（至）'] = chuushutsu_df[eki_itaru_cols].bfill(axis=1).iloc[:, 0]
            else:
                result_df['駅（至）'] = ''
            
            # Process each year column in 点数化列
            for col_index, col in enumerate(tensuuka_df.columns):
                if pd.isna(col):
                    continue
                    
                year = str(col)
                print(f"Processing year: {year}")
                
                # Get the operation formula from 演算子 sheet for this year
                if not enzanshi_df.empty and col_index < len(enzanshi_df.columns):
                    if len(enzanshi_df) > 1:
                        operation_formula = enzanshi_df.iloc[1, col_index]
                    else:
                        operation_formula = "A*B*C"
                else:
                    operation_formula = "A*B*C"
                
                # Get the column names for this year from 点数化列
                year_columns = tensuuka_df[col].dropna().tolist()
                
                if len(year_columns) == 0:
                    print(f"No columns found for {year}, skipping...")
                    continue
                
                # Find the corresponding weighted columns in 抽出シート
                weight_columns = []
                for year_col in year_columns:
                    weight_col_name = f"{year_col} 重み"
                    if weight_col_name in chuushutsu_df.columns:
                        weight_columns.append(weight_col_name)
                    else:
                        print(f"Weight column not found: {weight_col_name}")
                
                if len(weight_columns) == 0:
                    print(f"No weight columns found for {year}, skipping...")
                    continue
                
                # Calculate results for each row
                result_column_name = f"{year} 結果"
                result_values = []
                
                for row_idx, row in chuushutsu_df.iterrows():
                    try:
                        # Get the values for all available weighted columns
                        values = []
                        has_blank = False
                        
                        for weight_col in weight_columns:
                            value = row[weight_col]
                            
                            if pd.isna(value) or value == '':
                                has_blank = True
                                break
                            else:
                                values.append(float(value))
                        
                        if has_blank or len(values) == 0:
                            result_values.append('')  # Blank if any value is missing
                        else:
                            # Evaluate the operation dynamically
                            if operation_formula and isinstance(operation_formula, str):
                                # Replace A, B, C, D, etc. with actual values in the formula
                                expression = operation_formula
                                for i, val in enumerate(values):
                                    letter = chr(65 + i)  # A=65, B=66, C=67, D=68, etc.
                                    expression = expression.replace(letter, str(val))
                                
                                # Evaluate the expression
                                try:
                                    result = eval(expression)
                                except:
                                    result = 0
                            else:
                                # Default to multiplication if no formula
                                result = 1
                                for val in values:
                                    result *= val
                            
                            result_values.append(result)
                            
                    except Exception as e:
                        print(f"Error calculating row {row_idx}: {e}")
                        result_values.append('')
                
                # Add the result column
                result_df[result_column_name] = result_values
                print(f"Added column: {result_column_name}")
            
            # Write to Excel
            with pd.ExcelWriter(self.workbook, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Write all existing sheets first
                chuushutsu_df.to_excel(writer, sheet_name='抽出シート', index=False)
                tensuuka_df.to_excel(writer, sheet_name='点数化列', index=False)
                
                # Load and write lookup table
                lookup_df = pd.read_excel(self.workbook, sheet_name='重みテーブル')
                lookup_df.to_excel(writer, sheet_name='重みテーブル', index=False)
                
                # Write operation sheet if exists
                if not enzanshi_df.empty:
                    enzanshi_df.to_excel(writer, sheet_name='演算子', index=False)
                
                # Write the new result sheet
                result_df.to_excel(writer, sheet_name='演算結果', index=False)
            
            print("演算結果 sheet created successfully!")
            messagebox.showinfo("Success", "演算結果 sheet has been created successfully!")
            
        except Exception as e:
            print(f"Error creating 演算結果 sheet: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Error creating 演算結果 sheet: {str(e)}")

    def show_missing_values_choice(self, not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df):
        """Show initial choice dialog for handling missing values"""
        choice_window = tk.Toplevel(self.root)
        choice_window.title("Missing Values Found")
        choice_window.geometry("600x350")
        choice_window.grab_set()
        choice_window.resizable(False, False)
        
        # Center the window
        choice_window.transient(self.root)
        choice_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 50))

        # Main frame
        main_frame = tk.Frame(choice_window, padx=30, pady=30)
        main_frame.pack(fill="both", expand=True)

        # Title
        title_label = tk.Label(main_frame, text="Missing Weight Values Found", 
                              font=("Arial", 16, "bold"), fg="red")
        title_label.pack(pady=(0, 15))

        # Message
        msg_text = "Some values in your data don't have corresponding weights in the lookup table.\n\nHow would you like to proceed?"
        msg_label = tk.Label(main_frame, text=msg_text, justify="center", wraplength=500, font=("Arial", 12))
        msg_label.pack(pady=(0, 25))

        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=15)

        def assign_values():
            choice_window.destroy()
            self.show_assign_values_dialog(not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df)

        def skip_assignment():
            choice_window.destroy()
            self.ask_default_value(not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df)

        # Made buttons smaller as requested
        assign_btn = tk.Button(button_frame, text="Assign Values", command=assign_values,
                              bg="#4CAF50", fg="white", width=14, height=2, font=("Arial", 11))
        assign_btn.pack(side="left", padx=15)

        skip_btn = tk.Button(button_frame, text="Skip", command=skip_assignment,
                            bg="#f44336", fg="white", width=14, height=2, font=("Arial", 11))
        skip_btn.pack(side="left", padx=15)

    def ask_default_value(self, not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df):
        """Ask for default value to assign to all missing values"""
        default_value = simpledialog.askfloat("Default Weight", 
                                            "Enter the default weight to assign to all missing values:",
                                            minvalue=0, maxvalue=10)
        if default_value is not None:
            # Assign default value to all missing values
            for table_index, values in enumerate(not_found_values):
                for value in values:
                    if table_index < len(lookup_dicts):
                        lookup_dicts[table_index][value] = default_value
            
            # Continue processing
            self.recalculate_and_save(lookup_dicts, chuushutsu_df, tensuuka_df, lookup_df)

    def show_assign_values_dialog(self, not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df):
        """Show dialog for assigning individual values"""
        assign_window = tk.Toplevel(self.root)
        assign_window.title("Assign Weight Values")
        assign_window.geometry("1000x700")
        assign_window.grab_set()
        assign_window.resizable(True, True)
        
        # Center the window
        assign_window.transient(self.root)

        # Main frame
        main_frame = tk.Frame(assign_window, padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)

        # Title
        title_label = tk.Label(main_frame, text="Assign Weights to Missing Values", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 15))

        # Scrollable frame for tables
        canvas = tk.Canvas(main_frame, height=500)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Create entries for missing values
        entries = []
        not_found_values = [list(set(values)) for values in not_found_values]
        
        for table_index, values in enumerate(not_found_values):
            if values and table_index * 2 < len(lookup_df.columns):
                # Table frame
                table_frame = tk.LabelFrame(scrollable_frame, 
                                          text=f"Table {table_index + 1} - {lookup_df.columns[table_index * 2]}", 
                                          padx=15, pady=15, font=("Arial", 12, "bold"))
                table_frame.pack(fill="x", padx=15, pady=15)

                # Header
                header_frame = tk.Frame(table_frame)
                header_frame.pack(fill="x", pady=(0, 10))
                
                tk.Label(header_frame, text="Value", font=("Arial", 11, "bold"), width=35, anchor="w").pack(side="left")
                tk.Label(header_frame, text="Weight", font=("Arial", 11, "bold"), width=15, anchor="w").pack(side="left")

                # Values
                for value in values:
                    value_frame = tk.Frame(table_frame)
                    value_frame.pack(fill="x", pady=3)
                    
                    value_label = tk.Label(value_frame, text=str(value), width=35, anchor="w", 
                                         relief="solid", borderwidth=1, font=("Arial", 10))
                    value_label.pack(side="left", padx=(0, 10))
                    
                    entry = ttk.Combobox(value_frame, values=[i for i in range(21)], 
                                       state="normal", width=15, font=("Arial", 10))
                    entry.pack(side="left")
                    entries.append((value, entry, table_index))

        # Button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=20)

        def submit_values():
            # Get values from entries before destroying window
            entry_data = []
            for value, entry, table_index in entries:
                weight_text = entry.get()
                entry_data.append((value, weight_text, table_index))
            
            # Group new values by table_index
            new_values_by_table = {}
            for value, weight_text, table_index in entry_data:
                if weight_text:
                    try:
                        weight = float(weight_text)
                        if table_index < len(lookup_dicts):
                            lookup_dicts[table_index][value] = weight
                            
                            if table_index not in new_values_by_table:
                                new_values_by_table[table_index] = []
                            new_values_by_table[table_index].append((value, weight))
                            
                    except ValueError:
                        print(f"Invalid weight value for {value}: {weight_text}")
            
            # Add new values to lookup_df table by table
            for table_index, new_values in new_values_by_table.items():
                if table_index * 2 + 1 < len(lookup_df.columns):
                    key_col = lookup_df.columns[table_index * 2]
                    value_col = lookup_df.columns[table_index * 2 + 1]
                    
                    # Find the last row with data in this table
                    last_row = lookup_df[key_col].last_valid_index()
                    if last_row is not None:
                        # Add one separator blank line
                        separator_row_index = last_row + 2
                    else:
                        separator_row_index = 0
                    
                    # Extend dataframe if needed
                    total_new_rows = len(new_values) + 1  # +1 for separator
                    while len(lookup_df) <= separator_row_index + total_new_rows:
                        lookup_df.loc[len(lookup_df)] = [None] * len(lookup_df.columns)
                    
                    # Add separator row (blank line)
                    lookup_df.loc[separator_row_index, key_col] = ""
                    lookup_df.loc[separator_row_index, value_col] = ""
                    
                    # Add all new values consecutively after separator
                    for i, (value, weight) in enumerate(new_values):
                        new_row_index = separator_row_index + 1 + i
                        lookup_df.loc[new_row_index, key_col] = value
                        lookup_df.loc[new_row_index, value_col] = weight
                        print(f"Added new value to lookup table {table_index}: {value} = {weight}")
            
            assign_window.destroy()
            
            # Check for unassigned values
            unassigned_values = []
            for value, weight_text, table_index in entry_data:
                if not weight_text:
                    unassigned_values.append(value)
            
            if unassigned_values:
                self.handle_unassigned_values(unassigned_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df, not_found_values)
            else:
                # All values assigned, proceed
                self.recalculate_and_save(lookup_dicts, chuushutsu_df, tensuuka_df, lookup_df)

        submit_btn = tk.Button(button_frame, text="Submit", command=submit_values,
                            bg="#4CAF50", fg="white", width=20, height=2, font=("Arial", 12))
        submit_btn.pack(side="right", padx=15)

    def handle_unassigned_values(self, unassigned_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df, not_found_values):
        """Handle unassigned values after submission"""
        # Show dialog asking about unassigned values
        unassigned_window = tk.Toplevel(self.root)
        unassigned_window.title("Unassigned Values")
        unassigned_window.geometry("500x300")
        unassigned_window.grab_set()
        unassigned_window.resizable(False, False)
        
        # Center the window
        unassigned_window.transient(self.root)
        unassigned_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 150, self.root.winfo_rooty() + 100))

        main_frame = tk.Frame(unassigned_window, padx=25, pady=25)
        main_frame.pack(fill="both", expand=True)

        title_label = tk.Label(main_frame, text="Unassigned Values Found", 
                              font=("Arial", 14, "bold"), fg="orange")
        title_label.pack(pady=(0, 15))

        msg_text = f"Some values are still not assigned:\n{', '.join(unassigned_values[:5])}{'...' if len(unassigned_values) > 5 else ''}\n\nWhat would you like to do?"
        msg_label = tk.Label(main_frame, text=msg_text, justify="center", wraplength=450, font=("Arial", 11))
        msg_label.pack(pady=(0, 20))

        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=15)

        def assign_default():
            unassigned_window.destroy()
            self.ask_default_value_for_remaining(unassigned_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df)

        def go_back():
            unassigned_window.destroy()
            self.show_assign_values_dialog(not_found_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df)

        default_btn = tk.Button(button_frame, text="Assign Default Value", command=assign_default,
                               bg="#2196F3", fg="white", width=18, height=2, font=("Arial", 11))
        default_btn.pack(side="left", padx=10)

        back_btn = tk.Button(button_frame, text="Go Back", command=go_back,
                            bg="#FF9800", fg="white", width=18, height=2, font=("Arial", 11))
        back_btn.pack(side="left", padx=10)

    def ask_default_value_for_remaining(self, unassigned_values, lookup_dicts, lookup_df, chuushutsu_df, tensuuka_df):
        """Ask for default value for remaining unassigned values"""
        default_window = tk.Toplevel(self.root)
        default_window.title("Default Value for Remaining")
        default_window.geometry("450x250")
        default_window.grab_set()
        default_window.resizable(False, False)
        
        # Center the window
        default_window.transient(self.root)
        default_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 200, self.root.winfo_rooty() + 150))

        main_frame = tk.Frame(default_window, padx=25, pady=25)
        main_frame.pack(fill="both", expand=True)

        tk.Label(main_frame, text="Enter default weight for unassigned values:", 
                font=("Arial", 12)).pack(pady=15)

        entry_frame = tk.Frame(main_frame)
        entry_frame.pack(pady=15)

        tk.Label(entry_frame, text="Default Weight:", font=("Arial", 11)).pack(side="left")
        default_entry = tk.Entry(entry_frame, width=12, font=("Arial", 11))
        default_entry.pack(side="left", padx=15)
        default_entry.focus()

        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=25)

        def submit_default():
            try:
                default_value = float(default_entry.get())
                # Assign default value to unassigned values
                for value in unassigned_values:
                    # Assign to all lookup dictionaries since we don't know which table each value belongs to
                    for table_index in range(len(lookup_dicts)):
                        lookup_dicts[table_index][value] = default_value
                
                default_window.destroy()
                self.recalculate_and_save(lookup_dicts, chuushutsu_df, tensuuka_df, lookup_df)
                
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid number")

        submit_btn = tk.Button(button_frame, text="Submit", command=submit_default,
                            bg="#4CAF50", fg="white", width=18, height=2, font=("Arial", 12))
        submit_btn.pack()

        # Handle Enter key
        default_entry.bind('<Return>', lambda e: submit_default())

    def recalculate_and_save(self, lookup_dicts, chuushutsu_df, tensuuka_df, lookup_df):
        """Recalculate weights with updated lookup tables and save"""
        try:
            print("Recalculating weights with updated lookup tables...")
            
            # Recalculate all weight columns with updated lookup dictionaries
            for col_index, col in enumerate(tensuuka_df.columns):
                for i, column_name in enumerate(tensuuka_df[col].dropna()):
                    weight_col_name = f"{column_name} 重み"
                    
                    if column_name in chuushutsu_df.columns and weight_col_name in chuushutsu_df.columns:
                        dict_index = i % len(lookup_dicts)
                        lookup_dict = lookup_dicts[dict_index]
                        
                        # Recalculate weight values
                        weight_values = []
                        for val in chuushutsu_df[column_name]:
                            weight = self.lookup_weight(lookup_dict, val, None)  # None to avoid adding to not_found
                            weight_values.append(weight)
                        
                        # Update the weight column
                        chuushutsu_df[weight_col_name] = weight_values
                        print(f"Updated weight column: {weight_col_name}")

            # Write to Excel
            self.write_to_excel(chuushutsu_df, tensuuka_df, lookup_df)
            
        except Exception as e:
            print(f"Error in recalculation: {str(e)}")
            messagebox.showerror("Error", f"Error during recalculation: {str(e)}")

    def write_to_excel(self, chuushutsu_df, tensuuka_df, lookup_df):
        """Write the updated DataFrames back to Excel"""
        print("Writing updated data back to Excel...")
        
        # Create a new Excel writer
        with pd.ExcelWriter(self.workbook, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # Write the updated 抽出シート
            chuushutsu_df.to_excel(writer, sheet_name='抽出シート', index=False)
            
            # Also write the other sheets to preserve them
            tensuuka_df.to_excel(writer, sheet_name='点数化列', index=False)
            lookup_df.to_excel(writer, sheet_name='重みテーブル', index=False)
        
        print("Excel file updated successfully!")
        messagebox.showinfo("Success", f"Weights have been applied successfully!\nUpdated file: {self.workbook}")

    def create_lookup_dicts(self, lookup_df):
        lookup_dicts = []
        print("=== DEBUG: Creating Lookup Dictionaries ===")
        
        # Process pairs of columns (key, value)
        for i in range(0, len(lookup_df.columns), 2):
            if i + 1 < len(lookup_df.columns):
                key_col = lookup_df.columns[i]
                value_col = lookup_df.columns[i + 1]
                
                print(f"Processing pair: {key_col} -> {value_col}")
                
                # Create dictionary with proper key-value pairs
                lookup_dict = {}
                for idx, row in lookup_df.iterrows():
                    key = row[key_col]
                    value = row[value_col]
                    
                    # Skip NaN values
                    if pd.isna(key) or pd.isna(value):
                        continue
                    
                    # Convert key to string and apply character conversion
                    key_str = str(key).strip()
                    key_converted = self.convert_to_hankaku(key_str)
                    
                    # Store the mapping
                    lookup_dict[key_converted] = int(value) if pd.notna(value) else 0
                
                lookup_dicts.append(lookup_dict)
                print(f"Dictionary {len(lookup_dicts)-1} created with {len(lookup_dict)} entries")
        
        print("=== END DEBUG: Creating Lookup Dictionaries ===\n")
        return lookup_dicts

    def extract_first_value(self, text):
        """Extract the first value before delimiters (mimicking VBA ExtractFirstValue)"""
        if pd.isna(text) or text == '':
            return ''
        
        text = str(text).strip()
        
        # Define delimiters (Japanese and English commas, and other separators)
        delimiters = ['、', ',', ',', ' ', '　']  # Japanese comma, English comma, full-width comma, space, full-width space
        
        # Find the position of the first delimiter
        min_pos = len(text)
        for delimiter in delimiters:
            pos = text.find(delimiter)
            if pos != -1 and pos < min_pos:
                min_pos = pos
        
        # Extract the first value before the delimiter
        if min_pos < len(text):
            result = text[:min_pos].strip()
        else:
            result = text.strip()
            
        return result

    def lookup_weight(self, lookup_dict, value, not_found_values):
        # Handle NaN or empty values
        if pd.isna(value) or value == 'nan' or value == '' or value is None:
            return ''  # Return blank for NaN or empty values
        
        # Convert to string and extract first value
        value_str = str(value).strip()
        first_value = self.extract_first_value(value_str)
        
        # Convert to hankaku
        converted_value = self.convert_to_hankaku(first_value)
        
        # Check if the converted value exists in lookup dictionary
        if converted_value in lookup_dict:
            weight = lookup_dict[converted_value]
            return weight
        
        # If not found, try original value without conversion
        if first_value in lookup_dict:
            weight = lookup_dict[first_value]
            return weight
        
        # Try exact match with original input
        if value_str in lookup_dict:
            weight = lookup_dict[value_str]
            return weight
        
        # Add to not_found_values if not found and list is provided
        if not_found_values is not None and converted_value not in not_found_values:
            not_found_values.append(converted_value)
        return 1  # Default weight if not found

    def convert_to_hankaku(self, text):
        """Convert full-width characters to half-width (mimicking VBA ConvertToHankaku)"""
        if pd.isna(text) or text == '':
            return ''
        
        text = str(text)
        hankaku = ""
        
        for char in text:
            code = ord(char)
            # Full-width ASCII characters
            if 0xFF01 <= code <= 0xFF5E:
                hankaku += chr(code - 0xFF00 + 0x20)
            # Full-width digits
            elif 0xFF10 <= code <= 0xFF19:
                hankaku += chr(code - 0xFF10 + 0x30)
                        # Full-width uppercase letters  
            elif 0xFF21 <= code <= 0xFF3A:
                hankaku += chr(code - 0xFF21 + 0x41)
            # Full-width lowercase letters
            elif 0xFF41 <= code <= 0xFF5A:
                hankaku += chr(code - 0xFF41 + 0x61)
            else:
                hankaku += char
        
        return hankaku

# Create the main window
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()