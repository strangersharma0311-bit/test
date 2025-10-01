import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

class ObserFileGeneratorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Obser Files Generator")
        self.root.geometry("500x400")
        self.root.configure(bg="#f0f0f0")
        
        self.workbook_path = None
        self.nyuuryoku_params = {
            'data_count': 8,
            'prediction_years': 10,
            'lambda_constant': 0.02,
            'inspection_years': list(range(27, 43))
        }
        
        # Sheet mappings
        self.sheet_mappings = {
            'obser1.txt': '割算結果(補修考慮)',
            'obser2.txt': '割算結果(補修無視)', 
            'obser3.txt': '補修無視',
            'obser4.txt': '補修考慮',
            'obser5.txt': '新しい演算(補修無視)',
            'obser6.txt': '新しい演算(補修考慮)',
            'obser7.txt': '割算結果-新しい演算(補修無視)',
            'obser8.txt': '割算結果-新しい演算(補修考慮)'
        }
        
        self.create_main_gui()
    
    def create_main_gui(self):
        """Create main GUI"""
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=30, pady=30)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Obser Files Generator", 
                              font=("Arial", 16, "bold"), fg="#1565C0", bg="#f0f0f0")
        title_label.pack(pady=(0, 20))
        
        # File selection
        self.status_label = tk.Label(main_frame, text="Ready to select Excel file...", 
                                    font=("Arial", 10), fg="#666", bg="#f0f0f0")
        self.status_label.pack(pady=(0, 15))
        
        select_btn = tk.Button(main_frame, text="Select Excel File", 
                             command=self.select_workbook, 
                             bg="#4CAF50", fg="white", 
                             width=25, height=2, font=("Arial", 11))
        select_btn.pack(pady=5)
        
        # Actions frame (initially hidden)
        self.actions_frame = tk.Frame(main_frame, bg="#f0f0f0")
        
        tk.Label(self.actions_frame, text="What would you like to do?", 
                font=("Arial", 12, "bold"), fg="#333", bg="#f0f0f0").pack(pady=(20, 10))
        
        # Action buttons
        btn_frame = tk.Frame(self.actions_frame, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="View/Edit Parameters", 
                 command=self.show_parameter_editor,
                 bg="#2196F3", fg="white", width=20, height=2, 
                 font=("Arial", 11)).pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="Generate Files", 
                 command=self.generate_obser_files,
                 bg="#FF9800", fg="white", width=20, height=2, 
                 font=("Arial", 11)).pack(side="left", padx=5)

    def select_workbook(self):
        """Select and validate workbook"""
        self.workbook_path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not self.workbook_path:
            return
        
        try:
            wb = load_workbook(self.workbook_path)
            required_sheets = list(self.sheet_mappings.values())
            missing_sheets = [sheet for sheet in required_sheets 
                            if sheet not in wb.sheetnames]
            
            if missing_sheets:
                messagebox.showerror("Missing Sheets", 
                                   f"Required sheets not found:\n" + 
                                   "\n".join(missing_sheets))
                return
            
            # Load parameters from 入力値 sheet
            self.load_nyuuryoku_parameters()
            
            self.status_label.config(text=f"File loaded: {os.path.basename(self.workbook_path)}", 
                                   fg="#4CAF50")
            self.actions_frame.pack(fill="x", pady=20)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file:\n{str(e)}")

    def load_nyuuryoku_parameters(self):
        """Load parameters from 入力値 sheet"""
        try:
            # Load 入力値 sheet
            nyuuryoku_df = pd.read_excel(self.workbook_path, sheet_name='入力値', header=None)
            
            # Find parameters by matching column headers
            if len(nyuuryoku_df) >= 2:
                headers = nyuuryoku_df.iloc[0]  # First row as headers
                
                # Find and extract parameters
                for i, header in enumerate(headers):
                    if pd.notna(header):
                        header_str = str(header)
                        if 'データ個数' in header_str:
                            try:
                                self.nyuuryoku_params['data_count'] = int(nyuuryoku_df.iloc[1, i])
                            except (ValueError, TypeError):
                                pass
                        elif '予測年数' in header_str:
                            try:
                                self.nyuuryoku_params['prediction_years'] = int(nyuuryoku_df.iloc[1, i])
                            except (ValueError, TypeError):
                                pass
                        elif 'λ定数' in header_str:
                            try:
                                self.nyuuryoku_params['lambda_constant'] = float(nyuuryoku_df.iloc[1, i])
                            except (ValueError, TypeError):
                                pass
                        elif '点検年度に対応した年' in header_str:
                            # Extract years from this column
                            years = []
                            for row_idx in range(1, len(nyuuryoku_df)):
                                val = nyuuryoku_df.iloc[row_idx, i]
                                if pd.notna(val):
                                    try:
                                        year = int(val)
                                        if 20 <= year <= 50:
                                            years.append(year)
                                    except (ValueError, TypeError):
                                        break
                            
                            if years:
                                self.nyuuryoku_params['inspection_years'] = years
            
        except Exception as e:
            print(f"Could not load 入力値 sheet: {e}")

    def show_parameter_editor(self):
        """Show parameter editor"""
        editor = tk.Toplevel(self.root)
        editor.title("Edit Parameters")
        editor.geometry("550x500")
        editor.grab_set()
        editor.configure(bg="#f0f0f0")
        
        main_frame = tk.Frame(editor, bg="#f0f0f0", padx=25, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        tk.Label(main_frame, text="Edit Input Parameters", 
                font=("Arial", 14, "bold"), fg="#1565C0", bg="#f0f0f0").pack(pady=(0, 20))
        
        # Parameter fields
        fields_frame = tk.Frame(main_frame, bg="#f0f0f0")
        fields_frame.pack(fill="x", pady=(0, 20))
        
        # データ個数
        tk.Label(fields_frame, text="データ個数:", font=("Arial", 11), 
                bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5, padx=(0, 10))
        self.data_count_var = tk.StringVar(value=str(self.nyuuryoku_params['data_count']))
        tk.Entry(fields_frame, textvariable=self.data_count_var, width=15).grid(row=0, column=1, pady=5)
        
        # 予測年数
        tk.Label(fields_frame, text="予測年数:", font=("Arial", 11), 
                bg="#f0f0f0").grid(row=1, column=0, sticky="w", pady=5, padx=(0, 10))
        self.prediction_years_var = tk.StringVar(value=str(self.nyuuryoku_params['prediction_years']))
        tk.Entry(fields_frame, textvariable=self.prediction_years_var, width=15).grid(row=1, column=1, pady=5)
        
        # λ定数
        tk.Label(fields_frame, text="λ定数:", font=("Arial", 11), 
                bg="#f0f0f0").grid(row=2, column=0, sticky="w", pady=5, padx=(0, 10))
        self.lambda_var = tk.StringVar(value=str(self.nyuuryoku_params['lambda_constant']))
        tk.Entry(fields_frame, textvariable=self.lambda_var, width=15).grid(row=2, column=1, pady=5)
        
        # Years
        tk.Label(main_frame, text="点検年度に対応した年:", font=("Arial", 11, "bold"), 
                bg="#f0f0f0").pack(anchor="w", pady=(10, 5))
        
        years_frame = tk.Frame(main_frame, bg="#f0f0f0")
        years_frame.pack(fill="x", pady=(0, 20))
        
        self.years_entry = tk.Text(years_frame, height=4, width=50, font=("Arial", 10))
        years_scrollbar = ttk.Scrollbar(years_frame, orient="vertical", command=self.years_entry.yview)
        self.years_entry.configure(yscrollcommand=years_scrollbar.set)
        
        current_years = ' '.join(map(str, self.nyuuryoku_params['inspection_years']))
        self.years_entry.insert("1.0", current_years)
        
        self.years_entry.pack(side="left", fill="x", expand=True)
        years_scrollbar.pack(side="right", fill="y")
        
                # Buttons
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill="x", pady=20)
        
        def save_and_generate():
            if self.validate_and_save_params():
                editor.destroy()
                self.generate_obser_files()
        
        tk.Button(button_frame, text="Save & Generate", command=save_and_generate,
                 bg="#4CAF50", fg="white", width=20, height=2, 
                 font=("Arial", 11)).pack(side="left", padx=5)
        
        tk.Button(button_frame, text="Cancel", command=editor.destroy,
                 bg="#f44336", fg="white", width=15, height=2, 
                 font=("Arial", 11)).pack(side="left", padx=5)

    def validate_and_save_params(self):
        """Validate and save parameters"""
        try:
            # Validate inputs
            data_count = int(self.data_count_var.get())
            pred_years = int(self.prediction_years_var.get())
            lambda_const = float(self.lambda_var.get())
            
            if data_count <= 0 or pred_years <= 0 or lambda_const <= 0:
                messagebox.showerror("Error", "All values must be positive")
                return False
            
            # Parse years
            years_text = self.years_entry.get("1.0", tk.END).strip()
            years = []
            for year_str in years_text.split():
                try:
                    year = int(year_str)
                    if 20 <= year <= 50:
                        years.append(year)
                except ValueError:
                    pass
            
            if not years:
                messagebox.showerror("Error", "Please enter valid years (20-50 range)")
                return False
            
            # Update parameters
            self.nyuuryoku_params.update({
                'data_count': data_count,
                'prediction_years': pred_years,
                'lambda_constant': lambda_const,
                'inspection_years': years
            })
            
            # Save to Excel
            self.save_nyuuryoku_parameters()
            return True
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Error saving parameters: {str(e)}")
            return False

    def save_nyuuryoku_parameters(self):
        """Save parameters to 入力値 sheet"""
        try:
            wb = load_workbook(self.workbook_path)
            
            if '入力値' in wb.sheetnames:
                wb.remove(wb['入力値'])
            
            ws = wb.create_sheet('入力値')
            
            # Headers in row 1
            ws['A1'] = 'データ個数'
            ws['B1'] = '予測年数' 
            ws['C1'] = 'λ定数'
            ws['D1'] = '点検年度に対応した年'
            
            # Values in row 2
            ws['A2'] = self.nyuuryoku_params['data_count']
            ws['B2'] = self.nyuuryoku_params['prediction_years']
            ws['C2'] = self.nyuuryoku_params['lambda_constant']
            
            # Years in column D starting from row 2
            for i, year in enumerate(self.nyuuryoku_params['inspection_years']):
                ws[f'D{i+2}'] = year
            
            wb.save(self.workbook_path)
            wb.close()
            
        except Exception as e:
            raise Exception(f"Error saving to Excel: {str(e)}")

    def generate_obser_files(self):
        """Generate all obser files"""
        try:
            output_directory = os.path.dirname(self.workbook_path)
            generated_files = []
            
            # Progress window
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Generating Files")
            progress_window.geometry("400x150")
            progress_window.grab_set()
            progress_window.configure(bg="#f0f0f0")
            
            progress_frame = tk.Frame(progress_window, bg="#f0f0f0", padx=20, pady=20)
            progress_frame.pack(fill="both", expand=True)
            
            status_label = tk.Label(progress_frame, text="Generating obser files...", 
                                  font=("Arial", 11), bg="#f0f0f0")
            status_label.pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_frame, mode='determinate', maximum=8)
            progress_bar.pack(fill="x", pady=10)
            
            # Generate files
            for i, (obser_file, sheet_name) in enumerate(self.sheet_mappings.items(), 1):
                status_label.config(text=f"Generating {obser_file}...")
                progress_window.update()
                
                try:
                    output_path = os.path.join(output_directory, obser_file)
                    self.create_obser_file(sheet_name, output_path)
                    generated_files.append(obser_file)
                except Exception as e:
                    print(f"Error generating {obser_file}: {e}")
                
                progress_bar['value'] = i
                progress_window.update()
            
            progress_window.destroy()
            
            # Show completion
            completion_msg = f"✅ Generated {len(generated_files)} files:\n\n"
            completion_msg += "\n".join(f"• {file}" for file in generated_files)
            completion_msg += f"\n\n📁 Location: {output_directory}"
            
            messagebox.showinfo("Generation Complete", completion_msg)
            
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("Error", f"Error generating files: {str(e)}")

    def create_obser_file(self, sheet_name, output_path):
        """Create obser file with sorting and 0 value replacement"""
        try:
            # Load sheet data
            sheet_df = pd.read_excel(self.workbook_path, sheet_name=sheet_name)
            
            # Sort by last column in descending order (matching VBA logic)
            if len(sheet_df) > 0 and len(sheet_df.columns) > 0:
                last_col = sheet_df.columns[-1]
                sheet_df = sheet_df.sort_values(by=last_col, ascending=False)
                
                # Save the sorted data back to Excel sheet
                with pd.ExcelWriter(self.workbook_path, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                # First line: parameters separated by spaces
                f.write(f"{self.nyuuryoku_params['data_count']} {self.nyuuryoku_params['prediction_years']} {self.nyuuryoku_params['lambda_constant']}\n")
                
                # Second line: years separated by single spaces
                years_line = ' '.join(map(str, self.nyuuryoku_params['inspection_years']))
                f.write(f"{years_line}\n")
                
                # Third line: blank line
                f.write("\n")
                
                # Find 構造物番号 column and export from there onwards
                kozo_col_idx = None
                for i, col in enumerate(sheet_df.columns):
                    if '構造物番号' in str(col):
                        kozo_col_idx = i
                        break
                
                if kozo_col_idx is None:
                    raise Exception(f"構造物番号 column not found in {sheet_name}")
                
                # Get all columns from 構造物番号 onwards (including 構造物番号)
                columns_to_export = sheet_df.columns[kozo_col_idx:]
                
                # Write data rows (tab-separated)
                for _, row in sheet_df.iterrows():
                    row_data = []
                    for col in columns_to_export:
                        value = row[col]
                        
                        # Handle empty/NaN values
                        if pd.isna(value) or value == '':
                            row_data.append('')
                        else:
                            try:
                                numeric_val = float(value)
                                # Replace 0 with 0.1 (matching VBA logic)
                                if numeric_val == 0:
                                    row_data.append('0.1')
                                elif numeric_val == int(numeric_val):
                                    row_data.append(str(int(numeric_val)))
                                else:
                                    row_data.append(str(round(numeric_val, 3)))
                            except (ValueError, TypeError):
                                # Non-numeric values (like 構造物番号)
                                if str(value) == '0':
                                    row_data.append('0.1')
                                else:
                                    row_data.append(str(value))
                    
                    f.write('\t'.join(row_data) + '\n')
            
        except Exception as e:
            raise Exception(f"Error creating {output_path}: {str(e)}")
        

    def run(self):
        """Run the application"""
        self.root.mainloop()


if __name__ == "__main__":
    print("Obser Files Generator")
    print("===================")
    print("Sheet Mappings:")
    sheet_mappings = {
        'obser1.txt': '割算結果(補修考慮)',
        'obser2.txt': '割算結果(補修無視)', 
        'obser3.txt': '補修無視',
        'obser4.txt': '補修考慮',
        'obser5.txt': '新しい演算(補修無視)',
        'obser6.txt': '新しい演算(補修考慮)',
        'obser7.txt': '割算結果-新しい演算(補修無視)',
        'obser8.txt': '割算結果-新しい演算(補修考慮)'
    }
    
    for obser_file, sheet_name in sheet_mappings.items():
        print(f"• {obser_file} ← {sheet_name}")
    print("===================")
    
    app = ObserFileGeneratorApp()
    app.run()
        