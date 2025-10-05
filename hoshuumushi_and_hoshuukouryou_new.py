import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog
import os
import re
import threading
import time

class SimpleProcessorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Processor")
        self.root.geometry("450x250")
        self.root.resizable(False, False)
        
        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.root.winfo_screenheight() // 2) - (250 // 2)
        self.root.geometry(f"450x250+{x}+{y}")
        
        self.workbook_path = None
        self.grouped_df = None
        self.structure_df = None
        
        self.create_gui()
    
    def create_gui(self):
        main_frame = tk.Frame(self.root, padx=30, pady=30)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Excel Processor", 
                              font=("Arial", 16, "bold"), fg="navy")
        title_label.pack(pady=(0, 30))
        
        # Select file button
        self.select_btn = tk.Button(main_frame, text="Select Excel File", 
                                   command=self.select_and_process, 
                                   bg="#4CAF50", fg="white", 
                                   width=20, height=2, font=("Arial", 12))
        self.select_btn.pack()
        
        # Progress frame (initially hidden)
        self.progress_frame = tk.Frame(main_frame)
        
        self.status_label = tk.Label(self.progress_frame, text="", 
                                    font=("Arial", 12, "bold"), fg="blue")
        self.status_label.pack(pady=(20, 5))
        
        self.detail_label = tk.Label(self.progress_frame, text="", 
                                    font=("Arial", 10), fg="gray")
        self.detail_label.pack(pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, 
                                          length=350, mode='determinate')
        self.progress_bar.pack(pady=(0, 10))
        
        self.step_label = tk.Label(self.progress_frame, text="", 
                                  font=("Arial", 11, "bold"), fg="green")
        self.step_label.pack()

    def select_and_process(self):
        self.workbook_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if self.workbook_path:
            self.select_btn.pack_forget()
            self.progress_frame.pack(fill="both", expand=True)
            
            # Start processing in separate thread
            threading.Thread(target=self.process_file, daemon=True).start()

    def process_file(self):
        try:
            # Step 1: Loading file
            self.root.after(0, self.update_progress, 
                          "ファイル読み込み中...", 
                          "Excelファイルを開いています", 
                          10, "")
            time.sleep(0.5)  # Brief pause for visual feedback
            
            # Load data
            self.grouped_df = pd.read_excel(self.workbook_path, sheet_name='グループ化点検履歴')
            
            # Try to load structure data
            try:
                self.structure_df = pd.read_excel(self.workbook_path, sheet_name='構造物番号')
            except:
                self.structure_df = None
            
            # Find year columns
            year_columns = self.find_year_result_columns()
            
            # Step 2: Process 補修無視 (1/2)
            self.root.after(0, self.update_progress, 
                          "補修無視シート生成中...", 
                          "最大値関数を適用してデータを処理中", 
                          35, "1/2 完了")
            time.sleep(0.3)
            max_result_df = self.apply_max_function_enhanced(year_columns)
            
            # Step 3: Process 補修考慮 (2/2)  
            self.root.after(0, self.update_progress, 
                          "補修考慮シート生成中...", 
                          "補修ロジックを適用してデータを処理中", 
                          70, "2/2 完了")
            time.sleep(0.3)
            hoshuu_result_df = self.apply_hoshuu_kouryou_enhanced(year_columns)
            
            # Step 4: Save results
            self.root.after(0, self.update_progress, 
                          "ファイル保存中...", 
                          "補修無視・補修考慮シートをExcelに保存中", 
                          90, "")
            time.sleep(0.3)
            self.save_results(max_result_df, hoshuu_result_df)
            
            # Step 5: Complete
            self.root.after(0, self.update_progress, 
                          "処理完了！", 
                          "補修無視・補修考慮シートが正常に生成されました", 
                          100, "✅ 完了")
            time.sleep(1.5)
            self.root.after(0, self.close_app)
            
        except Exception as e:
            self.root.after(0, self.update_progress, 
                          "エラーが発生しました", 
                          f"処理中にエラー: {str(e)}", 
                          0, "❌ エラー")
            time.sleep(3)
            self.root.after(0, self.close_app)

    def update_progress(self, status, detail, progress, step):
        self.status_label.config(text=status)
        self.detail_label.config(text=detail)
        self.progress_bar['value'] = progress
        self.step_label.config(text=step)
        self.root.update()

    def close_app(self):
        self.root.quit()
        self.root.destroy()

    def abbreviate_sen_name(self, sen_name):
        if pd.isna(sen_name) or sen_name == '':
            return ''
        
        sen_name = str(sen_name).strip()
        abbreviation_map = {
            "東急多摩川線": "TM", "多摩川線": "TM", "東横線": "TY",
            "大井町線": "OM", "池上線": "IK", "田園都市線": "DT",
            "目黒線": "MG", "こどもの国線": "KD", "世田谷線": "SG"
        }
        return abbreviation_map.get(sen_name, sen_name)
    
    def lookup_structure_number(self, structure_df, rosen_name, kozo_name, ekikan):
        try:
            if structure_df is None or len(structure_df) == 0:
                return ''
                
            rosen_name = str(rosen_name).strip() if pd.notna(rosen_name) else ''
            
            # Match by structure name
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
            
            # Match by station interval
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
        except:
            return ''
    
    def add_enhanced_columns(self, df, structure_df=None):
        enhanced_df = df.copy()
        
        # Add 路線名略称
        if '路線名' in enhanced_df.columns:
            enhanced_df['路線名略称'] = enhanced_df['路線名'].apply(self.abbreviate_sen_name)
        else:
            enhanced_df['路線名略称'] = ''
        
        # Add 構造物番号
        enhanced_df['構造物番号'] = ''
        
        if structure_df is not None:
            for index, row in enhanced_df.iterrows():
                rosen_name = row.get('路線名', '')
                kozo_name = row.get('構造物名称', '')
                
                ekikan = ''
                if row.get('駅（始）', '') and row.get('駅（至）', ''):
                    ekikan = f"{row.get('駅（始）', '')}→{row.get('駅（至）', '')}"
                
                bangou = self.lookup_structure_number(structure_df, rosen_name, kozo_name, ekikan)
                enhanced_df.at[index, '構造物番号'] = bangou
        
        return enhanced_df
    
    def reorder_columns_enhanced(self, df):
        priority_columns = [
            'グループ化キー', 'グループ化方法', '種別', '構造物名称',
            '駅（始）', '駅（至）', '点検区分1', 'データ件数',
            '路線名', '路線名略称', '構造物番号'
        ]
        
        # Get year columns
        year_columns = []
        for col in df.columns:
            if str(col).endswith('結果') or any(year in str(col) for year in ['2018', '2019', '2020', '2021', '2022', '2023', '2024']):
                year_columns.append(col)
        
        # Sort year columns
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
        for col in priority_columns:
            if col in df.columns:
                final_columns.append(col)
        
        final_columns.extend(year_columns)
        
        remaining_columns = [col for col in df.columns if col not in final_columns]
        final_columns.extend(remaining_columns)
        
        return df[final_columns]

    def find_year_result_columns(self):
        year_columns = []
        year_pattern = re.compile(r'(20\d{2})\s*結果')
        
        for col in self.grouped_df.columns:
            if str(col).endswith('結果'):
                match = year_pattern.search(str(col))
                if match:
                    year = int(match.group(1))
                    year_columns.append((year, col))
        
        year_columns.sort(key=lambda x: x[0])
        return [col for year, col in year_columns]

    def apply_max_function_enhanced(self, year_columns):
        result_df = self.grouped_df.copy()
        
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
        
        enhanced_df = self.add_enhanced_columns(result_df, self.structure_df)
        return self.reorder_columns_enhanced(enhanced_df)

    def apply_hoshuu_kouryou_enhanced(self, year_columns):
        result_df = self.grouped_df.copy()
        
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
        
        enhanced_df = self.add_enhanced_columns(result_df, self.structure_df)
        return self.reorder_columns_enhanced(enhanced_df)

    def save_results(self, max_result_df, hoshuu_result_df):
        with pd.ExcelWriter(self.workbook_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            max_result_df.to_excel(writer, sheet_name="補修無視", index=False)
            hoshuu_result_df.to_excel(writer, sheet_name="補修考慮", index=False)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = SimpleProcessorApp()
    app.run()