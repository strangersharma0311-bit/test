import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
from collections import defaultdict

class EnhancedDataGroupingApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Enhanced Data Grouping System")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # Center the window on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")
        
        self.workbook_path = None
        self.rules_file = None
        self.rules = []
        self.enzan_kekka_df = None
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
    
    def create_enhanced_grouping_key(self, shubetsu, tenken_kubun, structure_name, eki_start, eki_end, group_method):
        """Create grouping key with enhanced logic - don't show 点検区分1 when All"""
        try:
            if group_method == "構造物名称":
                if tenken_kubun == "*":
                    # Don't include tenken_kubun in key when it's "*" (All)
                    key = f"{shubetsu}|{structure_name}"
                else:
                    key = f"{shubetsu}|{structure_name}|{tenken_kubun}"
            else:  # 駅間
                ekikan = f"{eki_start}→{eki_end}" if eki_start and eki_end else ""
                if tenken_kubun == "*":
                    # Don't include tenken_kubun in key when it's "*" (All)
                    key = f"{shubetsu}|{ekikan}"
                else:
                    key = f"{shubetsu}|{ekikan}|{tenken_kubun}"
            
            return key
            
        except Exception as e:
            print(f"Error creating grouping key: {e}")
            return "UNKNOWN"
    
    def create_main_gui(self):
        """Create main GUI for file selection"""
        main_frame = tk.Frame(self.root, padx=60, pady=60)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Enhanced Data Grouping System", 
                              font=("Arial", 24, "bold"), fg="navy")
        title_label.pack(pady=(0, 40))
        
        # Instructions
        instruction_text = ("Enhanced features:\n"
                          "• Adds 構造物番号 and 路線名略称 columns\n"
                          "• Improved column positioning\n"
                          "• Smart grouping key display\n"
                          "• Better rule management with 'All' option\n\n"
                          "Select Excel workbook with '演算結果' sheet")
        instruction_label = tk.Label(main_frame, text=instruction_text, 
                                   font=("Arial", 14), justify="center",
                                   wraplength=600)
        instruction_label.pack(pady=(0, 40))
        
        # Status label for feedback
        self.status_label = tk.Label(main_frame, text="Ready to select file...", 
                                    font=("Arial", 12), fg="gray")
        self.status_label.pack(pady=(0, 30))
        
        # Select file button
        select_btn = tk.Button(main_frame, text="📁 Browse & Select Excel File", 
                             command=self.select_workbook_with_feedback, 
                             bg="#4CAF50", fg="white", 
                             width=25, height=2, font=("Arial", 12, "bold"),
                             cursor="hand2")
        select_btn.pack(pady=20)
        
        # Exit button
        exit_frame = tk.Frame(main_frame)
        exit_frame.pack(pady=(30, 0))
        
        exit_btn = tk.Button(exit_frame, text="❌ Exit Application", 
                           command=self.confirm_exit, bg="#f44336", fg="white", 
                           width=20, height=2, font=("Arial", 10),
                           cursor="hand2")
        exit_btn.pack()
    
    def select_workbook_with_feedback(self):
        """Select workbook with user feedback"""
        self.status_label.config(text="Opening file browser...", fg="blue")
        self.root.update()
        
        self.workbook_path = filedialog.askopenfilename(
            title="Select Excel Workbook with '演算結果' Sheet",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=os.path.expanduser("~")
        )
        
        if not self.workbook_path:
            self.status_label.config(text="No file selected. Please try again.", fg="orange")
            return
        
        self.status_label.config(text="🔄 Loading and validating file...", fg="blue")
        self.root.update()
        
        self.root.after(100, self.validate_workbook)

    def validate_workbook(self):
        """Validate workbook with progress feedback"""
        try:
            if not os.path.exists(self.workbook_path):
                raise Exception("File not found")
            
            self.status_label.config(text="🔍 Checking required sheets...", fg="blue")
            self.root.update()
            
            wb = load_workbook(self.workbook_path)
            required_sheets = ['演算結果']
            missing_sheets = [sheet for sheet in required_sheets if sheet not in wb.sheetnames]
            
            if missing_sheets:
                self.status_label.config(text="❌ Required sheet not found!", fg="red")
                messagebox.showerror("Missing Sheet", 
                                   f"Required sheet not found: {', '.join(missing_sheets)}")
                self.status_label.config(text="Ready to select file...", fg="gray")
                return
            
            self.status_label.config(text="📊 Loading calculation data...", fg="blue")
            self.root.update()
            
            self.enzan_kekka_df = pd.read_excel(self.workbook_path, sheet_name='演算結果')
            
            if len(self.enzan_kekka_df) == 0:
                raise Exception("The calculation results sheet is empty")
            
            # Try to load structure data if it exists (for enhanced features)
            try:
                self.structure_df = pd.read_excel(self.workbook_path, sheet_name='構造物番号')
                print("Found 構造物番号 sheet - enhanced features enabled")
            except:
                self.structure_df = None
                print("No 構造物番号 sheet found - basic features only")
            
            self.status_label.config(text="⚙️ Setting up enhanced grouping system...", fg="blue")
            self.root.update()
            
            self.rules_file = os.path.join(os.path.dirname(self.workbook_path), "grouping_rules.json")
            self.rules = self.load_rules()
            
            self.status_label.config(text="✅ File loaded successfully!", fg="green")
            self.root.update()
            
            enhancement_status = "with 構造物番号 enhancements" if self.structure_df is not None else "basic version"
            messagebox.showinfo("Success", 
                               f"✅ Excel file loaded successfully {enhancement_status}!\n\n"
                               f"📁 File: {os.path.basename(self.workbook_path)}\n"
                               f"📊 Records: {len(self.enzan_kekka_df):,}\n\n"
                               f"Proceeding to enhanced grouping rules management...")
            
            self.root.withdraw()
            self.show_grouping_manager()
            
        except Exception as e:
            self.status_label.config(text="❌ Error loading file", fg="red")
            messagebox.showerror("Error", f"Failed to load Excel file:\n\n{str(e)}")
            self.status_label.config(text="Ready to select file...", fg="gray")

    def confirm_exit(self):
        """Confirm before exiting"""
        if messagebox.askyesno("Exit Application", 
                              "Are you sure you want to exit?\n\n"
                              "This will close the Enhanced Data Grouping System completely."):
            self.root.quit()
    
    def load_rules(self):
        """Load existing rules from JSON file"""
        default_rules = [
            {
                "shubetsu": "停車場",
                "tenken_kubun": "*",
                "group_by": "構造物名称",
                "description": "Station grouped by structure name - All inspection categories"
            },
            {
                "shubetsu": "擁壁・法面",
                "tenken_kubun": "*",
                "group_by": "駅間",
                "description": "Retaining walls grouped by station interval - All inspection categories"
            },
            {
                "shubetsu": "線路設備",
                "tenken_kubun": "*",
                "group_by": "駅間", 
                "description": "Track equipment grouped by station interval - All inspection categories"
            },
            {
                "shubetsu": "トンネル",
                "tenken_kubun": "*",
                "group_by": "構造物名称",
                "description": "Tunnels grouped by structure name - All inspection categories"
            },
            {
                "shubetsu": "高架橋",
                "tenken_kubun": "*",
                "group_by": "構造物名称",
                "description": "Elevated bridges grouped by structure name - All inspection categories"
            }
        ]
        
        if os.path.exists(self.rules_file):
            try:
                with open(self.rules_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return default_rules
        else:
            self.save_rules(default_rules)
            return default_rules
    
    def save_rules(self, rules=None):
        """Save rules to JSON file"""
        if rules is None:
            rules = self.rules
        with open(self.rules_file, 'w', encoding='utf-8') as f:
            json.dump(rules, f, ensure_ascii=False, indent=2)
    
    def show_grouping_manager(self):
        """Show enhanced grouping rules management window"""
        self.main_window = tk.Toplevel()
        self.main_window.title("Enhanced Data Grouping Rules Management")
        self.main_window.geometry("1400x900")
        self.main_window.minsize(1200, 700)
        self.main_window.grab_set()
        self.main_window.resizable(True, True)
        
        def on_closing():
            if messagebox.askyesno("Close Application", 
                                  "This will close the entire Enhanced Data Grouping System.\n"
                                  "Are you sure?"):
                self.main_window.destroy()
                self.root.quit()
        
        self.main_window.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Center window
        self.main_window.update_idletasks()
        x = (self.main_window.winfo_screenwidth() // 2) - (1200 // 2)
        y = (self.main_window.winfo_screenheight() // 2) - (800 // 2)
        self.main_window.geometry(f"1200x800+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(self.main_window, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Title and file info
        title_frame = tk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 20))
        
        title_label = tk.Label(title_frame, text="Enhanced Data Grouping Rules Management", 
                              font=("Arial", 12, "bold"), fg="navy")
        title_label.pack(anchor="w")
        
        file_label = tk.Label(title_frame, text=f"File: {os.path.basename(self.workbook_path)}", 
                             font=("Arial", 8), fg="gray")
        file_label.pack(anchor="w")
        
        data_info_label = tk.Label(title_frame, text=f"Data Count: {len(self.enzan_kekka_df):,} records", 
                                  font=("Arial", 10), fg="blue")
        data_info_label.pack(anchor="w")
        
        # Enhancement status
        enhancement_text = "✅ Enhanced with 構造物番号 & 路線名略称 columns" if self.structure_df is not None else "⚠️ Basic version (no 構造物番号 sheet found)"
        enhancement_label = tk.Label(title_frame, text=enhancement_text, 
                                   font=("Arial", 9), fg="green" if self.structure_df is not None else "orange")
        enhancement_label.pack(anchor="w")
        
        # Instructions
        instruction_text = ("Enhanced Data Grouping Features:\n"
                          "• Smart grouping with improved column layout\n"
                          "• データ件数 → 路線名 → 路線名略称 → 構造物番号 → Year columns\n"
                          "• Cleaner grouping keys (no 点検区分1 when 'All' selected)\n"
                          "• Auto-lookup 構造物番号 from structure data")
        instruction_label = tk.Label(main_frame, text=instruction_text, 
                                   font=("Arial", 11), justify="left", wraplength=900)
        instruction_label.pack(pady=(0, 20))
        
        # Rules display frame
        rules_frame = tk.LabelFrame(main_frame, text="Registered Rules", 
                                   font=("Arial", 12, "bold"), padx=15, pady=15)
        rules_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        # Create treeview for rules display
        columns = ("No", "Type", "Inspection Category", "Grouping Method", "Description")
        self.rules_tree = ttk.Treeview(rules_frame, columns=columns, show="headings", height=12)
        
        # Define column headings and widths
        column_widths = {"No": 50, "Type": 150, "Inspection Category": 200, 
                        "Grouping Method": 150, "Description": 350}
        
        for col in columns:
            self.rules_tree.heading(col, text=col)
            self.rules_tree.column(col, width=column_widths.get(col, 100))
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(rules_frame, orient="vertical", command=self.rules_tree.yview)
        self.rules_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        self.rules_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate rules
        self.refresh_rules_display()
        
        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=20)
        
        # Rule management buttons
        edit_btn = tk.Button(button_frame, text="✏️ Edit Selected Rule", 
                           command=self.edit_selected_rule, bg="#2196F3", fg="white", 
                           width=18, height=2, font=("Arial", 11), cursor="hand2")
        edit_btn.pack(side="left", padx=5)
        
        add_btn = tk.Button(button_frame, text="➕ Add New Rule", 
                          command=self.add_new_rule, bg="#4CAF50", fg="white", 
                          width=15, height=2, font=("Arial", 11), cursor="hand2")
        add_btn.pack(side="left", padx=5)
        
        delete_btn = tk.Button(button_frame, text="🗑️ Delete Rule", 
                             command=self.delete_selected_rule, bg="#f44336", fg="white", 
                             width=15, height=2, font=("Arial", 11), cursor="hand2")
        delete_btn.pack(side="left", padx=5)
        
        # Main action buttons
        action_frame = tk.Frame(main_frame)
        action_frame.pack(fill="x", pady=(20, 0))
        
        continue_btn = tk.Button(action_frame, text="🚀 Start Enhanced Grouping", 
                               command=self.start_enhanced_grouping_process, bg="#FF9800", fg="white", 
                               width=25, height=2, font=("Arial", 12, "bold"), cursor="hand2")
        continue_btn.pack(side="right", padx=15)
        
        back_btn = tk.Button(action_frame, text="⬅️ Back to File Selection", 
                           command=self.back_to_file_selection, bg="#9E9E9E", fg="white", 
                           width=20, height=2, font=("Arial", 11), cursor="hand2")
        back_btn.pack(side="right", padx=15)
    
    def refresh_rules_display(self):
        """Refresh the rules display in treeview"""
        # Clear existing items
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)
        
        # Add rules to treeview
        for i, rule in enumerate(self.rules, 1):
            tenken_display = "All" if rule["tenken_kubun"] == "*" else rule["tenken_kubun"]
            group_by_display = "Structure Name" if rule["group_by"] == "構造物名称" else "Station Interval"
            
            self.rules_tree.insert("", "end", values=(
                i, 
                rule["shubetsu"], 
                tenken_display,
                group_by_display,
                rule.get("description", "")
            ))
    
    def edit_selected_rule(self):
        """Edit the selected rule"""
        selection = self.rules_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a rule to edit.")
            return
        
        item = self.rules_tree.item(selection[0])
        rule_index = int(item['values'][0]) - 1
        self.show_rule_edit_dialog(rule_index)
    
    def add_new_rule(self):
        """Add a new rule"""
        self.show_rule_edit_dialog(-1)
    
    def delete_selected_rule(self):
        """Delete the selected rule"""
        selection = self.rules_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a rule to delete.")
            return
        
        item = self.rules_tree.item(selection[0])
        rule_name = f"{item['values'][1]} - {item['values'][2]}"
        
        if messagebox.askyesno("Confirm Deletion", 
                              f"Delete this rule?\n\n"
                              f"Rule: {rule_name}\n\n"
                              f"This action cannot be undone."):
            rule_index = int(item['values'][0]) - 1
            del self.rules[rule_index]
            self.save_rules()
            self.refresh_rules_display()
            
            messagebox.showinfo("Success", f"Rule deleted successfully:\n{rule_name}")
    
    def show_rule_edit_dialog(self, rule_index):
        """Show dialog for editing/adding rules with 'All' option"""
        edit_window = tk.Toplevel(self.main_window)
        is_new = rule_index == -1
        title = "Add New Rule" if is_new else "Edit Rule"
        edit_window.title(title)
        edit_window.geometry("600x500")
        edit_window.grab_set()
        edit_window.resizable(False, False)
        edit_window.transient(self.main_window)
        
        main_frame = tk.Frame(edit_window, padx=25, pady=25)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text=title, font=("Arial", 14, "bold"), fg="navy")
        title_label.pack(pady=(0, 20))
        
        # Get current rule data
        current_rule = self.rules[rule_index] if not is_new else {
            "shubetsu": "", "tenken_kubun": "*", "group_by": "構造物名称", "description": ""
        }
        
        # Form fields
        fields_frame = tk.Frame(main_frame)
        fields_frame.pack(fill="x", pady=20)
        
        # Type field
        tk.Label(fields_frame, text="Structure Type:", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", pady=10)
        shubetsu_var = tk.StringVar(value=current_rule["shubetsu"])
        shubetsu_entry = ttk.Combobox(fields_frame, textvariable=shubetsu_var, width=30, font=("Arial", 10))
        
        if '種別' in self.enzan_kekka_df.columns:
            unique_shubetsu = sorted(self.enzan_kekka_df['種別'].dropna().unique().tolist())
            shubetsu_entry['values'] = unique_shubetsu
        
        shubetsu_entry.grid(row=0, column=1, sticky="w", padx=(10,0), pady=10)
        
        # Inspection Category field - with "All" option instead of "*"
        tk.Label(fields_frame, text="Inspection Category:", font=("Arial", 11, "bold")).grid(row=1, column=0, sticky="w", pady=10)
        tenken_var = tk.StringVar(value="All" if current_rule["tenken_kubun"] == "*" else current_rule["tenken_kubun"])
        tenken_entry = ttk.Combobox(fields_frame, textvariable=tenken_var, width=30, font=("Arial", 10))
        
        if '点検区分1' in self.enzan_kekka_df.columns:
            unique_tenken = sorted(self.enzan_kekka_df['点検区分1'].dropna().unique().tolist())
            unique_tenken.insert(0, "All")  # Use "All" instead of "*"
            tenken_entry['values'] = unique_tenken
        else:
            tenken_entry['values'] = ["All"]
        
        tenken_entry.grid(row=1, column=1, sticky="w", padx=(10,0), pady=10)
        
        # Help text
        help_label = tk.Label(fields_frame, text="※ \"All\" applies to all inspection categories", 
                             font=("Arial", 9), fg="gray")
        help_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0,10))
        
        # Grouping Method field
        tk.Label(fields_frame, text="Grouping Method:", font=("Arial", 11, "bold")).grid(row=3, column=0, sticky="w", pady=10)
        group_by_var = tk.StringVar(value=current_rule["group_by"])
        
        radio_frame = tk.Frame(fields_frame)
        radio_frame.grid(row=3, column=1, sticky="w", padx=(10,0), pady=10)
        
        structure_radio = tk.Radiobutton(radio_frame, text="Structure Name", 
                                       variable=group_by_var, value="構造物名称", 
                                       font=("Arial", 10))
        structure_radio.pack(anchor="w")
        
        station_radio = tk.Radiobutton(radio_frame, text="Station Interval", 
                                     variable=group_by_var, value="駅間", 
                                     font=("Arial", 10))
        station_radio.pack(anchor="w")
        
        # Description field
        tk.Label(fields_frame, text="Description:", font=("Arial", 11, "bold")).grid(row=4, column=0, sticky="w", pady=10)
        description_var = tk.StringVar(value=current_rule.get("description", ""))
        description_entry = tk.Entry(fields_frame, textvariable=description_var, width=40, font=("Arial", 10))
        description_entry.grid(row=4, column=1, sticky="w", padx=(10,0), pady=10)
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        def save_rule():
            shubetsu = shubetsu_var.get().strip()
            tenken = tenken_var.get().strip()
            group_by = group_by_var.get()
            description = description_var.get().strip()
            
            # Convert "All" back to "*" for internal storage
            if tenken == "All":
                tenken = "*"
                        
            if not shubetsu or not tenken:
                messagebox.showerror("Error", "Please enter both Structure Type and Inspection Category.")
                return
            
            new_rule = {
                "shubetsu": shubetsu,
                "tenken_kubun": tenken,
                "group_by": group_by,
                "description": description
            }
            
            if is_new:
                self.rules.append(new_rule)
            else:
                self.rules[rule_index] = new_rule
            
            self.save_rules()
            self.refresh_rules_display()
            
            edit_window.destroy()
            messagebox.showinfo("Success", "Rule saved successfully.")
        
        save_btn = tk.Button(button_frame, text="Save", command=save_rule,
                           bg="#4CAF50", fg="white", width=15, height=2, font=("Arial", 11))
        save_btn.pack(side="right", padx=10)
        
        cancel_btn = tk.Button(button_frame, text="Cancel", command=edit_window.destroy,
                             bg="#9E9E9E", fg="white", width=15, height=2, font=("Arial", 11))
        cancel_btn.pack(side="right", padx=10)
    
    def back_to_file_selection(self):
        """Go back to file selection"""
        self.main_window.destroy()
        self.root.deiconify()
    
    def start_enhanced_grouping_process(self):
        """Start the enhanced grouping process"""
        try:
            if '種別' not in self.enzan_kekka_df.columns or '点検区分1' not in self.enzan_kekka_df.columns:
                raise Exception("Required columns (Type, Inspection Category) not found in data")
            
            unique_combinations = self.enzan_kekka_df[['種別', '点検区分1']].drop_duplicates()
            missing_rules = []
            
            # Check for missing rules
            for _, row in unique_combinations.iterrows():
                shubetsu = str(row['種別']) if pd.notna(row['種別']) else ""
                tenken = str(row['点検区分1']) if pd.notna(row['点検区分1']) else ""
                
                if not self.find_matching_rule(shubetsu, tenken):
                    missing_rules.append((shubetsu, tenken))
            
            if missing_rules:
                self.show_missing_rules_dialog(missing_rules)
            else:
                self.perform_enhanced_data_grouping()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error during data processing:\n{str(e)}")
    
    def find_matching_rule(self, shubetsu, tenken_kubun):
        """Find matching rule for given shubetsu and tenken_kubun"""
        for rule in self.rules:
            if rule["shubetsu"] == shubetsu:
                if rule["tenken_kubun"] == "*" or rule["tenken_kubun"] == tenken_kubun:
                    return rule
        return None
    
    def show_missing_rules_dialog(self, missing_rules):
        """Show intelligent dialog for missing rules with enhanced options"""
        missing_window = tk.Toplevel(self.main_window)
        missing_window.title("Configure Enhanced Grouping Rules")
        missing_window.geometry("1000x700")
        missing_window.grab_set()
        missing_window.resizable(True, True)
        missing_window.transient(self.main_window)
        
        main_frame = tk.Frame(missing_window, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Configure Enhanced Grouping Rules", 
                            font=("Arial", 16, "bold"), fg="navy")
        title_label.pack(pady=(0, 10))
        
        subtitle_label = tk.Label(main_frame, 
                                text="Choose how to group each structure type with enhanced features", 
                                font=("Arial", 11))
        subtitle_label.pack(pady=(0, 20))
        
        # Group missing rules by 種別
        shubetsu_groups = defaultdict(list)
        for shubetsu, tenken in missing_rules:
            shubetsu_groups[shubetsu].append(tenken)
        
        # Scrollable frame for 種別 groups
        canvas = tk.Canvas(main_frame, height=400)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Store user choices
        self.shubetsu_choices = {}
        
        row_count = 0
        for shubetsu, tenken_list in shubetsu_groups.items():
            # Create frame for each 種別
            shubetsu_frame = tk.LabelFrame(scrollable_frame, 
                                        text=f"種別: {shubetsu}", 
                                        font=("Arial", 12, "bold"), 
                                        padx=15, pady=15,
                                        bg="lightblue")
            shubetsu_frame.grid(row=row_count, column=0, sticky="ew", padx=10, pady=10)
            scrollable_frame.grid_columnconfigure(0, weight=1)
            
            # Show what 点検区分1 values exist for this 種別
            tenken_info = tk.Label(shubetsu_frame, 
                                text=f"Found 点検区分1 values: {', '.join(tenken_list)}", 
                                font=("Arial", 10), fg="darkblue", wraplength=400)
            tenken_info.pack(anchor="w", pady=(0, 10))
            
            # Choice variable for this 種別
            choice_var = tk.StringVar(value="universal")
            self.shubetsu_choices[shubetsu] = {
                'choice_var': choice_var,
                'tenken_list': tenken_list,
                'group_method_var': tk.StringVar(value="構造物名称"),
                'individual_methods': {}
            }
            
            # Option 1: Universal rule (use "All")
            universal_frame = tk.Frame(shubetsu_frame)
            universal_frame.pack(fill="x", pady=5)
            
            universal_radio = tk.Radiobutton(universal_frame, 
                                        text="Create ONE rule for ALL inspection categories", 
                                        variable=choice_var, value="universal",
                                        font=("Arial", 10, "bold"), fg="green")
            universal_radio.pack(anchor="w")
            
            # Grouping method for universal choice
            universal_method_frame = tk.Frame(universal_frame)
            universal_method_frame.pack(fill="x", padx=20, pady=5)
            
            tk.Label(universal_method_frame, text="Group by:", font=("Arial", 9)).pack(side="left")
            
            struktur_radio = tk.Radiobutton(universal_method_frame, text="構造物名称", 
                                        variable=self.shubetsu_choices[shubetsu]['group_method_var'], 
                                        value="構造物名称", font=("Arial", 9))
            struktur_radio.pack(side="left", padx=10)
            
            ekikan_radio = tk.Radiobutton(universal_method_frame, text="駅間", 
                                        variable=self.shubetsu_choices[shubetsu]['group_method_var'], 
                                        value="駅間", font=("Arial", 9))
            ekikan_radio.pack(side="left", padx=10)
            
            # Option 2: Individual rules
            individual_frame = tk.Frame(shubetsu_frame)
            individual_frame.pack(fill="x", pady=5)
            
            individual_radio = tk.Radiobutton(individual_frame, 
                                            text="Create SEPARATE rules for each inspection category", 
                                            variable=choice_var, value="individual",
                                            font=("Arial", 10, "bold"), fg="orange")
            individual_radio.pack(anchor="w")
            
            # Individual grouping methods for each 点検区分1
            individual_details_frame = tk.Frame(individual_frame)
            individual_details_frame.pack(fill="x", padx=20, pady=5)
            
            for tenken in tenken_list:
                tenken_frame = tk.Frame(individual_details_frame)
                tenken_frame.pack(fill="x", pady=2)
                
                tk.Label(tenken_frame, text=f"  {tenken}:", font=("Arial", 9), width=20, anchor="w").pack(side="left")
                
                method_var = tk.StringVar(value="構造物名称")
                self.shubetsu_choices[shubetsu]['individual_methods'][tenken] = method_var
                
                struktur_radio2 = tk.Radiobutton(tenken_frame, text="構造物名称", 
                                            variable=method_var, value="構造物名称", font=("Arial", 8))
                struktur_radio2.pack(side="left", padx=5)
                
                ekikan_radio2 = tk.Radiobutton(tenken_frame, text="駅間", 
                                            variable=method_var, value="駅間", font=("Arial", 8))
                ekikan_radio2.pack(side="left", padx=5)
            
            row_count += 1
        
        # Apply smart rules function
        def apply_smart_rules():
            """Apply the smart rules based on user choices"""
            new_rules = []
            
            for shubetsu, choice_data in self.shubetsu_choices.items():
                choice = choice_data['choice_var'].get()
                
                if choice == "universal":
                    # Create one rule with "*" for all 点検区分1
                    group_method = choice_data['group_method_var'].get()
                    new_rule = {
                        "shubetsu": shubetsu,
                        "tenken_kubun": "*",
                        "group_by": group_method,
                        "description": f"Universal rule for {shubetsu} - all inspection categories"
                    }
                    new_rules.append(new_rule)
                    print(f"Created universal rule: {shubetsu} -> {group_method}")
                    
                elif choice == "individual":
                    # Create separate rules for each 点検区分1
                    for tenken in choice_data['tenken_list']:
                        group_method = choice_data['individual_methods'][tenken].get()
                        new_rule = {
                            "shubetsu": shubetsu,
                            "tenken_kubun": tenken,
                            "group_by": group_method,
                            "description": f"Individual rule for {shubetsu} - {tenken}"
                        }
                        new_rules.append(new_rule)
                        print(f"Created individual rule: {shubetsu}|{tenken} -> {group_method}")
            
            # Add new rules to existing ones
            self.rules.extend(new_rules)
            self.save_rules()
            
            missing_window.destroy()
            
            messagebox.showinfo("Success", 
                            f"{len(new_rules)} new enhanced rules created successfully!\n\n"
                            f"Starting enhanced grouping process...")
            
            self.perform_enhanced_data_grouping()
        
        # Bulk assignment functions
        def set_all_universal_struktur():
            """Set all to universal with 構造物名称"""
            for choice_data in self.shubetsu_choices.values():
                choice_data['choice_var'].set("universal")
                choice_data['group_method_var'].set("構造物名称")
            messagebox.showinfo("Applied", "All set to Universal + 構造物名称")
        
        def set_all_universal_ekikan():
            """Set all to universal with 駅間"""
            for choice_data in self.shubetsu_choices.values():
                choice_data['choice_var'].set("universal")
                choice_data['group_method_var'].set("駅間")
            messagebox.showinfo("Applied", "All set to Universal + 駅間")
        
        def set_all_individual():
            """Set all to individual rules"""
            for choice_data in self.shubetsu_choices.values():
                choice_data['choice_var'].set("individual")
            messagebox.showinfo("Applied", "All set to Individual Rules")
        
        # Bulk assignment buttons
        bulk_frame = tk.LabelFrame(main_frame, text="Quick Actions", font=("Arial", 11, "bold"), padx=15, pady=10)
        bulk_frame.pack(fill="x", pady=(0, 20))
        
        bulk_button_frame = tk.Frame(bulk_frame)
        bulk_button_frame.pack(fill="x")
        
        tk.Button(bulk_button_frame, text="All Universal (構造物名称)", 
                command=set_all_universal_struktur, bg="#4CAF50", fg="white", 
                width=20, font=("Arial", 9)).pack(side="left", padx=5)
        
        tk.Button(bulk_button_frame, text="All Universal (駅間)", 
              command=set_all_universal_ekikan, bg="#2196F3", fg="white", 
              width=18, font=("Arial", 9)).pack(side="left", padx=5)
    
        tk.Button(bulk_button_frame, text="All Individual", 
                command=set_all_individual, bg="#FF9800", fg="white", 
                width=15, font=("Arial", 9)).pack(side="left", padx=5)
        
        # Main action buttons
        action_frame = tk.Frame(main_frame)
        action_frame.pack(fill="x", pady=(20, 0))
        
        apply_btn = tk.Button(action_frame, text="Apply Rules & Start Enhanced Grouping", 
                            command=apply_smart_rules, bg="#4CAF50", fg="white", 
                            width=30, height=2, font=("Arial", 11, "bold"))
        apply_btn.pack(side="right", padx=15)
        
        back_btn = tk.Button(action_frame, text="Back", 
                        command=missing_window.destroy, 
                        bg="#9E9E9E", fg="white", width=15, height=2, font=("Arial", 11))
        back_btn.pack(side="right", padx=15)

    def perform_enhanced_data_grouping(self):
        """Perform the enhanced data grouping with new columns"""
        try:
            # Show progress dialog
            progress_window = tk.Toplevel(self.main_window)
            progress_window.title("Performing Enhanced Grouping")
            progress_window.geometry("450x150")
            progress_window.grab_set()
            progress_window.resizable(False, False)
            progress_window.transient(self.main_window)
            
            progress_frame = tk.Frame(progress_window, padx=20, pady=20)
            progress_frame.pack(fill="both", expand=True)
            
            status_label = tk.Label(progress_frame, text="Enhanced grouping in progress...", 
                                font=("Arial", 11))
            status_label.pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
            progress_bar.pack(fill="x", pady=10)
            progress_bar.start()
            
            # Process data
            self.root.after(100, lambda: self.execute_enhanced_grouping(progress_window, status_label))
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during enhanced grouping process:\n{str(e)}")

    def execute_enhanced_grouping(self, progress_window, status_label):
        """Execute the enhanced grouping logic with new columns"""
        try:
            status_label.config(text="Analyzing data with enhancements...")
            self.root.update()
            
            # Create a copy of the data for processing
            df = self.enzan_kekka_df.copy()
            
            # Add enhanced grouping key column
            status_label.config(text="Generating enhanced grouping keys...")
            self.root.update()
            
            grouping_keys = []
            grouping_methods = []
            
            for _, row in df.iterrows():
                shubetsu = str(row.get('種別', '')) if pd.notna(row.get('種別')) else ''
                tenken = str(row.get('点検区分1', '')) if pd.notna(row.get('点検区分1')) else ''
                
                # Find matching rule
                rule = self.find_matching_rule(shubetsu, tenken)
                
                if rule:
                    group_method = rule['group_by']
                    grouping_methods.append(group_method)
                    
                    if group_method == "構造物名称":
                        # Group by structure name
                        structure_name = str(row.get('構造物名称', '')) if pd.notna(row.get('構造物名称')) else ''
                        key = self.create_enhanced_grouping_key(shubetsu, rule['tenken_kubun'], structure_name, None, None, group_method)
                    else:  # 駅間
                        # Group by station interval
                        eki_start = str(row.get('駅（始）', '')) if pd.notna(row.get('駅（始）')) else ''
                        eki_end = str(row.get('駅（至）', '')) if pd.notna(row.get('駅（至）')) else ''
                        key = self.create_enhanced_grouping_key(shubetsu, rule['tenken_kubun'], None, eki_start, eki_end, group_method)
                    
                    grouping_keys.append(key)
                else:
                    # Fallback
                    grouping_keys.append(f"UNKNOWN|{shubetsu}|{tenken}")
                    grouping_methods.append("構造物名称")
            
            # Add grouping columns to dataframe
            df['Grouping Key'] = grouping_keys
            df['Grouping Method'] = grouping_methods
            
            status_label.config(text="Aggregating data with enhanced features...")
            self.root.update()
            
            # Get year columns (result columns)
            year_columns = [col for col in df.columns if col.endswith(' 結果') and any(year in col for year in ['2024', '2023', '2022', '2021', '2020', '2019', '2018'])]
            
            # Group data and aggregate with enhanced columns
            grouped_data = []
            
            for group_key in df['Grouping Key'].unique():
                group_df = df[df['Grouping Key'] == group_key]
                
                # Get basic info from first row
                first_row = group_df.iloc[0]
                
                # Create result row with enhanced column structure
                result_row = {
                    'グループ化キー': group_key,
                    'グループ化方法': first_row['Grouping Method'],
                    'データ件数': len(group_df),  # BEFORE 路線名
                    '路線名': first_row.get('路線名', ''),  # stays
                    '路線名略称': self.abbreviate_sen_name(first_row.get('路線名', '')),  # after 路線名
                    '構造物番号': '',  # after 路線名略称 - will be filled later
                    '種別': first_row.get('種別', ''),
                    '構造物名称': first_row.get('構造物名称', '') if first_row['Grouping Method'] == "構造物名称" else '',
                    '駅（始）': first_row.get('駅（始）', '') if first_row['Grouping Method'] == "駅間" else '',
                    '駅（至）': first_row.get('駅（至）', '') if first_row['Grouping Method'] == "駅間" else '',
                    '点検区分1': first_row.get('点検区分1', '')
                }
                
                # Add 構造物番号 lookup if structure_df is available
                if self.structure_df is not None:
                    rosen_name = result_row['路線名']
                    kozo_name = result_row['構造物名称']
                    
                    # Create ekikan for lookup
                    ekikan = ''
                    if result_row['駅（始）'] and result_row['駅（至）']:
                        ekikan = f"{result_row['駅（始）']}→{result_row['駅（至）']}"
                    
                    # Lookup structure number
                    bangou = self.lookup_structure_number(self.structure_df, rosen_name, kozo_name, ekikan)
                    result_row['構造物番号'] = bangou
                
                # Aggregate year results
                for year_col in year_columns:
                    # Sum non-empty values
                    values = group_df[year_col].dropna()
                    values = values[values != '']
                    
                    if len(values) > 0:
                        try:
                            numeric_values = pd.to_numeric(values, errors='coerce').dropna()
                            if len(numeric_values) > 0:
                                result_row[year_col] = numeric_values.sum()
                            else:
                                result_row[year_col] = ''
                        except:
                            result_row[year_col] = ''
                    else:
                        result_row[year_col] = ''
                
                grouped_data.append(result_row)
            
            # Create enhanced grouped dataframe
            grouped_df = pd.DataFrame(grouped_data)
            
            # Sort by grouping key
            grouped_df = grouped_df.sort_values('グループ化キー')
            
            status_label.config(text="Saving enhanced results to Excel...")
            self.root.update()
            
            # Save to Excel with enhanced structure
            self.save_enhanced_grouped_data(grouped_df)
            
            # Close progress window
            progress_window.destroy()
            
            # Brief pause to show completion
            self.root.after(500, lambda: self.show_enhanced_completion_dialog(len(grouped_df), len(df)))
            
        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("Error", f"Error during enhanced grouping execution:\n{str(e)}")

    def save_enhanced_grouped_data(self, grouped_df):
        """Save enhanced grouped data to Excel sheet with proper column order"""
        # Enhanced column order: データ件数 → 路線名 → 路線名略称 → 構造物番号 → other columns → year columns
        base_columns = ['グループ化キー', 'グループ化方法', '種別','点検区分1', '構造物名称', '駅（始）', '駅（至）','データ件数', '路線名', '路線名略称', '構造物番号']
        
        # Get year columns dynamically and sort from oldest to newest
        year_columns = [col for col in grouped_df.columns if col.endswith(' 結果')]
        
        # Extract years and sort them to ensure correct chronological order
        def extract_year(col_name):
            """Extract year from column name like '2024 結果' """
            try:
                return int(col_name.split(' ')[0])
            except:
                return 0
        
        # Sort year columns by actual year value (oldest to newest)
        year_columns.sort(key=extract_year)
        
        final_columns = base_columns + year_columns
        existing_columns = [col for col in final_columns if col in grouped_df.columns]
        grouped_df = grouped_df[existing_columns]
        
        # Sort by the latest year column in descending order (highest values first)
        if year_columns:
            latest_year_col = year_columns[-1]  # Get the last (newest) year column
            # Convert to numeric for proper sorting, handle non-numeric values
            grouped_df[latest_year_col] = pd.to_numeric(grouped_df[latest_year_col], errors='coerce').fillna(0)
            grouped_df = grouped_df.sort_values(latest_year_col, ascending=False)
        
        # Save to Excel
        with pd.ExcelWriter(self.workbook_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            grouped_df.to_excel(writer, sheet_name='グループ化点検履歴', index=False)
            
            # Preserve other sheets
            try:
                original_wb = load_workbook(self.workbook_path)
                for sheet_name in original_wb.sheetnames:
                    if sheet_name != 'グループ化点検履歴':
                        try:
                            df_temp = pd.read_excel(self.workbook_path, sheet_name=sheet_name)
                            df_temp.to_excel(writer, sheet_name=sheet_name, index=False)
                        except Exception as e:
                            print(f"Could not preserve sheet {sheet_name}: {e}")
                            continue
            except Exception as e:
                print(f"Error preserving sheets: {e}")
                pass

    def show_enhanced_completion_dialog(self, grouped_count, original_count):
        """Show enhanced completion dialog with results summary"""
        completion_window = tk.Toplevel(self.main_window)
        completion_window.title("Enhanced Grouping Complete")
        completion_window.geometry("550x450")
        completion_window.grab_set()
        completion_window.resizable(False, False)
        completion_window.transient(self.main_window)
        
        main_frame = tk.Frame(completion_window, padx=30, pady=30)
        main_frame.pack(fill="both", expand=True)
        
        # Success title
        title_frame = tk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 20))
        
        success_label = tk.Label(title_frame, text="✓", font=("Arial", 24, "bold"), fg="green")
        success_label.pack(side="left")
        
        title_label = tk.Label(title_frame, text="Enhanced Grouping Complete!", 
                            font=("Arial", 16, "bold"), fg="navy")
        title_label.pack(side="left", padx=(10, 0))
        
        # Enhanced features summary
        enhancement_frame = tk.LabelFrame(main_frame, text="Enhanced Features Applied", 
                                        font=("Arial", 12, "bold"), padx=20, pady=15)
        enhancement_frame.pack(fill="x", pady=(0, 15))
        
        enhancement_details = [
            ("✓ Column Order:", "データ件数 → 路線名 → 路線名略称 → 構造物番号"),
            ("✓ Smart Grouping Keys:", "No 点検区分1 when 'All' selected"),
            ("✓ Route Abbreviations:", "東急多摩川線→TM, 東横線→TY, etc."),
            ("✓ Structure Numbers:", "Auto-lookup from 構造物番号 sheet" if self.structure_df is not None else "Not available (no 構造物番号 sheet)")
        ]
        
        for i, (feature, description) in enumerate(enhancement_details):
            detail_frame = tk.Frame(enhancement_frame)
            detail_frame.pack(fill="x", pady=3)
            
            feature_label = tk.Label(detail_frame, text=feature, font=("Arial", 10, "bold"), 
                                   width=20, anchor="w", fg="darkgreen")
            feature_label.pack(side="left")
            
            desc_label = tk.Label(detail_frame, text=description, font=("Arial", 10), 
                                fg="blue" if i < 3 else ("blue" if self.structure_df is not None else "orange"))
            desc_label.pack(side="left", padx=(5, 0))
        
        # Processing summary
        summary_frame = tk.LabelFrame(main_frame, text="Processing Summary", 
                                    font=("Arial", 12, "bold"), padx=20, pady=15)
        summary_frame.pack(fill="x", pady=(0, 20))
        
        summary_details = [
            ("Original Data Count:", f"{original_count:,} records"),
            ("Enhanced Grouped Count:", f"{grouped_count:,} records"),
            ("Reduction Rate:", f"{((original_count - grouped_count) / original_count * 100):.1f}% reduction"),
            ("New Sheet Name:", "\"グループ化点検履歴\" (Enhanced)"),
            ("Save Location:", "Same Excel file")
        ]
        
        for i, (label, value) in enumerate(summary_details):
            detail_frame = tk.Frame(summary_frame)
            detail_frame.pack(fill="x", pady=3)
            
            label_widget = tk.Label(detail_frame, text=label, font=("Arial", 10, "bold"), 
                                  width=20, anchor="w")
            label_widget.pack(side="left")
            
            value_widget = tk.Label(detail_frame, text=value, font=("Arial", 10), fg="blue")
            value_widget.pack(side="left", padx=(10, 0))
        
        # Next steps
        next_steps_frame = tk.LabelFrame(main_frame, text="Next Steps", 
                                        font=("Arial", 12, "bold"), padx=20, pady=15)
        next_steps_frame.pack(fill="x", pady=(0, 20))
        
        steps_text = ("1. Check the enhanced 'グループ化点検履歴' sheet\n"
                    "2. Review new columns: データ件数, 路線名略称, 構造物番号\n"
                    "3. Analyze grouped data with enhanced features\n"
                    "4. Proceed to next processing steps")
        
        steps_label = tk.Label(next_steps_frame, text=steps_text, font=("Arial", 10), 
                            justify="left", wraplength=450)
        steps_label.pack(anchor="w")
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        def open_excel():
            try:
                import os
                os.startfile(self.workbook_path)
                
                # Show brief message and then auto-close everything
                messagebox.showinfo("Excel Opened", 
                                "✅ Enhanced Excel file opened successfully!\n\n"
                                "Check the new enhanced columns in グループ化点検履歴 sheet.\n\n"
                                "The application will close automatically.")
                
                # Auto-close all windows after opening Excel
                completion_window.after(1000, lambda: [
                    completion_window.destroy(),
                    self.main_window.destroy() if hasattr(self, 'main_window') and self.main_window.winfo_exists() else None,
                    self.root.after(1000, self.root.quit)
                ])
                
            except Exception as e:
                messagebox.showinfo("Info", f"Please open the file manually:\n{self.workbook_path}")

        def close_and_exit():
            # Close completion dialog
            completion_window.destroy()
            
            # Close main window if it exists
            if hasattr(self, 'main_window'):
                try:
                    if self.main_window.winfo_exists():
                        self.main_window.destroy()
                except:
                    pass
            
            # Show final success message briefly
            messagebox.showinfo("Enhanced Process Complete", 
                            "✅ Enhanced data grouping completed successfully!\n\n"
                            f"📊 Results saved with enhanced features:\n"
                            f"• データ件数, 路線名略称, 構造物番号 columns\n"
                            f"• Smart grouping keys\n"
                            f"• Improved column layout\n\n"
                            "The application will now close automatically.")
            
            # Close the main application after showing message
            self.root.after(2000, self.root.quit)  # Close after 2 seconds
        
        # Buttons
        excel_btn = tk.Button(button_frame, text="Open Enhanced Excel", 
                            command=open_excel, bg="#4CAF50", fg="white", 
                            width=18, height=2, font=("Arial", 11))
        excel_btn.pack(side="left", padx=10)
        
        close_btn = tk.Button(button_frame, text="Complete", 
                            command=close_and_exit, bg="#2196F3", fg="white", 
                            width=15, height=2, font=("Arial", 11))
        close_btn.pack(side="right", padx=10)

    def run(self):
        """Run the enhanced application"""
        self.root.mainloop()


# Main execution
if __name__ == "__main__":
    print("Enhanced Data Grouping System Starting...")
    print("=" * 60)
    print("🚀 Enhanced Features:")
    print("• データ件数 column (before 路線名)")
    print("• 路線名略称 column (after 路線名)")
    print("• 構造物番号 column (after 路線名略称)")
    print("• Smart grouping keys (no 点検区分1 when 'All')")
    print("• Route name abbreviations (TM, TY, OM, etc.)")
    print("• Auto-lookup structure numbers from 構造物番号 sheet")
    print("• Enhanced column positioning")
    print("• 'All' option instead of '*' in UI")
    print("=" * 60)
    
    app = EnhancedDataGroupingApp()
    app.run()