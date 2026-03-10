import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import polars as pl
import os
from datetime import datetime, timedelta
import json

class MaterialDispatchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Material Dispatch Management System")
        self.root.geometry("1100x850")
        
        # Data storage paths
        self.inventory_file = "inventory.parquet"
        self.dispatched_file = "dispatched.parquet"
        self.config_file = "last_dispatch.json"
        self.security_file = "security_config.json"
        
        # Initialize Security & Data
        self.check_security()
        self.inventory = self.load_data(self.inventory_file)
        self.dispatched = self.load_data(self.dispatched_file)
        self.last_dispatch = self.load_config()
        
        # Hide main window until login
        self.root.withdraw()
        self.show_login()

    def check_security(self):
        """Initialize or check the 6-month trial period."""
        if not os.path.exists(self.security_file):
            first_run = datetime.now().strftime("%Y-%m-%d")
            with open(self.security_file, 'w') as f:
                json.dump({"first_run_date": first_run}, f)
            self.first_run_date = datetime.now()
        else:
            with open(self.security_file, 'r') as f:
                data = json.load(f)
                self.first_run_date = datetime.strptime(data["first_run_date"], "%Y-%m-%d")

    def show_login(self):
        login_win = tk.Toplevel(self.root)
        login_win.title("Security Login")
        login_win.geometry("400x300")
        login_win.resizable(False, False)
        
        # Center the login window
        login_win.update_idletasks()
        width = login_win.winfo_width()
        height = login_win.winfo_height()
        x = (login_win.winfo_screenwidth() // 2) - (width // 2)
        y = (login_win.winfo_screenheight() // 2) - (height // 2)
        login_win.geometry(f'{width}x{height}+{x}+{y}')

        # Determine if trial has expired (6 months = ~180 days)
        days_passed = (datetime.now() - self.first_run_date).days
        is_expired = days_passed > 180

        ttk.Label(login_win, text="System Login", font=('Helvetica', 14, 'bold')).pack(pady=20)
        
        if is_expired:
            ttk.Label(login_win, text="TRIAL EXPIRED: Enter Master Password", foreground="red").pack()
        
        frame = ttk.Frame(login_win, padding=20)
        frame.pack(fill="both")

        ttk.Label(frame, text="Username:").grid(row=0, column=0, sticky="w", pady=5)
        user_ent = ttk.Entry(frame)
        user_ent.grid(row=0, column=1, pady=5, padx=5)
        user_ent.insert(0, "Admin")

        ttk.Label(frame, text="Password:").grid(row=1, column=0, sticky="w", pady=5)
        pass_ent = ttk.Entry(frame, show="*")
        pass_ent.grid(row=1, column=1, pady=5, padx=5)

        def attempt_login():
            username = user_ent.get()
            password = pass_ent.get()
            
            if is_expired:
                if password == "Rajesh4568@123":
                    login_win.destroy()
                    self.setup_ui()
                    self.root.deiconify()
                else:
                    messagebox.showerror("Access Denied", "Invalid Master Password. Please contact administrator.")
            else:
                if (username == "Admin" and password == "123") or password == "Rajesh4568@123":
                    login_win.destroy()
                    self.setup_ui()
                    self.root.deiconify()
                else:
                    messagebox.showerror("Error", "Invalid Credentials")

        ttk.Button(login_win, text="Login", command=attempt_login).pack(pady=10)
        login_win.protocol("WM_DELETE_WINDOW", self.root.quit)

    def load_data(self, filepath):
        if os.path.exists(filepath):
            try:
                df = pl.read_parquet(filepath)
                # Force all columns to string and handle nulls
                return df.with_columns(pl.all().cast(pl.String).fill_null(""))
            except Exception:
                return pl.DataFrame()
        return pl.DataFrame()

    def save_data(self):
        if self.inventory.height > 0:
            self.inventory.write_parquet(self.inventory_file)
        elif os.path.exists(self.inventory_file):
            self.inventory.write_parquet(self.inventory_file)

        if self.dispatched.height > 0:
            self.dispatched.write_parquet(self.dispatched_file)
        elif os.path.exists(self.dispatched_file):
            self.dispatched.write_parquet(self.dispatched_file)

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                return json.load(f)
        return {}

    def save_config(self, details):
        with open(self.config_file, 'w') as f:
            json.dump(details, f)

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook.Tab", padding=[20, 10], font=('Helvetica', 10, 'bold'))
        style.configure("Treeview.Heading", font=('Helvetica', 9, 'bold'))
        
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(expand=1, fill="both", padx=10, pady=10)

        self.home_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.home_tab, text=" 🏠 Home ")
        self.setup_home_tab()

        self.dispatch_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.dispatch_tab, text=" 🚛 Dispatch ")
        self.setup_dispatch_tab()

    def setup_home_tab(self):
        top_frame = ttk.Frame(self.home_tab, padding=10)
        top_frame.pack(fill="x")

        ttk.Button(top_frame, text="📥 Import Packing Data (CSV/Excel)", command=self.import_data).pack(side="left", padx=5)
        ttk.Button(top_frame, text="📤 Export Remaining Data", command=self.export_remaining).pack(side="left", padx=5)
        ttk.Button(top_frame, text="🧹 Clear Table View", command=self.clear_home_tree).pack(side="left", padx=5)

        # View Filter UI
        ttk.Label(top_frame, text="View:").pack(side="left", padx=(15, 5))
        self.view_filter = tk.StringVar(value="In Inventory")
        filter_box = ttk.Combobox(top_frame, textvariable=self.view_filter, values=["In Inventory", "Dispatched", "All History"], state="readonly", width=15)
        filter_box.pack(side="left", padx=5)
        filter_box.bind("<<ComboboxSelected>>", lambda e: self.refresh_inventory_table())

        table_frame = ttk.Frame(self.home_tab, padding=10)
        table_frame.pack(expand=True, fill="both")

        cols = ("S.No.", "Date", "Shift", "Incharge", "Pallet_ID", "watt", "Binning", "JB_Type", "Module_Type", "Grade", "Status")
        self.tree = ttk.Treeview(table_frame, columns=cols, show='headings')
        
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90, anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", expand=True, fill="both")
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.home_tab, padding=10)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="🔍 View Module IDs", command=self.show_modules).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="ℹ️ Check Status / Details", command=self.check_status).pack(side="right", padx=5)

        self.refresh_inventory_table()

    def setup_dispatch_tab(self):
        main_frame = ttk.Frame(self.dispatch_tab, padding=20)
        main_frame.pack(expand=True, fill="both")

        # 1. Scanning Section
        scan_label = ttk.Label(main_frame, text="Scan Pallet ID", font=('Helvetica', 14, 'bold'))
        scan_label.pack(pady=(0, 5))

        self.scan_var = tk.StringVar()
        self.scan_entry = ttk.Entry(main_frame, textvariable=self.scan_var, font=('Helvetica', 16), justify='center', width=30)
        self.scan_entry.pack(pady=5)
        self.scan_entry.bind("<Return>", lambda e: self.process_scan())

        ttk.Button(main_frame, text="Process Scan", command=self.process_scan).pack(pady=10)

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15)
        
        # 2. Live Log Header
        header_row = ttk.Frame(main_frame)
        header_row.pack(fill="x", pady=5)
        ttk.Label(header_row, text="Live Dispatched Log", font=('Helvetica', 11, 'bold')).pack(side="left", anchor="w")
        ttk.Button(header_row, text="🧹 Clear Log View", command=self.clear_dispatch_tree).pack(side="right", padx=5)
        
        # 3. DISPATCH FILTER SECTION (New)
        filter_frame = ttk.Frame(main_frame, padding=(0, 5))
        filter_frame.pack(fill="x")
        
        ttk.Label(filter_frame, text="Filter By:").pack(side="left", padx=5)
        self.dispatch_search_col = tk.StringVar(value="Vehicle Number")
        search_col_box = ttk.Combobox(filter_frame, textvariable=self.dispatch_search_col, 
                                     values=["Vehicle Number", "Customer Name", "Dispatch Date", "Transporter"], 
                                     state="readonly", width=18)
        search_col_box.pack(side="left", padx=5)
        
        ttk.Label(filter_frame, text="Search:").pack(side="left", padx=5)
        self.dispatch_search_val = tk.StringVar()
        search_ent = ttk.Entry(filter_frame, textvariable=self.dispatch_search_val, width=25)
        search_ent.pack(side="left", padx=5)
        search_ent.bind("<KeyRelease>", lambda e: self.refresh_dispatch_table())
        
        ttk.Button(filter_frame, text="Apply", command=self.refresh_dispatch_table).pack(side="left", padx=5)
        
        def reset_dispatch_filter():
            self.dispatch_search_val.set("")
            self.refresh_dispatch_table()
            
        ttk.Button(filter_frame, text="Reset", command=reset_dispatch_filter).pack(side="left", padx=5)

        # 4. Table Section
        table_frame = ttk.Frame(main_frame, padding=(0, 10))
        table_frame.pack(expand=True, fill="both")

        dispatch_cols = ("Pallet_ID", "Customer Name", "Vehicle Number", "Driver Name", "Dispatch Date", "Transporter")
        self.dispatch_tree = ttk.Treeview(table_frame, columns=dispatch_cols, show='headings', height=10, selectmode='extended')
        
        for col in dispatch_cols:
            self.dispatch_tree.heading(col, text=col)
            self.dispatch_tree.column(col, width=120, anchor="center")

        d_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.dispatch_tree.yview)
        self.dispatch_tree.configure(yscrollcommand=d_scrollbar.set)
        
        self.dispatch_tree.pack(side="left", expand=True, fill="both")
        d_scrollbar.pack(side="right", fill="y")

        # 5. Bottom Actions
        btn_row = ttk.Frame(main_frame, padding=10)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="❌ Delete Selected Dispatch", command=self.delete_dispatch).pack(side="left", padx=5)
        ttk.Button(btn_row, text="📄 Generate Sheet (Vehicle Report)", command=self.generate_vehicle_report).pack(side="left", padx=5)

        export_frame = ttk.Frame(self.dispatch_tab, padding=20)
        export_frame.pack(side="bottom", fill="x")
        ttk.Button(export_frame, text="📑 Export All Dispatched Data", command=self.export_dispatched).pack(side="right")
        
        self.refresh_dispatch_table()

    def clear_home_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def clear_dispatch_tree(self):
        for item in self.dispatch_tree.get_children():
            self.dispatch_tree.delete(item)

    def delete_dispatch(self):
        selected_items = self.dispatch_tree.selection()
        if not selected_items:
            messagebox.showwarning("Selection Required", "Please select one or more pallets to delete.")
            return
        
        # Ensure ID is treated as normalized string
        pallet_ids = [str(self.dispatch_tree.item(item)['values'][0]).strip().upper() for item in selected_items]

        pw_win = tk.Toplevel(self.root)
        pw_win.title("Authorize Deletion")
        pw_win.geometry("300x150")
        pw_win.grab_set()

        ttk.Label(pw_win, text=f"Enter Password to Delete {len(pallet_ids)} item(s):").pack(pady=10)
        pw_ent = ttk.Entry(pw_win, show="*")
        pw_ent.pack(pady=5)
        pw_ent.focus_set()

        def confirm_delete():
            if pw_ent.get() == "Change@123":
                self.dispatched = self.dispatched.filter(~pl.col("Pallet_ID").is_in(pallet_ids))
                self.save_data()
                self.refresh_dispatch_table()
                self.refresh_inventory_table()
                pw_win.destroy()
                messagebox.showinfo("Deleted", f"{len(pallet_ids)} Pallet(s) removed from dispatched history.")
            else:
                messagebox.showerror("Error", "Incorrect Password")

        ttk.Button(pw_win, text="Confirm", command=confirm_delete).pack(pady=10)

    def generate_vehicle_report(self):
        """Exports currently visible dispatched data as a vehicle report."""
        if self.dispatched.height == 0:
            messagebox.showwarning("No Data", "No dispatched data available to generate report.")
            return
            
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            # Note: This exports all dispatched data currently in state, matching the live log logic
            self.dispatched.write_csv(path)
            messagebox.showinfo("Generated", f"Vehicle dispatch sheet generated successfully at:\n{path}")

    def import_data(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("All Compatible Files", "*.csv *.xlsx *.xls *.xlsm *.xlsb"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb")
        ])
        if not file_path:
            return
        
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.csv':
                try:
                    df = pl.read_csv(file_path, infer_schema_length=0, encoding="utf8")
                except Exception:
                    try:
                        df = pl.read_csv(file_path, infer_schema_length=0, encoding="latin-1")
                    except Exception as e:
                        messagebox.showerror("Import Error", f"Could not decode CSV file: {str(e)}")
                        return
            else:
                try:
                    df = pl.read_excel(file_path)
                except ImportError:
                    messagebox.showerror("Error", "To import Excel files, please install: pip install fastexcel calamine")
                    return

            df = df.with_columns(pl.all().cast(pl.String).fill_null(""))
            df = df.with_columns([
                pl.col("Pallet_ID").str.strip_chars().str.to_uppercase(),
                pl.col("Module_ID").str.strip_chars().str.to_uppercase()
            ])

            required = ["S.No.", "Date", "Shift", "Incharge", "watt", "Pallet_ID", "Module_ID", "Binning", "JB_Type", "Module_Type", "Grade"]
            for col in required:
                if col not in df.columns:
                    messagebox.showerror("Error", f"Missing column: {col}")
                    return

            df = df.filter(
                (pl.col("Pallet_ID") != "") & 
                (pl.col("Pallet_ID") != "NONE") & 
                (pl.col("Pallet_ID") != "NULL")
            )

            if self.dispatched.height > 0:
                dispatched_pallets = set(self.dispatched['Pallet_ID'].to_list())
                imported_pallets = set(df['Pallet_ID'].to_list())
                already_dispatched = imported_pallets.intersection(dispatched_pallets)
                
                if already_dispatched:
                    skipped_list = list(already_dispatched)
                    skipped_msg = ", ".join(skipped_list[:5])
                    if len(skipped_list) > 5:
                        skipped_msg += f" and {len(skipped_list)-5} others"
                    messagebox.showinfo("Import Info", f"Skipping {len(already_dispatched)} already dispatched pallets: {skipped_msg}.")
                    df = df.filter(~pl.col("Pallet_ID").is_in(list(already_dispatched)))

            if self.inventory.height > 0:
                current_inv_pallets = set(self.inventory['Pallet_ID'].to_list())
                imported_pallets = set(df['Pallet_ID'].to_list())
                already_in_inv = imported_pallets.intersection(current_inv_pallets)
                
                if already_in_inv:
                    df = df.filter(~pl.col("Pallet_ID").is_in(list(already_in_inv)))

            if df.height == 0:
                messagebox.showwarning("Empty Import", "No new pallets to import.")
                return

            if df['Module_ID'].n_unique() != df.height:
                messagebox.showerror("Duplicate Error", "The imported file contains internal duplicate Module IDs.")
                return

            if self.inventory.height > 0:
                existing_inv_mids = set(self.inventory['Module_ID'].to_list())
                new_mids = set(df['Module_ID'].to_list())
                dups = new_mids.intersection(existing_inv_mids)
                if dups:
                    messagebox.showerror("Duplicate Error", f"Duplicate Module IDs found in Inventory. Example: {list(dups)[0]}")
                    return

            if self.inventory.height == 0:
                self.inventory = df
            else:
                self.inventory = pl.concat([self.inventory, df], how="diagonal")

            self.save_data()
            self.refresh_inventory_table()
            messagebox.showinfo("Success", f"Successfully imported {df.unique(subset=['Pallet_ID']).height} new pallets.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import: {str(e)}")

    def refresh_inventory_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not hasattr(self, 'view_filter'):
            return 

        filter_type = self.view_filter.get()
        df_to_show = pl.DataFrame()

        if filter_type == "In Inventory":
            if self.inventory.height > 0:
                df_to_show = self.inventory.unique(subset=['Pallet_ID']).with_columns(pl.lit("In Inventory").alias("Display_Status"))
        elif filter_type == "Dispatched":
            if self.dispatched.height > 0:
                df_to_show = self.dispatched.unique(subset=['Pallet_ID']).with_columns(pl.lit("Dispatched").alias("Display_Status"))
        else: # All History
            inv_p = pl.DataFrame()
            if self.inventory.height > 0:
                inv_p = self.inventory.unique(subset=['Pallet_ID']).with_columns(pl.lit("In Inventory").alias("Display_Status"))
            disp_p = pl.DataFrame()
            if self.dispatched.height > 0:
                disp_p = self.dispatched.unique(subset=['Pallet_ID']).with_columns(pl.lit("Dispatched").alias("Display_Status"))
            
            if inv_p.height > 0 and disp_p.height > 0:
                df_to_show = pl.concat([inv_p, disp_p], how="diagonal")
            elif inv_p.height > 0:
                df_to_show = inv_p
            elif disp_p.height > 0:
                df_to_show = disp_p

        if df_to_show.height == 0:
            return

        for row in df_to_show.iter_rows(named=True):
            self.tree.insert("", "end", values=(
                row.get('S.No.', ''), row.get('Date', ''), row.get('Shift', ''), row.get('Incharge', ''),
                row.get('Pallet_ID', ''), row.get('watt', ''), row.get('Binning', ''), row.get('JB_Type', ''),
                row.get('Module_Type', ''), row.get('Grade', ''), row.get('Display_Status', '')
            ))

    def refresh_dispatch_table(self):
        for item in self.dispatch_tree.get_children():
            self.dispatch_tree.delete(item)
            
        if self.dispatched.height == 0:
            return

        # Filtering logic for Dispatch Tab (New)
        df_to_show = self.dispatched.unique(subset=['Pallet_ID'])
        
        search_col = self.dispatch_search_col.get()
        search_val = self.dispatch_search_val.get().strip().upper()
        
        if search_val:
            # Search logic based on selected column
            if search_col == "Vehicle Number":
                df_to_show = df_to_show.filter(pl.col("Vehicle Number").str.to_uppercase().str.contains(search_val))
            elif search_col == "Customer Name":
                df_to_show = df_to_show.filter(pl.col("Customer Name").str.to_uppercase().str.contains(search_val))
            elif search_col == "Dispatch Date":
                df_to_show = df_to_show.filter(pl.col("Dispatch Date").str.contains(search_val))
            elif search_col == "Transporter":
                df_to_show = df_to_show.filter(pl.col("Transporter").str.to_uppercase().str.contains(search_val))

        if df_to_show.height == 0:
            return

        for row in df_to_show.iter_rows(named=True):
            self.dispatch_tree.insert("", 0, values=(
                row['Pallet_ID'], 
                row['Customer Name'], 
                row['Vehicle Number'], 
                row['Driver Name'], 
                row['Dispatch Date'], 
                row.get('Transporter', '')
            ))

    def show_modules(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Select a pallet first")
            return
        
        pallet_id = str(self.tree.item(selected[0])['values'][4]).strip().upper()
        
        all_matches = []
        if self.inventory.height > 0:
            m_inv = self.inventory.filter(pl.col('Pallet_ID') == pallet_id)
            if m_inv.height > 0: all_matches.extend(m_inv['Module_ID'].to_list())
            
        if self.dispatched.height > 0:
            m_disp = self.dispatched.filter(pl.col('Pallet_ID') == pallet_id)
            if m_disp.height > 0: all_matches.extend(m_disp['Module_ID'].to_list())
            
        unique_modules = sorted(list(set(all_matches)))
            
        win = tk.Toplevel(self.root)
        win.title(f"Modules in {pallet_id}")
        win.geometry("400x500")
        txt = tk.Text(win, padx=10, pady=10)
        txt.pack(expand=True, fill="both")
        txt.insert("1.0", f"Total Modules: {len(unique_modules)}\n" + "-"*30 + "\n")
        for m in unique_modules:
            txt.insert("end", f"• {m}\n")
        txt.config(state="disabled")

    def check_status(self):
        selected = self.tree.selection()
        if not selected:
            return
        pallet_id = str(self.tree.item(selected[0])['values'][4]).strip().upper()
        match = pl.DataFrame()
        if self.dispatched.height > 0:
            match = self.dispatched.filter(pl.col('Pallet_ID') == pallet_id)
        if match.height > 0:
            details = match.row(0, named=True)
            msg = (f"PALLET DISPATCHED\n\nDate: {details['Dispatch Date']}\nCustomer: {details['Customer Name']}\n"
                   f"Vehicle: {details['Vehicle Number']}\nDriver: {details['Driver Name']}\nIncharge: {details['Dispatch Incharge']}")
            messagebox.showinfo("Dispatch Details", msg)
        else:
            messagebox.showinfo("Status", f"Pallet {pallet_id} is still in inventory.")

    def process_scan(self):
        pid = self.scan_var.get().strip().upper() 
        if not pid: return
        
        if self.dispatched.height > 0 and pid in self.dispatched['Pallet_ID'].to_list():
            messagebox.showerror("Error", f"Pallet {pid} already dispatched!")
            self.scan_var.set("")
            return
        if self.inventory.height == 0 or pid not in self.inventory['Pallet_ID'].to_list():
            messagebox.showerror("Error", f"Pallet {pid} not found in inventory.")
            self.scan_var.set("")
            return
        self.open_dispatch_modal(pid)

    def open_dispatch_modal(self, pallet_id):
        modal = tk.Toplevel(self.root)
        modal.title(f"Dispatch Details - {pallet_id}")
        modal.geometry("500x600")
        modal.grab_set()
        
        # Fields updated: Changed "Dispatch Shift" to "Transporter"
        fields = [
            ("Dispatch Date", datetime.now().strftime("%Y-%m-%d")), 
            ("Customer Name", ""), 
            ("Address", ""), 
            ("Driver Name", ""), 
            ("Vehicle Number", ""), 
            ("Driver Mob. No.", ""), 
            ("Dispatch Incharge", ""), 
            ("Transporter", "")
        ]
        
        entries = {}
        for label, default in fields:
            row = ttk.Frame(modal, padding=5)
            row.pack(fill="x")
            ttk.Label(row, text=label, width=20).pack(side="left")
            ent = ttk.Entry(row)
            ent.insert(0, default)
            ent.pack(side="right", expand=True, fill="x")
            entries[label] = ent
            
        def fill_previous():
            if self.last_dispatch:
                for k, v in self.last_dispatch.items():
                    if k in entries and k != "Dispatch Date":
                        entries[k].delete(0, tk.END); entries[k].insert(0, v)
                        
        ttk.Button(modal, text="📋 Select data same as previous", command=fill_previous).pack(pady=10)
        
        def on_done():
            data = {k: v.get() for k, v in entries.items()}
            if not data["Customer Name"] or not data["Vehicle Number"]:
                messagebox.showwarning("Incomplete", "Please fill Customer and Vehicle details.")
                return
            self.last_dispatch = data
            self.save_config(data)
            pallet_rows = self.inventory.filter(pl.col('Pallet_ID') == pallet_id)
            pallet_rows = pallet_rows.with_columns(pl.all().cast(pl.String))
            for k, v in data.items():
                pallet_rows = pallet_rows.with_columns(pl.lit(str(v)).alias(k))
            
            if self.dispatched.height == 0: 
                self.dispatched = pallet_rows
            else: 
                self.dispatched = pl.concat([self.dispatched, pallet_rows], how="diagonal")
            
            self.inventory = self.inventory.filter(pl.col('Pallet_ID') != pallet_id)
            self.save_data()
            self.refresh_inventory_table()
            self.refresh_dispatch_table()
            self.scan_var.set("")
            modal.destroy()
            messagebox.showinfo("Success", f"Pallet {pallet_id} dispatched.")
            
        btn_row = ttk.Frame(modal, padding=10)
        btn_row.pack(fill="x", side="bottom")
        ttk.Button(btn_row, text="Done", command=on_done).pack(side="right", padx=5)
        ttk.Button(btn_row, text="Close", command=modal.destroy).pack(side="right", padx=5)

    def export_remaining(self):
        if self.inventory.height == 0: return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path: self.inventory.write_csv(path); messagebox.showinfo("Exported", "Inventory exported.")

    def export_dispatched(self):
        if self.dispatched.height == 0: return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path: self.dispatched.write_csv(path); messagebox.showinfo("Exported", "Dispatch history exported.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialDispatchApp(root)
    root.mainloop()