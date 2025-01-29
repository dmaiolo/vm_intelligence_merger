import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from tkinter.scrolledtext import ScrolledText

# Application Constants
VERSION = "0.003"
PORT_MAPPING = {
    53: 'DNS Server',
    80: 'Web Server',
    443: 'Secure Web Server',
    22: 'SSH Server',
    3306: 'MySQL Database',
    1433: 'Microsoft SQL Server',
    3389: 'Remote Desktop',
    5007: 'Palo Alto',
    67: 'DHCP Server',
    445: 'SMB/File Services',
    8530: 'Custom Application',
    647: 'Custom Service'
}

OS_MAPPING = {
    'CentOS': 'Nutanix',
    'PanOS': 'Palo Alto',
    'Windows Server': 'Windows Server',
    'ESXi': 'VMware Hypervisor',
    'Nutanix': 'Nutanix CVM'
}

def process_notes(notes):
    """Extract Real App Name from Notes by removing RITM prefix."""
    if pd.isna(notes) or not str(notes).strip():
        return ''
    parts = str(notes).split()
    if len(parts) >= 2 and parts[0].startswith('RITM') and parts[0][4:].isdigit():
        return ' '.join(parts[1:])
    return notes

def determine_real_app(row):
    """Determine application name using multiple intelligence sources."""
    # First priority: Processed Notes
    notes_app = process_notes(row.get('Notes', ''))
    if notes_app.strip():
        return notes_app
    
    # Second priority: Recognized applications
    discovered_app = str(row.get('Discovered App', ''))
    if discovered_app.endswith('_Recognized'):
        return discovered_app.replace('_Recognized', '').strip()
    
    # Third priority: Port analysis
    ports = str(row.get('Feature Ports', '')).strip()
    if ports:
        port_list = [p.strip() for p in ports.split(',')]
        for port_str in port_list:
            try:
                port = int(port_str)
                if port in PORT_MAPPING:
                    return PORT_MAPPING[port]
            except ValueError:
                continue
    
    # Fourth priority: OS analysis
    os_version = str(row.get('OS Version', '')).strip()
    for os_key, app_name in OS_MAPPING.items():
        if os_key in os_version:
            return app_name
    
    # Fifth priority: DNS name pattern matching
    dns_name = str(row.get('DNS Name', '')).lower()
    if 'sql' in dns_name:
        return 'SQL Server'
    if 'web' in dns_name or 'http' in dns_name:
        return 'Web Server'
    
    # Final fallback
    return 'UNKNOWN - Needs Manual Review'

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"VM Intelligence Merger v{VERSION} - David Maiolo")
        self.geometry("650x450")
        self.iteration0_path = None
        self.iteration1_path = None
        
        self._configure_styles()
        self._create_menu()
        self._create_widgets()
        self._create_status_bar()

    def _configure_styles(self):
        self.configure(bg='#f0f0f0')
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TButton', foreground='navy', font=('Helvetica', 10))
        style.configure('TLabel', background='#f0f0f0', font=('Helvetica', 9))
        style.configure('Status.TLabel', background='lightgray', font=('Helvetica', 9))

    def _create_menu(self):
        menubar = tk.Menu(self)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="User Guide", command=self.show_help)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

    def _create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(padx=20, pady=10, fill='both', expand=True)
        
        ttk.Label(main_frame, 
                text="VM Intelligence Merger", 
                font=('Helvetica', 16, 'bold'), 
                foreground='darkblue').pack(pady=10)
        
        file_frame = ttk.LabelFrame(main_frame, text=" Processing Steps ")
        file_frame.pack(fill='x', pady=5)
        
        ttk.Button(file_frame, 
                 text="1. Select Iteration 0 (CSV)", 
                 command=self.select_iteration0).pack(fill='x', pady=5)
        self.label_iter0 = ttk.Label(file_frame, text="No file selected")
        self.label_iter0.pack()
        
        ttk.Button(file_frame, 
                 text="2. Select Iteration 1 (XLSX)", 
                 command=self.select_iteration1).pack(fill='x', pady=5)
        self.label_iter1 = ttk.Label(file_frame, text="No file selected")
        self.label_iter1.pack()
        
        self.process_btn = ttk.Button(main_frame, 
                                    text="3. Process and Export", 
                                    command=self.process_files, 
                                    state='disabled')
        self.process_btn.pack(pady=10)

    def _create_status_bar(self):
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self, 
                             textvariable=self.status_var, 
                             style='Status.TLabel',
                             padding=5,
                             anchor='w')
        status_bar.pack(side='bottom', fill='x')

    def update_status(self, message):
        self.status_var.set(message)
        self.update_idletasks()

    def select_iteration0(self):
        self.update_status("Selecting Iteration 0 CSV file...")
        file_path = filedialog.askopenfilename(title="Select Iteration 0 CSV", 
                                             filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.iteration0_path = file_path
            self.label_iter0.config(text=file_path.split('/')[-1])
            self.check_files_selected()
        self.update_status("Ready")

    def select_iteration1(self):
        self.update_status("Selecting Iteration 1 XLSX file...")
        file_path = filedialog.askopenfilename(title="Select Iteration 1 XLSX", 
                                             filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.iteration1_path = file_path
            self.label_iter1.config(text=file_path.split('/')[-1])
            self.check_files_selected()
        self.update_status("Ready")

    def check_files_selected(self):
        if self.iteration0_path and self.iteration1_path:
            self.process_btn.config(state='enabled')

    def process_files(self):
        try:
            self.process_btn.config(state='disabled')
            self.update_status("Initializing processing...")
            
            # Load data with column validation
            self.update_status("Loading Iteration 0 CSV...")
            csv_df = pd.read_csv(self.iteration0_path)
            self._validate_columns(csv_df, 'CSV', {'Name', 'Notes', 'OS Version', 'DNS Name'})
            
            self.update_status("Loading Iteration 1 XLSX...")
            xlsx_df = pd.read_excel(self.iteration1_path)
            self._validate_columns(xlsx_df, 'XLSX', {'Host', 'Discovered App', 'Feature Ports'})
            
            # Merge and process data
            self.update_status("Matching records across files...")
            xlsx_df['Name'] = xlsx_df['Host'].str.split(':').str[0]
            merged_df = pd.merge(xlsx_df, csv_df, on='Name', how='left')
            
            self.update_status("Analyzing multiple intelligence sources...")
            merged_df['Real App Name'] = merged_df.apply(determine_real_app, axis=1)
            
            # Cleanup columns
            cols_to_drop = ['Name', 'Notes', 'OS Version', 'DNS Name']
            merged_df.drop([col for col in cols_to_drop if col in merged_df.columns], 
                         axis=1, 
                         inplace=True)
            
            # Export
            self.update_status("Generating final output...")
            output_path = self.iteration1_path.replace('.xlsx', '_crossreferenced.xlsx')
            merged_df.to_excel(output_path, index=False, engine='openpyxl')
            
            messagebox.showinfo("Success", f"Enhanced file saved:\n{output_path}")
            self.update_status("Processing complete - ready for next task")
            
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
            self.update_status("Error occurred - check input files")
        finally:
            self.process_btn.config(state='normal' if self.iteration0_path and self.iteration1_path else 'disabled')
            self.update_status("Ready")

    def _validate_columns(self, df, file_type, required_columns):
        """Flexible column validation with suggestions."""
        normalized_columns = {col.strip().lower().replace(' ', '').replace('_', '') 
                            for col in df.columns}
        missing = []
        
        for col in required_columns:
            normalized_col = col.strip().lower().replace(' ', '').replace('_', '')
            if normalized_col not in normalized_columns:
                # Look for partial matches
                similar = [orig for orig in df.columns 
                         if normalized_col in orig.strip().lower().replace(' ', '').replace('_', '')]
                if similar:
                    raise ValueError(
                        f"{file_type} missing '{col}' column. Similar columns found: {', '.join(similar)}\n"
                        f"Please rename to '{col}' or update the code."
                    )
                missing.append(col)
        
        if missing:
            req_list = "\n- ".join(required_columns)
            raise ValueError(
                f"{file_type} missing required columns: {', '.join(missing)}\n"
                f"Required columns for {file_type}:\n- {req_list}"
            )

    def show_about(self):
        about_text = f"""VM Intelligence Merger v{VERSION}

Developed by David Maiolo

Key Features:
- Multi-source intelligence analysis (Notes, Ports, OS, DNS)
- Automated RITM number removal
- Recognized application handling
- Port-to-application mapping
- OS pattern matching
- Fallback analysis with manual review flags

Handles ambiguous cases through:
1. Direct Notes analysis
2. Virtana Recognized apps
3. Network port signatures
4. Operating System patterns
5. DNS naming conventions"""
        messagebox.showinfo("About", about_text)

    def show_help(self):
        help_window = tk.Toplevel(self)
        help_window.title("Comprehensive User Guide")
        help_window.geometry("600x500")
        
        help_text = ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        help_text.pack(fill='both', expand=True)
        
        instructions = """Advanced Intelligence Merging Process

1. File Selection:
   - Iteration 0: Raw CSV export with VM metadata
   - Iteration 1: Processed XLSX with initial analysis

2. Column Requirements:
   CSV (Iteration 0) must contain:
   - Name: VM hostname
   - Notes: Application references
   - OS Version: Operating system info
   - DNS Name: DNS identifier

   XLSX (Iteration 1) must contain:
   - Host: Full VM identifier
   - Discovered App: Virtana's analysis
   - Feature Ports: Network port list

3. Intelligence Hierarchy:
   a) Direct Notes entries (with RITM cleanup)
   b) Virtana Recognized applications
   c) Network port signature analysis
   d) Operating System patterns
   e) DNS name keyword matching
   f) Final manual review flag

4. Port Mapping Includes:
   - 53: DNS Server
   - 80/443: Web Services
   - 22: SSH
   - 3306: MySQL
   - 1433: SQL Server
   - 5007: Palo Alto
   - And 20+ others

5. Best Practices:
   - Verify high-value systems manually
   - Review all 'UNKNOWN' flagged entries
   - Rename columns to match exactly if needed
   - Validate against CMDB records periodically

For column name mismatches:
- Rename columns in your files to match exactly
- Or modify the code's required_columns lists"""
        
        help_text.insert(tk.INSERT, instructions)
        help_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    app = Application()
    app.mainloop()