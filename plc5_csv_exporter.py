import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import win32com.client as win32com
import csv
import os
import re
import threading
from datetime import datetime


class PLC5CSVExporter:
    def __init__(self, root):
        self.root = root
        self.root.title("PLC-5 RSP to CSV Exporter")
        self.root.geometry("700x500")
        
        self.rsp_file = None
        self.output_folder = None
        self.is_processing = False
        
        self.setup_ui()
    
    def setup_ui(self):
        # File selection frame
        file_frame = ttk.LabelFrame(self.root, text="File Selection", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(file_frame, text="RSP File:").grid(row=0, column=0, sticky="w", pady=5)
        self.file_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky="w", padx=10)
        ttk.Button(file_frame, text="Browse...", command=self.browse_rsp).grid(row=0, column=2, padx=5)
        
        ttk.Label(file_frame, text="Output Folder:").grid(row=1, column=0, sticky="w", pady=5)
        self.output_label = ttk.Label(file_frame, text="Same as RSP file", foreground="gray")
        self.output_label.grid(row=1, column=1, sticky="w", padx=10)
        ttk.Button(file_frame, text="Browse...", command=self.browse_output).grid(row=1, column=2, padx=5)
        
        # Export options
        options_frame = ttk.LabelFrame(self.root, text="Export Options", padding=10)
        options_frame.pack(fill="x", padx=10, pady=5)
        
        self.export_tags = tk.BooleanVar(value=True)
        self.export_timers = tk.BooleanVar(value=True)
        self.export_counters = tk.BooleanVar(value=True)
        self.export_controls = tk.BooleanVar(value=True)
        self.export_arrays = tk.BooleanVar(value=True)
        self.export_messages = tk.BooleanVar(value=True)
        self.export_io = tk.BooleanVar(value=True)
        self.export_rungs = tk.BooleanVar(value=True)
        self.export_datatable = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(options_frame, text="Tags/Addresses", variable=self.export_tags).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(options_frame, text="Timers", variable=self.export_timers).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(options_frame, text="Counters", variable=self.export_counters).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(options_frame, text="Controls", variable=self.export_controls).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(options_frame, text="Arrays", variable=self.export_arrays).grid(row=1, column=1, sticky="w")
        ttk.Checkbutton(options_frame, text="Messages", variable=self.export_messages).grid(row=1, column=2, sticky="w")
        ttk.Checkbutton(options_frame, text="I/O Points", variable=self.export_io).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(options_frame, text="Ladder Rungs", variable=self.export_rungs).grid(row=2, column=1, sticky="w")
        ttk.Checkbutton(options_frame, text="Data Table (slow)", variable=self.export_datatable).grid(row=2, column=2, sticky="w")
        
        # Progress frame
        progress_frame = ttk.LabelFrame(self.root, text="Progress", padding=10)
        progress_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.pack(fill="x", pady=5)
        
        self.log_text = tk.Text(progress_frame, height=10, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Action buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        self.export_btn = ttk.Button(button_frame, text="Export to CSV", command=self.start_export, state="disabled")
        self.export_btn.pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side="right", padx=5)
    
    def browse_rsp(self):
        filename = filedialog.askopenfilename(
            title="Select PLC-5 RSP File",
            filetypes=[("RSLogix 5 Files", "*.rsp"), ("All Files", "*.*")]
        )
        if filename:
            self.rsp_file = filename
            self.file_label.config(text=os.path.basename(filename), foreground="black")
            if not self.output_folder:
                self.output_folder = os.path.dirname(filename)
                self.output_label.config(text=self.output_folder, foreground="black")
            self.export_btn.config(state="normal")
            self.log(f"Selected file: {filename}")
    
    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_label.config(text=folder, foreground="black")
            self.log(f"Output folder: {folder}")
    
    def log(self, message):
        self.log_text.config(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update()
    
    def start_export(self):
        if self.is_processing:
            return
        
        if not self.rsp_file:
            messagebox.showerror("Error", "Please select an RSP file")
            return
        
        self.is_processing = True
        self.export_btn.config(state="disabled")
        self.progress_bar.start()
        
        thread = threading.Thread(target=self.export_data, daemon=True)
        thread.start()
    
    def export_data(self):
        try:
            self.log("Opening RSLogix5 Application...")
            rslogix5 = win32com.Dispatch("RSLogix5.Application.5")
            rslogix5.visible = True
            
            abs_path = os.path.abspath(self.rsp_file)
            self.log(f"Opening project: {abs_path}")
            project = rslogix5.FileOpen(abs_path, False, False, True)
            
            program_files = project.ProgramFiles
            addr_sym_records = project.AddrSymRecords
            datafiles = project.DataFiles
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.rsp_file))[0]
            
            # Export Timers, Counters, Tags, I/O, etc. by analyzing rungs
            # NOTE: Tags and I/O are now collected DURING ladder analysis
            if any([self.export_timers.get(), self.export_counters.get(), 
                   self.export_controls.get(), self.export_arrays.get(), 
                   self.export_messages.get(), self.export_tags.get(), self.export_io.get()]):
                self.log("Analyzing ladder logic...")
                try:
                    self.analyze_ladder_logic(program_files, addr_sym_records, datafiles, timestamp, base_name)
                except Exception as e:
                    self.log(f"  ERROR in ladder analysis: {str(e)}")
                    import traceback
                    self.log(f"  {traceback.format_exc()}")
            
            # Export Ladder Rungs
            if self.export_rungs.get():
                self.log("Extracting ladder rungs...")
                try:
                    self.export_rungs_to_csv(program_files, timestamp, base_name)
                except Exception as e:
                    self.log(f"  ERROR in rungs export: {str(e)}")
                    import traceback
                    self.log(f"  {traceback.format_exc()}")
            
            # Export Data Table
            if self.export_datatable.get():
                self.log("Extracting data table values...")
                self.export_datatable_to_csv(datafiles, timestamp, base_name)
            
            self.log("Closing project...")
            rslogix5.Quit(True, False)
            
            self.log("=" * 50)
            self.log(f"Export completed successfully!")
            self.log(f"Files saved to: {self.output_folder}")
            
            messagebox.showinfo("Success", "Export completed successfully!")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.log(f"ERROR: {str(e)}")
            self.log(f"Details: {error_details}")
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")
        finally:
            self.is_processing = False
            self.progress_bar.stop()
            self.export_btn.config(state="normal")
    
    def export_tags_to_csv(self, addr_sym_records, datafiles, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Tags_{timestamp}.csv")
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['PLC5_Address', 'Symbol', 'Description', 'DataType', 'Value'])
            
            count = 0
            # Try to get count safely
            try:
                record_count = addr_sym_records.Count()
                self.log(f"  Found {record_count} address records")
            except TypeError:
                # Count might be property, not method
                record_count = addr_sym_records.Count
                self.log(f"  Found {record_count} address records (property access)")
            
            for i in range(record_count):
                try:
                    # Use .Item() for addr_sym_records like working code
                    record = addr_sym_records.Item(i)
                    if record:
                        try:
                            address = record.Address
                            symbol = record.Symbol if record.Symbol else ""
                            desc = record.Description.replace('\r\n', ' | ') if record.Description else ""
                        except:
                            # Skip records with inaccessible properties
                            continue
                        
                        try:
                            value = datafiles.GetDataValue(address)
                        except:
                            value = ""
                        
                        data_type = self.get_data_type(address)
                        
                        writer.writerow([address, symbol, desc, data_type, value])
                        count += 1
                except Exception as e:
                    # Skip problematic records silently
                    continue
            
            self.log(f"  Exported {count} tags to {os.path.basename(csv_file)}")
    
    def analyze_ladder_logic(self, program_files, addr_sym_records, datafiles, timestamp, base_name):
        timers = {}
        counters = {}
        controls = {}
        arrays = []
        messages = {}
        all_addresses = {}  # Collect all addresses found in ladder
        io_addresses = {}   # Collect I/O addresses separately
        
        # Use Count() as method like working code
        for file_idx in range(2, program_files.Count()):
            try:
                # Use callable syntax like working code: object(index)
                ladder_file = program_files(file_idx)
                if ladder_file and ladder_file.NumberOfRungs() > 0:
                    self.log(f"  Analyzing file: {ladder_file.Name}")
                    
                    rung_count = ladder_file.NumberOfRungs()
                    for rung_idx in range(rung_count):
                        try:
                            rung_ascii = ladder_file.GetRungAsAscii(rung_idx)
                            
                            # Extract all addresses from this rung
                            self.extract_addresses_from_rung(rung_ascii, all_addresses, io_addresses, 
                                                            addr_sym_records, datafiles)
                            
                            # Extract timers
                            if self.export_timers.get() and any(x in rung_ascii for x in ['TON ', 'TOF ', 'RTO ']):
                                self.extract_timers(rung_ascii, timers, addr_sym_records, datafiles)
                            
                            # Extract counters
                            if self.export_counters.get() and any(x in rung_ascii for x in ['CTU ', 'CTD ']):
                                self.extract_counters(rung_ascii, counters, addr_sym_records, datafiles)
                            
                            # Extract arrays and controls
                            if self.export_arrays.get() or self.export_controls.get():
                                self.extract_arrays_controls(rung_ascii, arrays, controls, addr_sym_records)
                            
                            # Extract messages
                            if self.export_messages.get() and 'MSG ' in rung_ascii:
                                self.extract_messages(rung_ascii, messages, addr_sym_records)
                        except Exception as e:
                            self.log(f"  Warning: Error in rung {rung_idx}: {str(e)}")
                            continue
            except Exception as e:
                self.log(f"  Warning: Error in file {file_idx}: {str(e)}")
                continue
        
        # Write CSVs
        if self.export_timers.get() and timers:
            self.write_timers_csv(timers, timestamp, base_name)
        
        if self.export_counters.get() and counters:
            self.write_counters_csv(counters, timestamp, base_name)
        
        if self.export_controls.get() and controls:
            self.write_controls_csv(controls, timestamp, base_name)
        
        if self.export_arrays.get() and arrays:
            self.write_arrays_csv(arrays, timestamp, base_name)
        
        if self.export_messages.get() and messages:
            self.write_messages_csv(messages, timestamp, base_name)
        
        # Write Tags and I/O from collected addresses
        if self.export_tags.get() and all_addresses:
            self.write_tags_from_addresses(all_addresses, timestamp, base_name)
        
        if self.export_io.get() and io_addresses:
            self.write_io_from_addresses(io_addresses, timestamp, base_name)
    
    def extract_addresses_from_rung(self, rung_ascii, all_addresses, io_addresses, addr_sym_records, datafiles):
        """Extract all PLC addresses from a rung and store them with their info"""
        # Pattern to match PLC-5 addresses: I:, O:, B, N, F, T, C, R, etc.
        address_pattern = re.compile(r'\b([IONBFLTCRS]:\d+(?:/\d+)?|[BNTCR]\d+:\d+(?:/\d+)?)\b')
        matches = address_pattern.findall(rung_ascii)
        
        for address in matches:
            # Skip if already processed
            if address in all_addresses or address in io_addresses:
                continue
            
            # Get symbol and description using GetRecordViaAddrOrSym (the working method)
            symbol = ""
            desc = ""
            try:
                record = addr_sym_records.GetRecordViaAddrOrSym(address, 0)
                if record:
                    symbol = record.Symbol if record.Symbol else ""
                    desc = record.Description.replace('\r\n', ' | ') if record.Description else ""
            except:
                pass
            
            # Get value
            value = ""
            try:
                value = datafiles.GetDataValue(address)
            except:
                pass
            
            data_type = self.get_data_type(address)
            
            # Categorize as I/O or general address
            if address.startswith('I:'):
                io_addresses[address] = {
                    'Type': 'Input',
                    'Address': address,
                    'Symbol': symbol,
                    'Description': desc,
                    'Value': value
                }
            elif address.startswith('O:'):
                io_addresses[address] = {
                    'Type': 'Output',
                    'Address': address,
                    'Symbol': symbol,
                    'Description': desc,
                    'Value': value
                }
            else:
                all_addresses[address] = {
                    'Address': address,
                    'Symbol': symbol,
                    'Description': desc,
                    'DataType': data_type,
                    'Value': value
                }
    
    def extract_timers(self, rung, timers, addr_sym_records, datafiles):
        timer_pattern = re.compile(r'(TON|TOF|RTO)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)')
        matches = timer_pattern.findall(rung)
        
        for match in matches:
            timer_type, address, base, pre, acc = match
            if address not in timers:
                symbol, desc = self.get_symbol_desc(address, addr_sym_records)
                timers[address] = {
                    'Type': timer_type,
                    'Address': address,
                    'Symbol': symbol,
                    'Description': desc,
                    'Base': base,
                    'PRE': pre,
                    'ACC': acc
                }
    
    def extract_counters(self, rung, counters, addr_sym_records, datafiles):
        counter_pattern = re.compile(r'(CTU|CTD)\s+(\S+)\s+(\S+)\s+(\S+)')
        matches = counter_pattern.findall(rung)
        
        for match in matches:
            counter_type, address, pre, acc = match
            if address not in counters:
                symbol, desc = self.get_symbol_desc(address, addr_sym_records)
                counters[address] = {
                    'Type': counter_type,
                    'Address': address,
                    'Symbol': symbol,
                    'Description': desc,
                    'PRE': pre,
                    'ACC': acc
                }
    
    def extract_arrays_controls(self, rung, arrays, controls, addr_sym_records):
        control_pattern = re.compile(r'(FAL|FSC|FFL|FFU|COP|DDT|FBC)\s+(\S+)\s+(\S+)\s+(\S+)')
        matches = control_pattern.findall(rung)
        
        for match in matches:
            inst, control, length, pos = match[:4]
            if control not in controls:
                symbol, desc = self.get_symbol_desc(control, addr_sym_records)
                controls[control] = {
                    'Instruction': inst,
                    'Address': control,
                    'Symbol': symbol,
                    'Description': desc,
                    'Length': length,
                    'Position': pos
                }
    
    def extract_messages(self, rung, messages, addr_sym_records):
        msg_pattern = re.compile(r'MSG\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)')
        matches = msg_pattern.findall(rung)
        
        for match in matches:
            if len(match) >= 7:
                address = match[0]
                if address not in messages:
                    symbol, desc = self.get_symbol_desc(address, addr_sym_records)
                    messages[address] = {
                        'Address': address,
                        'Symbol': symbol,
                        'Description': desc,
                        'Type': match[1],
                        'ThisPLC': match[2],
                        'Length': match[3],
                        'Port': match[4],
                        'Target': match[5],
                        'Node': match[6]
                    }
    
    def write_timers_csv(self, timers, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Timers_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Type', 'Address', 'Symbol', 'Description', 'Base', 'PRE', 'ACC'])
            for timer in timers.values():
                writer.writerow([timer['Type'], timer['Address'], timer['Symbol'], 
                               timer['Description'], timer['Base'], timer['PRE'], timer['ACC']])
        self.log(f"  Exported {len(timers)} timers to {os.path.basename(csv_file)}")
    
    def write_counters_csv(self, counters, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Counters_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Type', 'Address', 'Symbol', 'Description', 'PRE', 'ACC'])
            for counter in counters.values():
                writer.writerow([counter['Type'], counter['Address'], counter['Symbol'], 
                               counter['Description'], counter['PRE'], counter['ACC']])
        self.log(f"  Exported {len(counters)} counters to {os.path.basename(csv_file)}")
    
    def write_controls_csv(self, controls, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Controls_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Instruction', 'Address', 'Symbol', 'Description', 'Length', 'Position'])
            for control in controls.values():
                writer.writerow([control['Instruction'], control['Address'], control['Symbol'], 
                               control['Description'], control['Length'], control['Position']])
        self.log(f"  Exported {len(controls)} controls to {os.path.basename(csv_file)}")
    
    def write_arrays_csv(self, arrays, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Arrays_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Address', 'Length', 'Instruction', 'Symbol', 'Description'])
            for array in arrays:
                writer.writerow([array['Address'], array['Length'], array['Instruction'], 
                               array['Symbol'], array['Description']])
        self.log(f"  Exported {len(arrays)} arrays to {os.path.basename(csv_file)}")
    
    def write_messages_csv(self, messages, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Messages_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Address', 'Symbol', 'Description', 'Type', 'ThisPLC', 'Length', 'Port', 'Target', 'Node'])
            for msg in messages.values():
                writer.writerow([msg['Address'], msg['Symbol'], msg['Description'], msg['Type'], 
                               msg['ThisPLC'], msg['Length'], msg['Port'], msg['Target'], msg['Node']])
        self.log(f"  Exported {len(messages)} messages to {os.path.basename(csv_file)}")
    
    def write_tags_from_addresses(self, all_addresses, timestamp, base_name):
        """Write tags collected during ladder analysis"""
        csv_file = os.path.join(self.output_folder, f"{base_name}_Tags_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['PLC5_Address', 'Symbol', 'Description', 'DataType', 'Value'])
            for tag in all_addresses.values():
                writer.writerow([tag['Address'], tag['Symbol'], tag['Description'], 
                               tag['DataType'], tag['Value']])
        self.log(f"  Exported {len(all_addresses)} tags to {os.path.basename(csv_file)}")
    
    def write_io_from_addresses(self, io_addresses, timestamp, base_name):
        """Write I/O points collected during ladder analysis"""
        csv_file = os.path.join(self.output_folder, f"{base_name}_IO_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Type', 'Address', 'Symbol', 'Description', 'Value'])
            for io in io_addresses.values():
                writer.writerow([io['Type'], io['Address'], io['Symbol'], 
                               io['Description'], io['Value']])
        self.log(f"  Exported {len(io_addresses)} I/O points to {os.path.basename(csv_file)}")
    
    def export_io_to_csv(self, addr_sym_records, datafiles, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_IO_{timestamp}.csv")
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Type', 'Address', 'Symbol', 'Description', 'Value'])
            
            count = 0
            # Use Count() as method like working code
            for i in range(addr_sym_records.Count()):
                try:
                    record = addr_sym_records.Item(i)
                    if record:
                        address = record.Address
                        if address.startswith('I:') or address.startswith('O:'):
                            symbol = record.Symbol if record.Symbol else ""
                            desc = record.Description.replace('\r\n', ' | ') if record.Description else ""
                            
                            try:
                                value = datafiles.GetDataValue(address)
                            except:
                                value = ""
                            
                            io_type = 'Input' if address.startswith('I:') else 'Output'
                            writer.writerow([io_type, address, symbol, desc, value])
                            count += 1
                except Exception as e:
                    continue
            
            self.log(f"  Exported {count} I/O points to {os.path.basename(csv_file)}")
    
    def export_rungs_to_csv(self, program_files, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_Rungs_{timestamp}.csv")
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['File_Name', 'File_Number', 'Rung_Number', 'Rung_ASCII'])
            
            total_rungs = 0
            # Use Count() as method like working code
            for file_idx in range(2, program_files.Count()):
                try:
                    ladder_file = program_files.Item(file_idx)
                    if ladder_file and ladder_file.NumberOfRungs() > 0:
                        file_name = ladder_file.Name
                        file_number = ladder_file.FileNumber
                        
                        rung_count = ladder_file.NumberOfRungs()
                        for rung_idx in range(rung_count):
                            try:
                                rung_ascii = ladder_file.GetRungAsAscii(rung_idx)
                                writer.writerow([file_name, file_number, rung_idx, rung_ascii])
                                total_rungs += 1
                            except Exception as e:
                                self.log(f"  Warning: Error reading rung {rung_idx} in file {file_name}: {str(e)}")
                                continue
                except Exception as e:
                    self.log(f"  Warning: Error processing file {file_idx}: {str(e)}")
                    continue
            
            self.log(f"  Exported {total_rungs} ladder rungs to {os.path.basename(csv_file)}")
    
    def export_datatable_to_csv(self, datafiles, timestamp, base_name):
        csv_file = os.path.join(self.output_folder, f"{base_name}_DataTable_{timestamp}.csv")
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['FileType', 'FileNumber', 'Element', 'Address', 'Value'])
            
            count = 0
            # Use Count() as method like working code
            for file_idx in range(datafiles.Count()):
                try:
                    # Use callable syntax like working code: object(index)
                    datafile = datafiles(file_idx)
                    if datafile:
                        file_type = datafile.TypeAsString
                        file_num = datafile.FileNumber
                        length = datafile.NumberOfElements
                        
                        if file_type in ['B', 'N', 'F', 'L', 'T', 'C', 'R']:
                            for elem in range(min(length, 1000)):
                                address = f"{file_type}{file_num}:{elem}"
                                try:
                                    value = datafiles.GetDataValue(address)
                                    writer.writerow([file_type, file_num, elem, address, value])
                                    count += 1
                                except:
                                    pass
                except Exception as e:
                    continue
            
            self.log(f"  Exported {count} data table values to {os.path.basename(csv_file)}")
    
    def get_symbol_desc(self, address, addr_sym_records):
        try:
            record = addr_sym_records.GetRecordViaAddrOrSym(address, 0)
            if record:
                symbol = record.Symbol if record.Symbol else ""
                desc = record.Description.replace('\r\n', ' | ') if record.Description else ""
                return symbol, desc
        except:
            pass
        return "", ""
    
    def get_data_type(self, address):
        if ':' not in address:
            return "Unknown"
        
        prefix = address.split(':')[0].replace('#', '')
        
        type_map = {
            'I': 'Input',
            'O': 'Output',
            'B': 'Bit',
            'N': 'Integer',
            'F': 'Float',
            'L': 'Long',
            'T': 'Timer',
            'C': 'Counter',
            'R': 'Control',
            'S': 'Status'
        }
        
        return type_map.get(prefix, prefix)


if __name__ == "__main__":
    root = tk.Tk()
    app = PLC5CSVExporter(root)
    root.mainloop()
