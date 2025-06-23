import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import pdfplumber
import os
import threading
from pathlib import Path
from datetime import datetime
from queue import Queue, Empty
import re

class PDFToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Meta Ads PDF to CSV Converter")
        self.root.geometry("1000x800")
        self.root.minsize(800, 650)
        
        # Configure modern style
        self.setup_style()

        self.pdf_files = []
        self.csv_output_path = tk.StringVar()

        self.columns = [
            "Sort row", "Source", "Posted date", "Merch (Ref #)", "Transaction ID",
            "Total per Campaign", "Campaign name", "Company", "Branch", "Account",
            "Event Code", "Client (if applicable)", "Description"
        ]

        self.message_queue = Queue()
        self.cancel_processing = False
        self.processing_thread = None

        self.setup_ui()
        self.check_queue()

    def setup_style(self):
        """Configure modern, professional styling"""
        style = ttk.Style()
        
        # Configure colors and fonts
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#2C3E50')
        style.configure('Header.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#34495E')
        style.configure('Info.TLabel', font=('Segoe UI', 9), foreground='#7F8C8D')
        style.configure('Success.TLabel', font=('Segoe UI', 9), foreground='#27AE60')
        style.configure('Error.TLabel', font=('Segoe UI', 9), foreground='#E74C3C')
        
        # Configure button styles
        style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('Secondary.TButton', font=('Segoe UI', 10))
        
        # Configure entry styles
        style.configure('Modern.TEntry', fieldbackground='white', borderwidth=1, relief='solid')
        
        # Configure frame styles
        style.configure('Card.TFrame', relief='raised', borderwidth=1)
        style.configure('Main.TFrame', padding=20)

    def setup_ui(self):
        # Main container with padding
        main_container = ttk.Frame(self.root, style='Main.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Title section
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill=tk.X, pady=(0, 30))
        
        title_label = ttk.Label(title_frame, text="Meta Ads PDF to CSV Converter", style='Title.TLabel')
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(title_frame, text="Convert multiple Meta Ads PDF receipts to a single CSV file", style='Info.TLabel')
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_container, text="File Selection", padding=20)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # PDF files selection
        pdf_header_frame = ttk.Frame(file_frame)
        pdf_header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(pdf_header_frame, text="PDF Files:", style='Header.TLabel').pack(side=tk.LEFT)
        
        pdf_buttons_frame = ttk.Frame(pdf_header_frame)
        pdf_buttons_frame.pack(side=tk.RIGHT)
        
        ttk.Button(pdf_buttons_frame, text="Add PDFs", command=self.add_pdfs, style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(pdf_buttons_frame, text="Clear All", command=self.clear_pdfs, style='Secondary.TButton').pack(side=tk.LEFT)
        
        # PDF files listbox with scrollbar
        listbox_frame = ttk.Frame(file_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Create listbox with scrollbar
        self.pdf_listbox = tk.Listbox(listbox_frame, font=('Segoe UI', 9), selectmode=tk.EXTENDED)
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.pdf_listbox.yview)
        self.pdf_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.pdf_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Remove selected button
        ttk.Button(file_frame, text="Remove Selected", command=self.remove_selected_pdfs, style='Secondary.TButton').pack(pady=(0, 15))
        
        # CSV output selection
        ttk.Label(file_frame, text="Output CSV File:", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 5))
        
        csv_frame = ttk.Frame(file_frame)
        csv_frame.pack(fill=tk.X)
        csv_frame.columnconfigure(0, weight=1)
        
        self.csv_entry = ttk.Entry(csv_frame, textvariable=self.csv_output_path, style='Modern.TEntry', font=('Segoe UI', 10))
        self.csv_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(csv_frame, text="Browse CSV", command=self.browse_csv, style='Secondary.TButton').grid(row=0, column=1)
        
        # Control buttons section
        control_frame = ttk.Frame(main_container)
        control_frame.pack(fill=tk.X, pady=(0, 20))
        
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(anchor=tk.CENTER)
        
        self.convert_btn = ttk.Button(button_frame, text="Convert All PDFs to CSV", command=self.convert, style='Primary.TButton')
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.cancel, style='Secondary.TButton', state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT)
        
        # Status section
        status_frame = ttk.LabelFrame(main_container, text="Status", padding=15)
        status_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.status_var = tk.StringVar(value="Ready to convert - Add PDF files to begin")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, style='Info.TLabel', font=('Segoe UI', 10))
        self.status_label.pack(anchor=tk.W)
        
        # Progress bar
        self.progress = ttk.Progressbar(status_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(10, 0))
        
        # Log section
        log_frame = ttk.LabelFrame(main_container, text="Processing Log", padding=15)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=12, 
            font=('Consolas', 9),
            bg='#F8F9FA',
            fg='#2C3E50',
            insertbackground='#2C3E50'
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def add_pdfs(self):
        """Add PDF files to the list"""
        filenames = filedialog.askopenfilenames(
            title="Select Meta Ads PDF Receipts",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if filenames:
            added_count = 0
            for filename in filenames:
                if filename not in self.pdf_files:
                    self.pdf_files.append(filename)
                    self.pdf_listbox.insert(tk.END, Path(filename).name)
                    added_count += 1
            
            if added_count > 0:
                self.log(f"Added {added_count} PDF file(s)")
                self.update_status()
                
                # Auto-generate CSV output path if not set
                if not self.csv_output_path.get() and self.pdf_files:
                    first_pdf = Path(self.pdf_files[0])
                    default_csv = first_pdf.parent / f"{first_pdf.stem}_combined.csv"
                    self.csv_output_path.set(str(default_csv))

    def remove_selected_pdfs(self):
        """Remove selected PDF files from the list"""
        selected_indices = self.pdf_listbox.curselection()
        if selected_indices:
            # Remove in reverse order to maintain indices
            for index in reversed(selected_indices):
                del self.pdf_files[index]
                self.pdf_listbox.delete(index)
            
            self.log(f"Removed {len(selected_indices)} PDF file(s)")
            self.update_status()

    def clear_pdfs(self):
        """Clear all PDF files from the list"""
        if self.pdf_files:
            count = len(self.pdf_files)
            self.pdf_files.clear()
            self.pdf_listbox.delete(0, tk.END)
            self.log(f"Cleared all {count} PDF files")
            self.update_status()

    def update_status(self):
        """Update the status message based on current state"""
        file_count = len(self.pdf_files)
        if file_count == 0:
            self.status_var.set("Ready to convert - Add PDF files to begin")
        elif file_count == 1:
            self.status_var.set(f"Ready to convert 1 PDF file")
        else:
            self.status_var.set(f"Ready to convert {file_count} PDF files")

    def browse_csv(self):
        """Browse for CSV output location"""
        filename = filedialog.asksaveasfilename(
            title="Save CSV File As",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_output_path.set(filename)
            self.log(f"Output CSV: {Path(filename).name}")

    def check_queue(self):
        """Check for messages from processing thread"""
        try:
            while True:
                msg = self.message_queue.get_nowait()
                if msg['type'] == 'log':
                    self.log(msg['text'])
                elif msg['type'] == 'status':
                    self.status_var.set(msg['text'])
                    if 'error' in msg['text'].lower():
                        self.status_label.configure(style='Error.TLabel')
                    elif 'success' in msg['text'].lower() or 'completed' in msg['text'].lower():
                        self.status_label.configure(style='Success.TLabel')
                    else:
                        self.status_label.configure(style='Info.TLabel')
                elif msg['type'] == 'progress':
                    self.progress['value'] = msg['value']
                elif msg['type'] == 'progress_max':
                    self.progress['maximum'] = msg['value']
                elif msg['type'] == 'done':
                    self.convert_btn.config(state=tk.NORMAL)
                    self.cancel_btn.config(state=tk.DISABLED)
                    self.progress['value'] = 0
        except Empty:
            pass
        self.root.after(100, self.check_queue)

    def log(self, message):
        """Add message to log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def cancel(self):
        """Cancel processing"""
        self.cancel_processing = True
        self.message_queue.put({'type': 'status', 'text': "Cancelling operation..."})
        self.log("Cancellation requested by user")

    def convert(self):
        """Start conversion process"""
        if not self.pdf_files:
            messagebox.showerror("Error", "Please add at least one PDF file")
            return
            
        csv_path = self.csv_output_path.get().strip()
        if not csv_path:
            messagebox.showerror("Error", "Please specify an output CSV file")
            return

        # Validate PDF files exist
        missing_files = [f for f in self.pdf_files if not os.path.exists(f)]
        if missing_files:
            messagebox.showerror("Error", f"PDF files not found:\n" + "\n".join(Path(f).name for f in missing_files[:5]))
            return

        # Reset UI state
        self.cancel_processing = False
        self.convert_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.status_label.configure(style='Info.TLabel')
        self.progress['value'] = 0

        # Start processing
        self.log(f"Starting conversion of {len(self.pdf_files)} PDF files...")
        self.message_queue.put({'type': 'status', 'text': f"Processing {len(self.pdf_files)} PDF files..."})
        self.message_queue.put({'type': 'progress_max', 'value': len(self.pdf_files)})

        self.processing_thread = threading.Thread(
            target=self.run_conversion, 
            args=(self.pdf_files.copy(), csv_path), 
            daemon=True
        )
        self.processing_thread.start()

    def run_conversion(self, pdf_files, csv_path):
        """Run the conversion process in a separate thread"""
        all_data = []
        sort_row = 1
        
        try:
            for i, pdf_path in enumerate(pdf_files, 1):
                if self.cancel_processing:
                    self.message_queue.put({'type': 'status', 'text': "Operation cancelled"})
                    self.message_queue.put({'type': 'done'})
                    return
                
                self.message_queue.put({'type': 'log', 'text': f"Processing file {i}/{len(pdf_files)}: {Path(pdf_path).name}"})
                self.message_queue.put({'type': 'progress', 'value': i})
                
                try:
                    # Extract text from current PDF
                    text = self.extract_text_from_pdf(pdf_path)
                    
                    if self.cancel_processing:
                        break
                    
                    # Extract campaign data from current PDF
                    df = self.extract_campaign_data(text, sort_row)
                    
                    if not df.empty:
                        all_data.append(df)
                        sort_row += len(df)
                        self.message_queue.put({'type': 'log', 'text': f"  Found {len(df)} campaigns in {Path(pdf_path).name}"})
                    else:
                        self.message_queue.put({'type': 'log', 'text': f"  No campaigns found in {Path(pdf_path).name}"})
                        
                except Exception as e:
                    self.message_queue.put({'type': 'log', 'text': f"  Error processing {Path(pdf_path).name}: {str(e)}"})
                    continue
            
            if self.cancel_processing:
                self.message_queue.put({'type': 'status', 'text': "Operation cancelled"})
                self.message_queue.put({'type': 'done'})
                return
            
            # Combine all data
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                
                self.message_queue.put({'type': 'log', 'text': f"Saving {len(combined_df)} total campaigns to CSV..."})
                
                # Save to CSV
                combined_df.to_csv(csv_path, index=False, encoding='utf-8')
                
                self.message_queue.put({'type': 'log', 'text': "Conversion completed successfully!"})
                self.message_queue.put({'type': 'status', 'text': f"Success! Saved {len(combined_df)} campaigns from {len(pdf_files)} PDFs to CSV"})
                
            else:
                self.message_queue.put({'type': 'log', 'text': "No campaign data found in any PDF files"})
                self.message_queue.put({'type': 'status', 'text': "No data extracted from PDF files"})
                
        except Exception as e:
            error_msg = f"Error during conversion: {str(e)}"
            self.message_queue.put({'type': 'log', 'text': error_msg})
            self.message_queue.put({'type': 'status', 'text': "Conversion failed"})
        
        finally:
            self.message_queue.put({'type': 'done'})

    def extract_text_from_pdf(self, filepath):
        """Extract text from PDF file"""
        try:
            with pdfplumber.open(filepath) as pdf:
                text_parts = []
                for page in pdf.pages:
                    if self.cancel_processing:
                        break
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
                
                return "\n".join(text_parts)
        except Exception as e:
            raise Exception(f"Failed to extract text from PDF: {str(e)}")

    def extract_campaign_data(self, text, start_sort_row):
        """Extract campaign data from PDF text - campaign total is the line above 'From'"""
        extracted_data = []
        sort_row = start_sort_row
        
        # Extract header information
        source = ""
        posted_date = ""
        ref_number = ""
        transaction_id = ""
        account_id = ""

        lines = text.split('\n')
        header_lines = lines[:25]

        # Extract header information
        for idx, line in enumerate(header_lines):
            if self.cancel_processing:
                break
                
            line = line.strip()
            
            # Extract source
            if not source and line.startswith("Receipt for"):
                match = re.search(r"Receipt for (.+)", line)
                if match:
                    source = match.group(1).strip()

            # Extract posted date
            if not posted_date and "Invoice/Payment Date" in line and idx + 1 < len(header_lines):
                next_line = header_lines[idx + 1].strip()
                date_match = re.search(r'([A-Za-z]+ \d{1,2}, \d{4}(?:,? \d{1,2}:\d{2} [APMapm]{2})?)', next_line)
                if date_match:
                    posted_date = date_match.group(1).strip()

            # Extract reference number
            if not ref_number:
                ref_match = re.search(r'Reference Number\s*:?\s*([A-Z0-9\-]+)', line)
                if ref_match:
                    ref_number = ref_match.group(1).strip()

            # Extract transaction ID
            if not transaction_id and "Transaction ID" in line and idx + 1 < len(header_lines):
                next_line = header_lines[idx + 1].strip()
                trans_match = re.search(r'([0-9\-]{15,})', next_line)
                if trans_match:
                    transaction_id = trans_match.group(1).strip()

            # Extract account ID
            if not account_id:
                acc_match = re.search(r'Account ID\s*:?\s*([0-9]+)', line)
                if acc_match:
                    account_id = acc_match.group(1).strip()

        # Extract campaign information
        for i in range(len(lines)):
            if self.cancel_processing:
                break
                
            line = lines[i].strip()

            # Look for lines that start with "From" (date range)
            if line.lower().startswith("from") and i >= 2:
                # Campaign name is 2 lines above "From"
                campaign_name_line = lines[i - 2].strip() if i >= 2 else ""
                # Campaign total is 1 line above "From" (the line directly above)
                campaign_total_line = lines[i - 1].strip() if i >= 1 else ""
                
                # Extract the dollar amount from the total line
                total_match = re.search(r'^\$(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)$', campaign_total_line)
                
                if campaign_name_line and total_match:
                    campaign_total = total_match.group(1).replace(",", "")
                    
                    extracted_data.append({
                        "Sort row": str(sort_row),
                        "Source": source,
                        "Posted date": posted_date,
                        "Merch (Ref #)": ref_number,
                        "Transaction ID": transaction_id,
                        "Total per Campaign": campaign_total,
                        "Campaign name": campaign_name_line,
                        "Company": "",
                        "Branch": "",
                        "Account": account_id,
                        "Event Code": "",
                        "Client (if applicable)": "",
                        "Description": ""
                    })
                    
                    sort_row += 1

        return pd.DataFrame(extracted_data, columns=self.columns)

def main():
    root = tk.Tk()
    
    # Set window icon if available
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    app = PDFToCSVConverter(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 3) - (root.winfo_width() // 3)
    y = (root.winfo_screenheight() // 14) - (root.winfo_height() // 14)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()