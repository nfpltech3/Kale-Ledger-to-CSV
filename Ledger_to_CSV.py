import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import os
import sys
from datetime import datetime
import logging
import re

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Setup logging to file
logging.basicConfig(
    filename='ledger_to_purchase.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()

# --- Color Palette ---
BG_COLOR = "#F4F6F8"  # Nagarkot Light Background
CARD_BG = "#FFFFFF"   # Panel White
ACCENT = "#1F3F6E"    # Nagarkot Primary Blue
ACCENT_HOVER = "#2A528F" # Hover Blue
ACCENT_LIGHT = "#E3F2FD" # Light Blue
TEXT_PRIMARY = "#1E1E1E" # Dark Text
TEXT_SECONDARY = "#6B7280" # Muted Gray
BORDER_COLOR = "#E5E7EB" # Border Gray
SUCCESS_GREEN = "#1F3F6E" # Blue for success (Brand Rule)
ERROR_RED = "#D8232A"     # Nagarkot Red for errors
LOG_BG = "#FAFBFC"
LOG_FG = "#1E1E1E"

# Custom handler to display logs in GUI
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        try:
            if self.text_widget.winfo_exists():
                msg = self.format(record)
                self.text_widget.config(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                # Auto-scroll, but respect tags if implemented later
                self.text_widget.config(state='disabled')
                self.text_widget.see(tk.END)
        except Exception:
            pass

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Global variable to store Job Register path
JOB_REGISTER_PATH = None

# Function to get Job Number from Job Register CSV or Excel
def get_job_number(boe_number, log_callback):
    global JOB_REGISTER_PATH
    if JOB_REGISTER_PATH is None:
        log_callback("Job Register file not set.")
        return "NA"
    try:
        # Read Job Register based on file extension
        if JOB_REGISTER_PATH.endswith('.csv'):
            df = pd.read_csv(JOB_REGISTER_PATH)
        elif JOB_REGISTER_PATH.endswith('.xlsx'):
            df = pd.read_excel(JOB_REGISTER_PATH, engine='openpyxl')
        else:
            log_callback(f"Unsupported Job Register file format: {JOB_REGISTER_PATH}")
            logger.error(f"Unsupported Job Register file format: {JOB_REGISTER_PATH}")
            return "NA"
        
        # Try multiple possible BOE column names
        boe_column = None
        possible_boe_columns = ["BOE No", "BE No.", "BE No", "BOE No.", "BOE Number", "Bill of Entry No"]
        for col in possible_boe_columns:
            if col in df.columns:
                boe_column = col
                break
        if boe_column is None:
            log_callback(f"BOE column not found in Job Register file. Available columns: {list(df.columns)}")
            logger.error(f"BOE column not found in Job Register file. Available columns: {list(df.columns)}")
            return "NA"
        
        # Try multiple possible Job No column names
        job_column = None
        possible_job_columns = ["Job No.", "Job No", "Job Number", "Ref No", "Reference No"]
        for col in possible_job_columns:
            if col in df.columns:
                job_column = col
                break
        if job_column is None:
            log_callback(f"Job No column not found in Job Register file. Available columns: {list(df.columns)}")
            logger.error(f"Job No column not found in Job Register file. Available columns: {list(df.columns)}")
            return "NA"
        
        # Clean BOE numbers for matching
        df[boe_column] = df[boe_column].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        boe_number_clean = str(boe_number).strip()
        match = df[df[boe_column].str.lower() == boe_number_clean.lower()][job_column]
        if not match.empty:
            job_no = match.iloc[0]
            log_callback(f"Found Job No: {job_no} for BOE No.: {boe_number}")
            return job_no
        else:
            log_callback(f"No Job No found for BOE No.: {boe_number}")
            return "NA"
    except Exception as e:
        log_callback(f"Error reading Job Register file: {str(e)}")
        logger.error(f"Error reading Job Register file: {e}")
        return "NA"

# Function to create CSV
def create_csv(ledger_data, output_path, log_callback):
    log_callback("Creating CSV file...")
    try:
        today = datetime.now().strftime("%d-%b-%Y")  # e.g., 14-Jun-2025
        data_list = []
        for idx, row in ledger_data.iterrows():
            # Skip rows with empty or missing Receipt No.
            receipt_no = row.get('Receipt No.')
            if pd.isna(receipt_no) or str(receipt_no).strip() == '':
                log_callback(f"Skipping row {idx} due to missing Receipt No.: {receipt_no}")
                logger.warning(f"Skipping row {idx} due to missing Receipt No.: {receipt_no}")
                continue
            
            # Skip rows with empty or missing BOE No.
            boe_no = row.get('BOE No.')
            if pd.isna(boe_no) or str(boe_no).strip() == '':
                log_callback(f"Skipping row {idx} with Receipt No.: {receipt_no} due to missing BOE No.: {boe_no}")
                logger.warning(f"Skipping row {idx} with Receipt No.: {receipt_no} due to missing BOE No.: {boe_no}")
                continue

            # Handle Txn Date
            try:
                txn_date = pd.to_datetime(row['Txn Date'])
                if pd.isna(txn_date):  # Check for NaT
                    log_callback(f"Skipping row {idx} with Receipt No.: {receipt_no} due to missing or invalid Txn Date: {row['Txn Date']}")
                    logger.warning(f"Skipping row {idx} with Receipt No.: {receipt_no} due to missing or invalid Txn Date: {row['Txn Date']}")
                    continue
                vendor_inv_date = txn_date.strftime("%d-%b-%Y")
            except Exception as e:
                log_callback(f"Skipping row {idx} with Receipt No.: {receipt_no} due to invalid Txn Date: {str(e)}")
                logger.warning(f"Skipping row {idx} with Receipt No.: {receipt_no} due to invalid Txn Date: {e}")
                continue

            # Custom logic for ABBOTT HEALTHCARE PRIVATE LIMITED
            consignee_name = row.get('Consignee Name', '').strip()
            # Match any Consignee Name that starts with 'ABBOTT HEALTHCARE' (case-insensitive)
            if consignee_name.upper().startswith("ABBOTT HEALTHCARE"):
                charge_or_gl_name = "GATE PASS CHARGES - REIM"
                charge_or_gl_amount = "336"
                taxcode1 = ""
                taxcode1_amt = ""
                taxcode2 = ""
                taxcode2_amt = ""
                amount = "336"
                avail_tax_credit = "No"
            else:
                charge_or_gl_name = "GATE PASS CHARGES CCL"
                charge_or_gl_amount = "285"
                taxcode1 = "Central GST"
                taxcode1_amt = "25.65"
                taxcode2 = "State GST"
                taxcode2_amt = "25.65"
                avail_tax_credit = "100"
                amount = "285"

            job_no = get_job_number(boe_no, log_callback)
            if job_no and job_no != "NA":
                narration = f"Being Entry posted for Gatepass / Kale Logistics / {job_no}"
            else:
                narration = "Being Entry posted for Gatepass / Kale Logistics"
            data = {
                "Entry Date": today,
                "Posting Date": today,
                "Organization": "KALE LOGISTICS SOLUTIONS PVT LTD",
                "Organization Branch": "THANE",
                "Vendor Inv No": receipt_no,
                "Vendor Inv Date": vendor_inv_date,
                "Currency": "INR",
                "ExchRate": "1",
                "Narration": narration,
                "Due Date": "",
                "Charge or GL": "Charge",
                "Charge or GL Name": charge_or_gl_name,
                "Charge or GL Amount": charge_or_gl_amount,
                "DR or CR": "Dr",
                "Cost Center": "",
                "Branch": "HO",
                " Charge Narration": "GATE PASS CHARGES",
                "TaxGroup": "GSTIN",
                "Tax Type": "Taxable",
                "SAC or HSN": "996712",
                "Taxcode1": taxcode1,
                "Taxcode1 Amt": taxcode1_amt,
                "Taxcode2": taxcode2,
                "Taxcode2 Amt": taxcode2_amt,
                "Taxcode3": "",
                "Taxcode3 Amt": "",
                "Taxcode4": "",
                "Taxcode4 Amt": "",
                "Avail Tax Credit": avail_tax_credit,
                "LOB": "CCL IMP",
                "Ref Type": "",
                "Ref No": job_no,
                "Amount": amount,
                "Start Date": "",
                "End Date": "",
                "WH Tax Code": "",
                "WH Tax Percentage": "",
                "WH Tax Taxable": "",
                "WH Tax Amount": "",
                "Round Off": "Yes",
                "CC Code": ""
            }
            data_list.append(data)
        if not data_list:
            log_callback("No valid rows to process for CSV creation.")
            logger.warning("No valid rows to process for CSV creation.")
            return False
        df = pd.DataFrame(data_list)
        df.to_csv(output_path, index=False)
        log_callback(f"CSV saved to {output_path} with {len(data_list)} records")
        return True
    except Exception as e:
        log_callback(f"Failed to create CSV: {str(e)}")
        logger.error(f"Failed to create CSV: {e}")
        return False

# Tkinter GUI
class LedgerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ledger to Purchase CSV Converter")
        
        # Configure Fullscreen or Zoomed state
        try:
            self.root.state("zoomed")
        except:
            self.root.attributes("-fullscreen", True)
            
        self.root.configure(bg=BG_COLOR)

        # Variables
        self.ledger_path = None
        self.job_register_path = None
        self._logo_image = None

        # Setup Styles
        self._setup_styles()

        # Build UI
        self._create_widgets()

        # Logging Setup
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass

        # Card Style
        style.configure(
            "Card.TLabelframe",
            background=CARD_BG,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=CARD_BG,
            foreground=TEXT_PRIMARY,
            font=("Segoe UI", 10, "bold"),
        )

        # Button Styles
        style.configure(
            "Modern.TButton",
            font=("Segoe UI", 9),
            padding=(14, 6),
            background=CARD_BG,
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Modern.TButton",
            background=[("active", "#F5F5F5"), ("pressed", "#EEEEEE")],
        )

        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(20, 8),
            foreground="#FFFFFF",
            background=ACCENT,
            borderwidth=0,
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_HOVER), ("pressed", ACCENT_HOVER), ("disabled", "#90CAF9")],
            foreground=[("disabled", "#FFFFFF")],
        )

    def _create_widgets(self):
        # MAIN CONTAINER
        main_frame = tk.Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # HEADER
        header_frame = tk.Frame(main_frame, bg=CARD_BG, pady=16, padx=24)
        header_frame.pack(fill=tk.X)
        tk.Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=tk.X)

        # Logo
        logo_path = resource_path("logo.png")
        if HAS_PIL and os.path.isfile(logo_path):
            try:
                img = Image.open(logo_path)
                h = 40
                w = int(img.width * h / img.height)
                img = img.resize((w, h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                tk.Label(header_frame, image=self._logo_image, bg=CARD_BG).pack(side=tk.LEFT)
            except Exception:
                tk.Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=CARD_BG).pack(side=tk.LEFT)
        else:
            tk.Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=CARD_BG).pack(side=tk.LEFT)

        # Centered Title
        title_label = tk.Label(
            header_frame,
            text="Ledger to Purchase CSV Converter",
            font=("Segoe UI", 16, "bold"),
            bg=CARD_BG,
            fg=TEXT_PRIMARY,
        )
        title_label.place(relx=0.5, rely=0.3, anchor="center")

        subtitle_label = tk.Label(
            header_frame,
            text="Merge Ledger Reports with Job Registers for Logisys Upload",
            font=("Segoe UI", 9),
            bg=CARD_BG,
            fg=TEXT_SECONDARY,
        )
        subtitle_label.place(relx=0.5, rely=0.75, anchor="center")

        # BODY
        body = tk.Frame(main_frame, bg=BG_COLOR, padx=40, pady=30)
        body.pack(fill=tk.BOTH, expand=True)

        # File Selection Card
        file_card = ttk.LabelFrame(body, text="  Input Files  ", style="Card.TLabelframe", padding=20)
        file_card.pack(fill=tk.X, pady=(0, 20))
        
        file_inner = tk.Frame(file_card, bg=CARD_BG)
        file_inner.pack(fill=tk.BOTH, expand=True)

        # File Status Labels
        self.job_status_label = tk.Label(file_inner, text="Job Register: Not Selected", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 9))
        self.job_status_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.ledger_status_label = tk.Label(file_inner, text="Ledger Report: Not Selected", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 9))
        self.ledger_status_label.pack(anchor=tk.W, pady=(0, 15))

        btn_frame = tk.Frame(file_inner, bg=CARD_BG)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Select Job Register", command=self.select_job_register, style="Modern.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Select Ledger Report", command=self.select_ledger, style="Modern.TButton").pack(side=tk.LEFT)

        # Action Area
        action_frame = tk.Frame(body, bg=BG_COLOR)
        action_frame.pack(fill=tk.X, pady=(0, 20))

        self.process_button = ttk.Button(
            action_frame,
            text="\u25B6  Process & Generate CSV",
            command=self.process_files,
            style="Accent.TButton"
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 20))

        self.status_label_main = tk.Label(action_frame, text="Ready", fg=TEXT_SECONDARY, bg=BG_COLOR, font=("Segoe UI", 9))
        self.status_label_main.pack(side=tk.LEFT)

        # Log Card
        log_card = ttk.LabelFrame(body, text="  Processing Log  ", style="Card.TLabelframe", padding=15)
        log_card.pack(fill=tk.BOTH, expand=True)
        
        log_inner = tk.Frame(log_card, bg=CARD_BG)
        log_inner.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_inner, height=10, wrap=tk.WORD, state="disabled",
            bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9),
            relief="flat", padx=10, pady=10,
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # FOOTER
        footer_frame = tk.Frame(main_frame, bg=CARD_BG, padx=24, pady=10)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=tk.X, side=tk.BOTTOM)

        tk.Label(
            footer_frame,
            text="Nagarkot Forwarders Pvt. Ltd. \u00A9",
            fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 8),
        ).pack(side=tk.LEFT)

        ttk.Button(footer_frame, text="Exit", command=self.root.destroy, style="Modern.TButton").pack(side=tk.RIGHT)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')}: {message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.root.update()

    def select_job_register(self):
        global JOB_REGISTER_PATH
        csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if not csv_path:
            self.log("No Job Register file selected.")
            return
        JOB_REGISTER_PATH = csv_path
        self.job_register_path = csv_path
        self.job_status_label.config(text=f"Job Register: {os.path.basename(csv_path)}", fg=TEXT_PRIMARY)
        self.log(f"Selected Job Register: {os.path.basename(csv_path)}")
        logger.info(f"Job Register file selected: {csv_path}")

    def select_ledger(self):
        global JOB_REGISTER_PATH
        if JOB_REGISTER_PATH is None:
            self.status_label_main.config(text="Please select Job Register first.", fg=ERROR_RED)
            self.log("Job Register file not selected.")
            messagebox.showerror("Error", "Please select Job Register file before selecting Ledger Report.")
            return
        ledger_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not ledger_path:
            self.log("No Ledger Report selected.")
            return
        self.ledger_path = ledger_path
        self.ledger_status_label.config(text=f"Ledger Report: {os.path.basename(ledger_path)}", fg=TEXT_PRIMARY)
        self.log(f"Selected Ledger Report: {os.path.basename(ledger_path)}")
        logger.info(f"Ledger Report selected: {ledger_path}")

    def process_files(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        logger.info("Starting file processing")
        
        if not self.ledger_path:
            messagebox.showerror("Error", "Please select a Ledger Report")
            self.log("No Ledger Report selected.")
            return
        if not self.job_register_path:
            messagebox.showerror("Error", "Please select a Job Register file")
            self.log("No Job Register file selected.")
            return
            
        self.status_label_main.config(text="Processing...", fg=ACCENT)
        self.process_button.state(['disabled'])
        self.log("Starting processing...")
        self.root.update()

        # Create output directory
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, 'Kale Output')
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"Output directory: {output_dir}")

        # Generate output CSV path
        timestamp = datetime.now().strftime("%d-%m-%y %H-%M")
        output_csv = os.path.join(output_dir, f"purchase_{timestamp}.csv")
        logger.info(f"Output CSV: {output_csv}")

        # Check if CSV exists
        if os.path.exists(output_csv):
            response = messagebox.askyesno(
                "File Exists",
                f"CSV file {os.path.basename(output_csv)} already exists. Overwrite?",
                parent=self.root
            )
            if not response:
                self.log(f"Cancelled overwrite.")
                self.status_label_main.config(text="Cancelled", fg=TEXT_SECONDARY)
                self.process_button.state(['!disabled'])
                return

        # Read Ledger Report
        try:
            ledger_data = pd.read_excel(self.ledger_path, engine='openpyxl')
            self.log(f"Loaded Ledger Report: {len(ledger_data)} rows")
        except Exception as e:
            self.log(f"Failed to load Ledger Report: {str(e)}")
            self.status_label_main.config(text="Error loading file", fg=ERROR_RED)
            messagebox.showerror("Error", f"Failed to load Ledger Report: {str(e)}")
            self.process_button.state(['!disabled'])
            return

        # Create CSV
        if create_csv(ledger_data, output_csv, self.log):
            self.status_label_main.config(text="Completed Successfully", fg=SUCCESS_GREEN)
            self.log(f"CSV generated: {os.path.basename(output_csv)}")
            messagebox.showinfo("Success", f"CSV saved to {output_csv}")
        else:
            self.status_label_main.config(text="Failed", fg=ERROR_RED)
            self.log("Failed to generate CSV.")
            messagebox.showerror("Error", "Failed to generate CSV.")
            
        self.process_button.state(['!disabled'])

# Main
def main():
    try:
        logger.info("Starting Ledger to Purchase Converter")
        root = tk.Tk()
        app = LedgerApp(root)
        root.mainloop()
        logger.info("Application closed")
    except Exception as e:
        logger.error(f"Application error: {e}")
        messagebox.showerror("Error", f"Application error: {e}")

if __name__ == "__main__":
    main()
