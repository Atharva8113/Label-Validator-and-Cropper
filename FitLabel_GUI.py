import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pdfplumber
from pypdf import PdfReader, PdfWriter
import os
import threading
import re

import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class PDFCropperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Label Validator & Cropper")
        self.root.geometry("800x650")
        self.root.configure(bg="#F5F5F5")
        self.root.minsize(600, 500)
        
        # Colors
        self.PRIMARY_BLUE = "#1E3A5F"
        self.ACCENT_BLUE = "#0078D4"
        self.SUCCESS_GREEN = "#107C10"
        self.ERROR_RED = "#C42B1C"
        self.WARNING_ORANGE = "#D83B01"
        self.BG_COLOR = "#F5F5F5"
        self.CARD_BG = "#FFFFFF"
        self.TEXT_PRIMARY = "#333333"
        self.TEXT_SECONDARY = "#666666"
        self.BORDER_COLOR = "#E0E0E0"
        
        # Data
        self.invoice_files: list[str] = []
        self.valid_pairs: list[tuple[str, str]] = []
        self.label_files: list[str] = []
        self.validated_labels: list[tuple] = []
        self.adc_file: str | None = None
        self.adc_data: dict[str, str] = {}
        
        self.setup_ui()
        
    def setup_ui(self):
        # === HEADER ===
        header_frame = tk.Frame(self.root, bg=self.CARD_BG, height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        header_inner = tk.Frame(header_frame, bg=self.CARD_BG)
        header_inner.pack(fill=tk.BOTH, expand=True, padx=20)
        
        # Logo (Left aligned)
        try:
            logo_path = resource_path("Nagarkot Logo.png")
            logo_img = Image.open(logo_path)
            # Calculate proportional resize for horizontal logo
            aspect_ratio = logo_img.width / logo_img.height
            new_height = 20
            new_width = int(new_height * aspect_ratio)
            logo_img = logo_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(header_inner, image=self.logo_photo, bg=self.CARD_BG)
            logo_label.pack(side=tk.LEFT, anchor=tk.W)
        except Exception as e:
            print(f"Logo error: {e}")
        
        # Title and subtitle (Center aligned)
        title_frame = tk.Frame(header_inner, bg=self.CARD_BG)
        title_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        title_label = tk.Label(title_frame, text="Label Cropper", 
                               font=('Segoe UI', 16, 'bold'), fg=self.PRIMARY_BLUE, bg=self.CARD_BG)
        title_label.pack(anchor=tk.CENTER)
        
        subtitle_label = tk.Label(title_frame, text="Validate invoice → Verify ADC → Crop & Rename labels", 
                                  font=('Segoe UI', 9), fg=self.TEXT_SECONDARY, bg=self.CARD_BG)
        subtitle_label.pack(anchor=tk.CENTER)
        
        # Header separator
        tk.Frame(self.root, height=1, bg=self.BORDER_COLOR).pack(fill=tk.X)
        
        # === MAIN CONTENT ===
        main_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # --- Card Row 1: Invoice and Labels ---
        cards_row1 = tk.Frame(main_frame, bg=self.BG_COLOR)
        cards_row1.pack(fill=tk.X, pady=(0, 10))
        
        # Invoice Card
        self.create_file_card(cards_row1, "📄 Invoice PDFs", "invoice", side=tk.LEFT)
        
        # Spacer
        tk.Frame(cards_row1, width=15, bg=self.BG_COLOR).pack(side=tk.LEFT)
        
        # Labels Card
        self.create_file_card(cards_row1, "🏷️ Label PDFs", "label", side=tk.LEFT)
        
        # --- Card Row 2: ADC ---
        cards_row2 = tk.Frame(main_frame, bg=self.BG_COLOR)
        cards_row2.pack(fill=tk.X, pady=(0, 10))
        
        # ADC Card (half width)
        self.create_file_card(cards_row2, "📋 ADC PDF", "adc", side=tk.LEFT, single=True)
        
        # --- Results Section ---
        results_frame = tk.Frame(main_frame, bg=self.CARD_BG, 
                                 highlightbackground=self.BORDER_COLOR, highlightthickness=1)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Results header
        results_header = tk.Frame(results_frame, bg=self.CARD_BG)
        results_header.pack(fill=tk.X, padx=15, pady=(12, 5))
        
        tk.Label(results_header, text="Preview", font=('Segoe UI', 10, 'bold'), 
                 fg=self.TEXT_PRIMARY, bg=self.CARD_BG).pack(side=tk.LEFT)
        
        self.lbl_preview_count = tk.Label(results_header, text="(0 items)", 
                                          font=('Segoe UI', 9), fg=self.TEXT_SECONDARY, bg=self.CARD_BG)
        self.lbl_preview_count.pack(side=tk.LEFT, padx=(5, 0))
        
        # Results text
        results_container = tk.Frame(results_frame, bg=self.CARD_BG)
        results_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 12))
        
        self.results_text = tk.Text(results_container, wrap=tk.WORD, bg="#FAFAFA", fg=self.TEXT_PRIMARY,
                                     font=('Consolas', 9), relief=tk.FLAT, padx=10, pady=10,
                                     state=tk.DISABLED, cursor="arrow", height=10,
                                     highlightbackground=self.BORDER_COLOR, highlightthickness=1)
        scrollbar = ttk.Scrollbar(results_container, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Text tags
        self.results_text.tag_configure("valid", foreground=self.SUCCESS_GREEN)
        self.results_text.tag_configure("invalid", foreground=self.ERROR_RED, font=('Consolas', 9, 'bold'))
        self.results_text.tag_configure("warning", foreground=self.WARNING_ORANGE)
        self.results_text.tag_configure("header", font=('Consolas', 9, 'bold'))
        self.results_text.tag_configure("normal", foreground=self.TEXT_PRIMARY)
        
        # === FOOTER ===
        footer_frame = tk.Frame(self.root, bg=self.BG_COLOR, height=60)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        footer_inner = tk.Frame(footer_frame, bg=self.BG_COLOR)
        footer_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Copyright
        copyright_label = tk.Label(footer_inner, text="© Nagarkot Forwarders Pvt Ltd", 
                                   font=('Segoe UI', 8), fg=self.TEXT_SECONDARY, bg=self.BG_COLOR)
        copyright_label.pack(side=tk.LEFT, anchor=tk.S)
        
        # Generate button
        self.btn_process = tk.Button(footer_inner, text="📥  Generate Cropped Labels", 
                                     command=self.start_processing, state=tk.DISABLED,
                                     bg=self.ACCENT_BLUE, fg="white", font=('Segoe UI', 10, 'bold'),
                                     relief=tk.FLAT, padx=25, pady=10, cursor="hand2",
                                     activebackground="#005A9E", activeforeground="white",
                                     disabledforeground="#FFFFFF")
        self.btn_process.pack(side=tk.RIGHT)
        
        # Status label
        self.status_var = tk.StringVar()
        self.lbl_status = tk.Label(footer_inner, textvariable=self.status_var, 
                                   font=('Segoe UI', 9), fg=self.ACCENT_BLUE, bg=self.BG_COLOR)
        self.lbl_status.pack(side=tk.RIGHT, padx=(0, 15))
        
    def create_file_card(self, parent, title, card_type, side=tk.LEFT, single=False):
        """Create a file selection card."""
        card = tk.Frame(parent, bg=self.CARD_BG, 
                        highlightbackground=self.BORDER_COLOR, highlightthickness=1)
        if single:
            card.pack(side=side, fill=tk.X, expand=False)
            card.configure(width=350)
        else:
            card.pack(side=side, fill=tk.BOTH, expand=True)
        
        card_inner = tk.Frame(card, bg=self.CARD_BG)
        card_inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=12)
        
        # Title row
        title_row = tk.Frame(card_inner, bg=self.CARD_BG)
        title_row.pack(fill=tk.X)
        
        tk.Label(title_row, text=title, font=('Segoe UI', 10, 'bold'), 
                 fg=self.TEXT_PRIMARY, bg=self.CARD_BG).pack(side=tk.LEFT)
        
        # Status label
        if card_type == "invoice":
            self.lbl_invoice_status = tk.Label(card_inner, text="No files selected", 
                                               font=('Segoe UI', 9), fg=self.TEXT_SECONDARY, bg=self.CARD_BG)
            self.lbl_invoice_status.pack(anchor=tk.W, pady=(3, 8))
        elif card_type == "label":
            self.lbl_labels_status = tk.Label(card_inner, text="No files selected", 
                                              font=('Segoe UI', 9), fg=self.TEXT_SECONDARY, bg=self.CARD_BG)
            self.lbl_labels_status.pack(anchor=tk.W, pady=(3, 8))
        else:
            self.lbl_adc_status = tk.Label(card_inner, text="Load ADC to validate Import License", 
                                           font=('Segoe UI', 9), fg=self.TEXT_SECONDARY, bg=self.CARD_BG)
            self.lbl_adc_status.pack(anchor=tk.W, pady=(3, 8))
        
        # Buttons row
        btn_row = tk.Frame(card_inner, bg=self.CARD_BG)
        btn_row.pack(fill=tk.X)
        
        if card_type == "invoice":
            btn_select = tk.Button(btn_row, text="Select", command=self.select_invoices,
                                   bg="#F0F0F0", fg=self.TEXT_PRIMARY, font=('Segoe UI', 9),
                                   relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                                   activebackground="#E0E0E0")
            btn_select.pack(side=tk.LEFT, padx=(0, 8))
            
            btn_clear = tk.Button(btn_row, text="Clear", command=self.clear_invoices,
                                  bg="#F0F0F0", fg=self.TEXT_PRIMARY, font=('Segoe UI', 9),
                                  relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                                  activebackground="#E0E0E0")
            btn_clear.pack(side=tk.LEFT)
            
        elif card_type == "label":
            btn_select = tk.Button(btn_row, text="Select", command=self.select_labels,
                                   bg="#F0F0F0", fg=self.TEXT_PRIMARY, font=('Segoe UI', 9),
                                   relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                                   activebackground="#E0E0E0")
            btn_select.pack(side=tk.LEFT, padx=(0, 8))
            
            btn_clear = tk.Button(btn_row, text="Clear", command=self.clear_labels,
                                  bg="#F0F0F0", fg=self.TEXT_PRIMARY, font=('Segoe UI', 9),
                                  relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                                  activebackground="#E0E0E0")
            btn_clear.pack(side=tk.LEFT)
            
        else:
            btn_select = tk.Button(btn_row, text="Select", command=self.select_adc,
                                   bg="#F0F0F0", fg=self.TEXT_PRIMARY, font=('Segoe UI', 9),
                                   relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                                   activebackground="#E0E0E0")
            btn_select.pack(side=tk.LEFT, padx=(0, 8))
            
            btn_clear = tk.Button(btn_row, text="Clear", command=self.clear_adc,
                                  bg="#F0F0F0", fg=self.TEXT_PRIMARY, font=('Segoe UI', 9),
                                  relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                                  activebackground="#E0E0E0")
            btn_clear.pack(side=tk.LEFT)
    
    def clear_invoices(self):
        self.invoice_files = []
        self.valid_pairs = []
        self.lbl_invoice_status.config(text="No files selected", fg=self.TEXT_SECONDARY)
        self.update_preview()
        
    def clear_labels(self):
        self.label_files = []
        self.validated_labels = []
        self.lbl_labels_status.config(text="No files selected", fg=self.TEXT_SECONDARY)
        self.btn_process.config(state=tk.DISABLED)
        self.update_preview()
        
    def clear_adc(self):
        self.adc_file = None
        self.adc_data = {}
        self.lbl_adc_status.config(text="Load ADC to validate Import License", fg=self.TEXT_SECONDARY)
        # Re-validate labels without ADC
        if self.validated_labels:
            self.validated_labels = [(p, l, lo, lic, iv, None) for p, l, lo, lic, iv, _ in self.validated_labels]
            self.display_validation_results()
    
    def update_results_text(self, content_list):
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        for item in content_list:
            text = item.get("text", "")
            tag = item.get("tag", "normal")
            self.results_text.insert(tk.END, text, tag)
        self.results_text.config(state=tk.DISABLED)
        
    def update_preview(self):
        """Update the preview section with current data."""
        content = []
        count = 0
        
        if self.valid_pairs:
            content.append({"text": "📄 Invoice Data:\n", "tag": "header"})
            # Group lots by item for cleaner display
            grouped_invoice: dict[str, list[str]] = {}
            for item, lot in self.valid_pairs:
                if item not in grouped_invoice:
                    grouped_invoice[item] = []
                if lot not in grouped_invoice[item]:
                    grouped_invoice[item].append(lot)
            
            for item, lots in grouped_invoice.items():
                lots_str = ", ".join(lots)
                content.append({"text": f"   Item: {item} → Lot: {lots_str}\n", "tag": "normal"})
                count += 1
            content.append({"text": "\n", "tag": "normal"})
        
        if self.validated_labels:
            content.append({"text": "🏷️ Label Validation:\n", "tag": "header"})
            for (path, list_no, lot_no, import_lic, invoice_valid, adc_valid) in self.validated_labels:
                filename = os.path.basename(path)
                if self.valid_pairs:
                    show_valid = invoice_valid
                else:
                    show_valid = list_no and lot_no
                
                if show_valid:
                    content.append({"text": f"   ✓ {filename} ({list_no}, {lot_no})", "tag": "valid"})
                    if adc_valid is True:
                        content.append({"text": " [License: ✓]\n", "tag": "valid"})
                    elif adc_valid is False:
                        content.append({"text": " [License: ✗]\n", "tag": "invalid"})
                    else:
                        content.append({"text": "\n", "tag": "normal"})
                else:
                    content.append({"text": f"   ✗ {filename}", "tag": "invalid"})
                    if list_no and lot_no:
                        content.append({"text": f" ({list_no}, {lot_no}) - NOT in invoice\n", "tag": "invalid"})
                    else:
                        content.append({"text": " - Could not extract data\n", "tag": "invalid"})
                count += 1
        
        self.lbl_preview_count.config(text=f"({count} items)")
        
        if content:
            self.update_results_text(content)
        else:
            self.update_results_text([{"text": "Select files to see preview...", "tag": "normal"}])
    
    def normalize_import_license(self, lic):
        """Normalize import license by removing slashes and leading zeros."""
        if not lic:
            return ""
        normalized = re.sub(r'[/\s]', '', lic.upper())
        match = re.match(r'([A-Z]+)(\d{4})(\d+)', normalized)
        if match:
            prefix = match.group(1)
            year = match.group(2)
            number = str(int(match.group(3)))
            return prefix + year + number
        return normalized
        
    # --- Invoice Handling ---
    def select_invoices(self):
        files = filedialog.askopenfilenames(title="Select Invoice PDF(s)", filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.invoice_files = list(files)
            self.status_var.set("Extracting invoice data...")
            self.root.update()
            self.extract_invoice_data()
            
    def extract_invoice_data(self) -> None:
        self.valid_pairs = []
        
        for inv_path in self.invoice_files:
            try:
                with pdfplumber.open(inv_path) as pdf:
                    current_item = None
                    for page in pdf.pages:
                        text = page.extract_text()
                        if not text:
                            continue
                        
                        lines = text.split('\n')
                        for line in lines:
                            # Reset item context if we reach summary/footer sections
                            if re.search(r'Subtotal|Total|Page\s+\d+|BIC-Code|IBAN', line, re.IGNORECASE):
                                current_item = None
                                
                            item_match = re.match(r'^([0-9][A-Z0-9]+)\s+.+\s+\d+\s+each', line, re.IGNORECASE)
                            if item_match:
                                current_item = item_match.group(1).upper()
                            
                            lot_match = re.search(r'Lot Number\s+(\S+)\s+Qty\.', line, re.IGNORECASE)
                            if lot_match and current_item:
                                lot_number = lot_match.group(1).upper()
                                # current_item is known to be str here due to truthy check
                                pair = (str(current_item), lot_number)
                                if pair not in self.valid_pairs:
                                    self.valid_pairs.append(pair)
                                # current_item is intentionally not reset here to handle multiple lots
                                
            except Exception as e:
                print(f"Error reading invoice {inv_path}: {e}")
        
        count = len(self.valid_pairs)
        if count > 0:
            self.lbl_invoice_status.config(text=f"✓ {len(self.invoice_files)} file(s), {count} items found", 
                                            fg=self.SUCCESS_GREEN)
        else:
            self.lbl_invoice_status.config(text="⚠ No item-lot pairs found", fg=self.WARNING_ORANGE)
            
        self.status_var.set("")
        
        # Re-validate labels if already loaded
        if self.label_files:
            self.validate_labels()
        else:
            self.update_preview()
        
    # --- Label Handling ---
    def select_labels(self):
        files = filedialog.askopenfilenames(title="Select Label PDF(s)", filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.label_files = list(files)
            self.validate_labels()
            
    def validate_labels(self):
        self.validated_labels = []
        
        for label_path in self.label_files:
            try:
                with pdfplumber.open(label_path) as pdf:
                    if not pdf.pages:
                        continue
                    text = pdf.pages[0].extract_text()
                    if not text:
                        continue
                    
                    match = re.search(r"List No\.:\s*([\w]+)\s+Lot No\.:\s*([\w]+)", text, re.IGNORECASE)
                    lic_match = re.search(r"Import License No\.:\s*(\S+)", text, re.IGNORECASE)
                    
                    if match:
                        list_no = match.group(1).upper()
                        lot_no = match.group(2).upper()
                        import_lic = lic_match.group(1) if lic_match else None
                        invoice_valid = (list_no, lot_no) in self.valid_pairs
                        self.validated_labels.append((label_path, list_no, lot_no, import_lic, invoice_valid, None))
                    else:
                        self.validated_labels.append((label_path, None, None, None, False, None))
                        
            except Exception as e:
                print(f"Error reading label {label_path}: {e}")
                self.validated_labels.append((label_path, None, None, None, False, None))
        
        if not self.valid_pairs:
            valid_count = sum(1 for v in self.validated_labels if v[1] and v[2])
            self.lbl_labels_status.config(
                text=f"{len(self.label_files)} file(s) selected (no invoice validation)",
                fg=self.TEXT_SECONDARY
            )
        else:
            valid_count = sum(1 for v in self.validated_labels if v[4])
            invalid_count = len(self.validated_labels) - valid_count
            self.lbl_labels_status.config(
                text=f"✓ {valid_count} valid, ✗ {invalid_count} invalid",
                fg=self.SUCCESS_GREEN if valid_count > 0 else self.ERROR_RED
            )
        
        self.display_validation_results()
        
        if len(self.validated_labels) > 0:
            self.btn_process.config(state=tk.NORMAL)
        else:
            self.btn_process.config(state=tk.DISABLED)
    
    def display_validation_results(self):
        self.update_preview()
        
        # Also run ADC validation if ADC is loaded
        if self.adc_data and self.validated_labels:
            self.validate_adc()
            
    # --- ADC Handling ---
    def select_adc(self):
        file = filedialog.askopenfilename(title="Select ADC PDF", filetypes=[("PDF Files", "*.pdf")])
        if file:
            self.adc_file = file
            self.status_var.set("Extracting ADC data...")
            self.root.update()
            self.extract_adc_data()
            
    def extract_adc_data(self):
        self.adc_data = {}
        product_pattern = r'(\d[A-Z]\d{4})'
        
        try:
            with pdfplumber.open(self.adc_file) as pdf:
                full_text = ''
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + ' '
                
                products = set(re.findall(product_pattern, full_text))
                
                for prod in products:
                    pattern = rf'{re.escape(prod)}.*?(IMP/[A-Z]+/\d+/\d+)'
                    match = re.search(pattern, full_text, re.IGNORECASE | re.DOTALL)
                    if match:
                        self.adc_data[prod.upper()] = match.group(1).upper()
                                
        except Exception as e:
            print(f"Error reading ADC: {e}")
        
        if self.adc_data:
            self.lbl_adc_status.config(text=f"✓ {len(self.adc_data)} products found", fg=self.SUCCESS_GREEN)
            self.validate_adc()
        else:
            self.lbl_adc_status.config(text="⚠ No product data found", fg=self.WARNING_ORANGE)
            
        self.status_var.set("")
        
    def validate_adc(self):
        updated_labels = []
        
        for (path, list_no, lot_no, import_lic, invoice_valid, _) in self.validated_labels:
            adc_valid = None
            
            list_no_upper = list_no.upper() if list_no else None
            if list_no_upper and import_lic and list_no_upper in self.adc_data:
                label_lic_norm = self.normalize_import_license(import_lic)
                adc_lic_norm = self.normalize_import_license(self.adc_data[list_no_upper])
                adc_valid = (label_lic_norm == adc_lic_norm)
            
            updated_labels.append((path, list_no, lot_no, import_lic, invoice_valid, adc_valid))
        
        self.validated_labels = updated_labels
        self.update_preview()
            
    # --- Processing ---
    def start_processing(self):
        output_dir = filedialog.askdirectory(title="Select Output Folder for Cropped Labels")
        if not output_dir:
            return
            
        self.btn_process.config(state=tk.DISABLED)
        self.status_var.set("Processing...")
        
        thread = threading.Thread(target=self.process_files, args=(output_dir,))
        thread.start()
        
    def get_pdf_info(self, pdf_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if not pdf.pages:
                    return None, None, None, None
                page = pdf.pages[0]
                text = page.extract_text()
                
                objects = []
                if hasattr(page, 'chars'): objects.extend(page.chars)
                if hasattr(page, 'lines'): objects.extend(page.lines)
                if hasattr(page, 'rects'): objects.extend(page.rects)
                if hasattr(page, 'images'): objects.extend(page.images)
                if hasattr(page, 'curves'): objects.extend(page.curves)
                
                if not objects:
                    return None, None, None, text
                    
                x0 = min(obj['x0'] for obj in objects)
                top = min(obj['top'] for obj in objects)
                x1 = max(obj['x1'] for obj in objects)
                bottom = max(obj['bottom'] for obj in objects)
                
                return (x0, top, x1, bottom), page.width, page.height, text
        except Exception as e:
            print(f"Error reading {pdf_path}: {e}")
            return None, None, None, None

    def process_files(self, output_dir):
        processed_count = 0
        skipped_count = 0
        errors = 0
        PADDING = 5
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        for (pdf_path, list_no, lot_no, import_lic, invoice_valid, adc_valid) in self.validated_labels:
            filename = os.path.basename(pdf_path)
            
            if self.valid_pairs:
                should_process = invoice_valid
            else:
                should_process = list_no and lot_no
            
            if not should_process:
                skipped_count += 1
                continue
                
            self.status_var.set(f"Cropping {filename}...")
            
            try:
                result, width, height, text = self.get_pdf_info(pdf_path)
                
                if not result:
                    errors += 1
                    continue
                    
                (p_x0, p_top, p_x1, p_bottom) = result
                
                new_filename = f"{list_no} {lot_no}.pdf"
                
                crop_x_min = max(0, p_x0 - PADDING)
                crop_x_max = min(width, p_x1 + PADDING)
                crop_y_max = min(height, (height - p_top) + PADDING)
                crop_y_min = max(0, (height - p_bottom) - PADDING)
                
                reader = PdfReader(pdf_path)
                writer = PdfWriter()
                
                for i in range(len(reader.pages)):
                    page = reader.pages[i]
                    page.cropbox.lower_left = (crop_x_min, crop_y_min)
                    page.cropbox.upper_right = (crop_x_max, crop_y_max)
                    page.mediabox.lower_left = (crop_x_min, crop_y_min)
                    page.mediabox.upper_right = (crop_x_max, crop_y_max)
                    writer.add_page(page)
                
                output_path = os.path.join(output_dir, new_filename)
                with open(output_path, "wb") as f:
                    writer.write(f)
                    
                processed_count += 1
                
            except Exception as e:
                print(f"Failed to process {filename}: {e}")
                errors += 1
        
        self.status_var.set(f"✓ {processed_count} saved, {skipped_count} skipped")
        self.root.after(0, lambda: messagebox.showinfo("Complete", 
            f"Processing Complete!\n\nCropped & Saved: {processed_count}\nSkipped: {skipped_count}\nErrors: {errors}"))
        self.root.after(0, lambda: self.btn_process.config(state=tk.NORMAL))

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCropperApp(root)
    root.mainloop()
