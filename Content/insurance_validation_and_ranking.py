import cv2
import pytesseract
import pandas as pd
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import numpy as np
import os
from datetime import datetime

# Set Tesseract path (update this for your system)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

class InsuranceClaimProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Insurance Claim Processing System")
        self.root.geometry("1400x900")
        self.root.configure(bg='#f0f8ff')
        
        # Initialize variables
        self.forms = []
        self.records = []
        self.claim_id_counter = 101
        self.current_image_index = -1
        self.current_image = None
        
        # Create GUI
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Insurance Claim Processing System", 
                               font=("Arial", 16, "bold"), foreground="#2c3e50")
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))
        
        # Upload section
        ttk.Label(main_frame, text="Upload Forms:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Button(main_frame, text="Select Form Images", command=self.upload_forms).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Process buttons
        ttk.Button(main_frame, text="Process All Forms", command=self.process_forms).grid(row=2, column=0, pady=10, sticky=tk.W)
        ttk.Button(main_frame, text="Export Results", command=self.export_results).grid(row=2, column=1, pady=10, sticky=tk.W)
        ttk.Button(main_frame, text="Clear All", command=self.clear_all).grid(row=2, column=2, pady=10, sticky=tk.W)
        
        # Image preview
        ttk.Label(main_frame, text="Form Preview:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky=tk.W, pady=(20, 5))
        self.image_label = ttk.Label(main_frame, text="No image selected", background="white", anchor="center")
        self.image_label.grid(row=4, column=0, rowspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Navigation buttons
        nav_frame = ttk.Frame(main_frame)
        nav_frame.grid(row=8, column=0, sticky=tk.W, pady=5)
        ttk.Button(nav_frame, text="Previous", command=self.previous_image).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(nav_frame, text="Next", command=self.next_image).grid(row=0, column=1)
        
        # Extracted text
        ttk.Label(main_frame, text="Extracted Text:", font=("Arial", 10, "bold")).grid(row=3, column=1, sticky=tk.W, pady=(20, 5))
        self.text_area = scrolledtext.ScrolledText(main_frame, width=40, height=15, wrap=tk.WORD)
        self.text_area.grid(row=4, column=1, rowspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Validation results
        ttk.Label(main_frame, text="Validation Results:", font=("Arial", 10, "bold")).grid(row=3, column=2, sticky=tk.W, pady=(20, 5))
        self.validation_tree = ttk.Treeview(main_frame, columns=("Field", "Status", "Value"), show="headings", height=8)
        self.validation_tree.heading("Field", text="Field")
        self.validation_tree.heading("Status", text="Status")
        self.validation_tree.heading("Value", text="Value")
        self.validation_tree.column("Field", width=120)
        self.validation_tree.column("Status", width=80)
        self.validation_tree.column("Value", width=120)
        self.validation_tree.grid(row=4, column=2, rowspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Results section
        ttk.Label(main_frame, text="Processing Results:", font=("Arial", 10, "bold")).grid(row=9, column=0, sticky=tk.W, pady=(20, 5))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 10))
        
        # Validation results frame (List View)
        validation_frame = ttk.Frame(self.notebook, padding="5")
        
        # Add filter controls
        filter_frame = ttk.Frame(validation_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(filter_frame, text="Filter by:").grid(row=0, column=0, padx=(0, 5))
        self.filter_var = tk.StringVar()
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, 
                                   values=["All", "Valid", "Invalid"], state="readonly", width=10)
        filter_combo.set("All")
        filter_combo.grid(row=0, column=1, padx=(0, 10))
        filter_combo.bind("<<ComboboxSelected>>", self.filter_records)
        
        ttk.Label(filter_frame, text="Search:").grid(row=0, column=2, padx=(10, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=15)
        search_entry.grid(row=0, column=3, padx=(0, 5))
        search_entry.bind("<KeyRelease>", self.search_records)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(validation_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.validation_results_tree = ttk.Treeview(tree_frame, 
                                                   columns=("ClaimID", "Name", "ClaimType", "ClaimAmount", "Status", "Reasons"), 
                                                   show="headings", height=10)
        
        # Define headings
        self.validation_results_tree.heading("ClaimID", text="Claim ID", anchor=tk.CENTER)
        self.validation_results_tree.heading("Name", text="Name", anchor=tk.CENTER)
        self.validation_results_tree.heading("ClaimType", text="Claim Type", anchor=tk.CENTER)
        self.validation_results_tree.heading("ClaimAmount", text="Amount", anchor=tk.CENTER)
        self.validation_results_tree.heading("Status", text="Status", anchor=tk.CENTER)
        self.validation_results_tree.heading("Reasons", text="Reasons", anchor=tk.CENTER)
        
        # Define columns
        self.validation_results_tree.column("ClaimID", width=80, anchor=tk.CENTER)
        self.validation_results_tree.column("Name", width=150, anchor=tk.W)
        self.validation_results_tree.column("ClaimType", width=100, anchor=tk.CENTER)
        self.validation_results_tree.column("ClaimAmount", width=100, anchor=tk.E)
        self.validation_results_tree.column("Status", width=80, anchor=tk.CENTER)
        self.validation_results_tree.column("Reasons", width=200, anchor=tk.W)
        
        # Add scrollbar
        validation_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.validation_results_tree.yview)
        self.validation_results_tree.configure(yscrollcommand=validation_scrollbar.set)
        
        self.validation_results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        validation_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add tag configurations for color coding
        self.validation_results_tree.tag_configure("valid", background="#d4edda")  # Light green
        self.validation_results_tree.tag_configure("invalid", background="#f8d7da")  # Light red
        
        self.notebook.add(validation_frame, text="Validation Results")
        
        # Priority ranking frame
        priority_frame = ttk.Frame(self.notebook, padding="5")
        
        # Create treeview with scrollbar
        priority_tree_frame = ttk.Frame(priority_frame)
        priority_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.priority_tree = ttk.Treeview(priority_tree_frame, 
                                         columns=("Priority", "ClaimID", "Name", "ClaimType", "ClaimAmount"), 
                                         show="headings", height=10)
        
        # Define headings
        self.priority_tree.heading("Priority", text="Priority", anchor=tk.CENTER)
        self.priority_tree.heading("ClaimID", text="Claim ID", anchor=tk.CENTER)
        self.priority_tree.heading("Name", text="Name", anchor=tk.CENTER)
        self.priority_tree.heading("ClaimType", text="Claim Type", anchor=tk.CENTER)
        self.priority_tree.heading("ClaimAmount", text="Amount", anchor=tk.CENTER)
        
        # Define columns
        self.priority_tree.column("Priority", width=80, anchor=tk.CENTER)
        self.priority_tree.column("ClaimID", width=80, anchor=tk.CENTER)
        self.priority_tree.column("Name", width=150, anchor=tk.W)
        self.priority_tree.column("ClaimType", width=100, anchor=tk.CENTER)
        self.priority_tree.column("ClaimAmount", width=100, anchor=tk.E)
        
        # Add scrollbar
        priority_scrollbar = ttk.Scrollbar(priority_tree_frame, orient="vertical", command=self.priority_tree.yview)
        self.priority_tree.configure(yscrollcommand=priority_scrollbar.set)
        
        self.priority_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        priority_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.notebook.add(priority_frame, text="Priority Ranking")
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to process forms")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Statistics frame
        stats_frame = ttk.Frame(main_frame)
        stats_frame.grid(row=11, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        self.stats_var = tk.StringVar()
        self.stats_var.set("Total: 0 forms | Valid: 0 | Invalid: 0")
        stats_label = ttk.Label(stats_frame, textvariable=self.stats_var, font=("Arial", 9))
        stats_label.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(stats_frame, mode='determinate')
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
    def upload_forms(self):
        filepaths = filedialog.askopenfilenames(
            title="Select Form Images",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.tiff;*.bmp"), ("All files", "*.*")]
        )
        
        if filepaths:
            self.forms = list(filepaths)
            self.status_var.set(f"Loaded {len(self.forms)} form(s)")
            self.current_image_index = 0
            self.show_image()
            
    def show_image(self):
        if not self.forms or self.current_image_index < 0:
            return
            
        image_path = self.forms[self.current_image_index]
        self.current_image = cv2.imread(image_path)
        
        # Resize image for display
        display_image = self.resize_image(self.current_image, 300, 400)
        
        # Convert to PIL format and then to ImageTk format
        display_image = cv2.cvtColor(display_image, cv2.COLOR_BGR2RGB)
        pil_image = Image.fromarray(display_image)
        tk_image = ImageTk.PhotoImage(pil_image)
        
        # Update image label
        self.image_label.configure(image=tk_image, text="")
        self.image_label.image = tk_image
        
        # Update status
        self.status_var.set(f"Image {self.current_image_index + 1} of {len(self.forms)}")
        
        # Extract and show text
        self.extract_text()
        
    def resize_image(self, image, max_width, max_height):
        height, width = image.shape[:2]
        
        if width > max_width or height > max_height:
            # Calculate scaling factor
            scale = min(max_width/width, max_height/height)
            new_width = int(width * scale)
            new_height = int(height * scale)
            
            # Resize the image
            return cv2.resize(image, (new_width, new_height))
        return image
        
    def extract_text(self):
        if self.current_image is None:
            return
            
        # Preprocess image for better OCR
        gray = cv2.cvtColor(self.current_image, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        
        # Extract text using Tesseract
        text = pytesseract.image_to_string(thresh)
        
        # Display extracted text
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, text)
        
        # Validate the form
        self.validate_form(text)
        
    def validate_form(self, text):
        # Clear previous validation results
        for item in self.validation_tree.get_children():
            self.validation_tree.delete(item)
            
        claim = {"ClaimID": self.claim_id_counter, "Name": None, "ClaimType": None,
                 "ClaimAmount": None, "validation_status": None, "validation_reasons": ""}
        
        # Extract name
        name_match = re.search(r"Name of Insured\s*[:]?\s*([A-Za-z ]+)", text, re.IGNORECASE)
        if name_match:
            claim["Name"] = name_match.group(1).strip()
            self.validation_tree.insert("", "end", values=("Name", "✓ Found", claim["Name"]))
        else:
            self.validation_tree.insert("", "end", values=("Name", "✗ Missing", "Not found"))
        
        # Extract claim type
        claim_type = None
        if re.search(r"Health", text, re.IGNORECASE):
            claim_type = "Health"
        elif re.search(r"Auto", text, re.IGNORECASE):
            claim_type = "Auto"
        elif re.search(r"Life", text, re.IGNORECASE):
            claim_type = "Life"
            
        if claim_type:
            claim["ClaimType"] = claim_type
            self.validation_tree.insert("", "end", values=("Claim Type", "✓ Found", claim_type))
        else:
            self.validation_tree.insert("", "end", values=("Claim Type", "✗ Missing", "Not found"))
        
        # Extract claim amount
        amount_match = re.search(r"Claim Amount\s*[:]?\s*[\$]?\s*(\d+[.,]?\d*)", text, re.IGNORECASE)
        if amount_match:
            amount_str = amount_match.group(1).replace(',', '').replace('.', '')
            claim["ClaimAmount"] = int(amount_str)
            self.validation_tree.insert("", "end", values=("Claim Amount", "✓ Found", f"${claim['ClaimAmount']}"))
        else:
            # Fallback: look for any dollar amount in the text
            fallback_match = re.search(r"\$?\s*(\d+[.,]?\d*)", text)
            if fallback_match:
                amount_str = fallback_match.group(1).replace(',', '').replace('.', '')
                claim["ClaimAmount"] = int(amount_str)
                self.validation_tree.insert("", "end", values=("Claim Amount", "✓ Found", f"${claim['ClaimAmount']}"))
            else:
                self.validation_tree.insert("", "end", values=("Claim Amount", "✗ Missing", "Not found"))
        
        # Validate the claim
        reasons = []
        if not claim["Name"]:
            reasons.append("Missing Name")
        if claim["ClaimAmount"] is None or claim["ClaimAmount"] <= 0:
            reasons.append("Invalid ClaimAmount")
        if not claim["ClaimType"]:
            reasons.append("Invalid ClaimType")

        if reasons:
            claim["validation_status"] = "Invalid"
            claim["validation_reasons"] = ", ".join(reasons)
            self.validation_tree.insert("", "end", values=("Overall", "✗ Invalid", ", ".join(reasons)))
        else:
            claim["validation_status"] = "Valid"
            self.validation_tree.insert("", "end", values=("Overall", "✓ Valid", "All fields OK"))
        
        # Store the claim for later processing
        self.records.append(claim)
        
        # Update the results display after each form is processed
        self.update_results_display()
        
    def process_forms(self):
        if not self.forms:
            messagebox.showwarning("Warning", "Please upload form images first.")
            return
            
        self.records = []
        self.claim_id_counter = 101
        
        # Configure progress bar
        self.progress['maximum'] = len(self.forms)
        self.progress['value'] = 0
        
        # Process all forms
        for i, form in enumerate(self.forms):
            self.status_var.set(f"Processing form {i+1} of {len(self.forms)}")
            self.progress['value'] = i + 1
            self.root.update_idletasks()
            
            img = cv2.imread(form)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
            
            text = pytesseract.image_to_string(thresh)
            
            claim = {"ClaimID": self.claim_id_counter, "Name": None, "ClaimType": None,
                     "ClaimAmount": None, "validation_status": None, "validation_reasons": ""}

            # Extract name
            match = re.search(r"Name of Insured\s*[:]?\s*([A-Za-z ]+)", text, re.IGNORECASE)
            if match:
                claim["Name"] = match.group(1).strip()

            # Extract claim type
            if re.search(r"Health", text, re.IGNORECASE):
                claim["ClaimType"] = "Health"
            elif re.search(r"Auto", text, re.IGNORECASE):
                claim["ClaimType"] = "Auto"
            elif re.search(r"Life", text, re.IGNORECASE):
                claim["ClaimType"] = "Life"

            # Extract claim amount
            amount_match = re.search(r"Claim Amount\s*[:]?\s*[\$]?\s*(\d+[.,]?\d*)", text, re.IGNORECASE)
            if amount_match:
                amount_str = amount_match.group(1).replace(',', '').replace('.', '')
                claim["ClaimAmount"] = int(amount_str)
            else:
                # Fallback: look for any dollar amount in the text
                fallback_match = re.search(r"\$?\s*(\d+[.,]?\d*)", text)
                if fallback_match:
                    amount_str = fallback_match.group(1).replace(',', '').replace('.', '')
                    claim["ClaimAmount"] = int(amount_str)

            # Validate the claim
            reasons = []
            if not claim["Name"]:
                reasons.append("Missing Name")
            if claim["ClaimAmount"] is None or claim["ClaimAmount"] <= 0:
                reasons.append("Invalid ClaimAmount")
            if not claim["ClaimType"]:
                reasons.append("Invalid ClaimType")

            if reasons:
                claim["validation_status"] = "Invalid"
                claim["validation_reasons"] = ", ".join(reasons)
            else:
                claim["validation_status"] = "Valid"

            self.records.append(claim)
            self.claim_id_counter += 1
            
            # Update the display after each form
            self.update_results_display()
        
        # Display results
        self.status_var.set(f"Processed {len(self.forms)} form(s)")
        
    def update_results_display(self):
        # Clear previous results
        for item in self.validation_results_tree.get_children():
            self.validation_results_tree.delete(item)
            
        for item in self.priority_tree.get_children():
            self.priority_tree.delete(item)
        
        # Display validation results
        for record in self.records:
            tag = "valid" if record["validation_status"] == "Valid" else "invalid"
            self.validation_results_tree.insert("", "end", values=(
                record["ClaimID"],
                record["Name"] or "N/A",
                record["ClaimType"] or "N/A",
                f"${record['ClaimAmount']}" if record["ClaimAmount"] else "N/A",
                record["validation_status"],
                record["validation_reasons"]
            ), tags=(tag,))
        
        # Display priority ranking
        valid_claims = [claim for claim in self.records if claim["validation_status"] == "Valid"]
        valid_claims.sort(key=lambda x: x["ClaimAmount"], reverse=True)
        
        for i, claim in enumerate(valid_claims):
            self.priority_tree.insert("", "end", values=(
                i+1,
                claim["ClaimID"],
                claim["Name"],
                claim["ClaimType"],
                f"${claim['ClaimAmount']}"
            ))
        
        # Update statistics
        total = len(self.records)
        valid = len([r for r in self.records if r["validation_status"] == "Valid"])
        invalid = total - valid
        self.stats_var.set(f"Total: {total} forms | Valid: {valid} | Invalid: {invalid}")
        
    def filter_records(self, event=None):
        # Clear previous results
        for item in self.validation_results_tree.get_children():
            self.validation_results_tree.delete(item)
            
        # Filter records based on selection
        filter_value = self.filter_var.get()
        
        if filter_value == "All":
            filtered_records = self.records
        else:
            filtered_records = [r for r in self.records if r["validation_status"] == filter_value.capitalize()]
        
        # Display filtered results
        for record in filtered_records:
            tag = "valid" if record["validation_status"] == "Valid" else "invalid"
            self.validation_results_tree.insert("", "end", values=(
                record["ClaimID"],
                record["Name"] or "N/A",
                record["ClaimType"] or "N/A",
                f"${record['ClaimAmount']}" if record["ClaimAmount"] else "N/A",
                record["validation_status"],
                record["validation_reasons"]
            ), tags=(tag,))
            
    def search_records(self, event=None):
        # Clear previous results
        for item in self.validation_results_tree.get_children():
            self.validation_results_tree.delete(item)
            
        # Search records based on text
        search_text = self.search_var.get().lower()
        
        if not search_text:
            # If search is empty, show all records
            filtered_records = self.records
        else:
            # Filter records that match search text
            filtered_records = []
            for record in self.records:
                if (search_text in str(record["ClaimID"]).lower() or
                    search_text in (record["Name"] or "").lower() or
                    search_text in (record["ClaimType"] or "").lower() or
                    search_text in (record["validation_status"] or "").lower() or
                    search_text in (record["validation_reasons"] or "").lower()):
                    filtered_records.append(record)
        
        # Apply additional filter if selected
        filter_value = self.filter_var.get()
        if filter_value != "All":
            filtered_records = [r for r in filtered_records if r["validation_status"] == filter_value.capitalize()]
        
        # Display filtered results
        for record in filtered_records:
            tag = "valid" if record["validation_status"] == "Valid" else "invalid"
            self.validation_results_tree.insert("", "end", values=(
                record["ClaimID"],
                record["Name"] or "N/A",
                record["ClaimType"] or "N/A",
                f"${record['ClaimAmount']}" if record["ClaimAmount"] else "N/A",
                record["validation_status"],
                record["validation_reasons"]
            ), tags=(tag,))
        
    def export_results(self):
        if not self.records:
            messagebox.showwarning("Warning", "No data to export. Please process forms first.")
            return
            
        # Create DataFrames
        df = pd.DataFrame(self.records)
        valid_claims = df[df["validation_status"] == "Valid"].copy()
        valid_claims = valid_claims.sort_values(by="ClaimAmount", ascending=False)
        valid_claims["priority_score"] = range(1, len(valid_claims) + 1)
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save Results As"
        )
        
        if file_path:
            if file_path.endswith('.xlsx'):
                with pd.ExcelWriter(file_path) as writer:
                    df.to_excel(writer, sheet_name="Validation Results", index=False)
                    valid_claims.to_excel(writer, sheet_name="Priority Ranking", index=False)
            else:
                df.to_csv(file_path, index=False)
                
            messagebox.showinfo("Success", f"Results exported successfully to {file_path}")
        
    def clear_all(self):
        self.forms = []
        self.records = []
        self.claim_id_counter = 101
        self.current_image_index = -1
        self.current_image = None
        
        # Clear UI elements
        self.image_label.configure(image="", text="No image selected")
        self.text_area.delete(1.0, tk.END)
        self.search_var.set("")
        self.filter_var.set("All")
        self.stats_var.set("Total: 0 forms | Valid: 0 | Invalid: 0")
        self.progress['value'] = 0
        
        for item in self.validation_tree.get_children():
            self.validation_tree.delete(item)
            
        for item in self.validation_results_tree.get_children():
            self.validation_results_tree.delete(item)
            
        for item in self.priority_tree.get_children():
            self.priority_tree.delete(item)
            
        self.status_var.set("Ready to process forms")
        
    def next_image(self):
        if self.forms and self.current_image_index < len(self.forms) - 1:
            self.current_image_index += 1
            self.show_image()
            
    def previous_image(self):
        if self.forms and self.current_image_index > 0:
            self.current_image_index -= 1
            self.show_image()

if __name__ == "__main__":
    root = tk.Tk()
    app = InsuranceClaimProcessor(root)
    root.mainloop()