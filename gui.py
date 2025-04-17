import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import re

class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice Processor v2.0")
        master.geometry("600x300")

        # Script configuration
        self.script_mapping = {
            ('USA', 'Type1'): 'script_usa_type1.py',
            ('Canada', 'Type2'): 'script_canada_type2.py',
            # Add other mappings...
        }

        # UI initialization
        self.create_widgets()

    def create_widgets(self):
        # PDF Folder
        ttk.Label(self.master, text="PDF Folder:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        self.input_dir = tk.StringVar()
        ttk.Entry(self.master, textvariable=self.input_dir, width=45).grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(self.master, text="Browse", command=self.select_input_dir).grid(row=0, column=2, padx=5, pady=5)

        # Output File
        ttk.Label(self.master, text="Output File:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
        self.output_file = tk.StringVar()
        ttk.Entry(self.master, textvariable=self.output_file, width=45).grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(
            self.master, 
            text="Browse", 
            command=self.select_output_file
        ).grid(row=1, column=2, padx=5, pady=5)

        # Country Selection
        ttk.Label(self.master, text="Country:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
        self.country = ttk.Combobox(
            self.master, 
            values=['USA', 'Canada', 'Mexico'], 
            state='readonly',
            width=28
        )
        self.country.grid(row=2, column=1, padx=5, pady=5, sticky='ew')

        # Invoice Type Selection
        ttk.Label(self.master, text="Type:").grid(row=3, column=0, padx=10, pady=10, sticky='e')
        self.invoice_type = ttk.Combobox(
            self.master, 
            values=['Advertising', 'Seller Fees', 'FBA Fulfillment', 'FBA Fulfillment Non-Amazon'], 
            state='readonly',
            width=28
        )
        self.invoice_type.grid(row=3, column=1, padx=5, pady=5, sticky='ew')

        # Process Button
        ttk.Button(self.master, text="Start Processing", command=self.run_script).grid(row=4, column=1, pady=20)

        # Status Label
        self.status_label = ttk.Label(self.master, text="Ready", foreground="gray")
        self.status_label.grid(row=5, column=1, sticky='w')

        # Layout configuration
        self.master.columnconfigure(1, weight=1)
        for i in [0, 1, 2]:
            self.master.rowconfigure(i, weight=1)

    def select_input_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.input_dir.set(dir_path)
            self.status_label.config(text="Input folder selected", foreground="green")

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_file.set(file_path)
            self.status_label.config(text="Output path set", foreground="green")

    def validate_inputs(self):
        required_fields = [
            (self.input_dir.get(), "Please select PDF folder"),
            (self.output_file.get(), "Please specify output file"),
            (self.country.get(), "Please select country"),
            (self.invoice_type.get(), "Please select invoice type")
        ]

        for value, error_msg in required_fields:
            if not value:
                messagebox.showerror("Error", error_msg)
                return False

        output_path = self.output_file.get()
        if not output_path.lower().endswith(".xlsx"):
            messagebox.showerror("Error", "Output file must be .xlsx format")
            return False

        if re.search(r'[<>:"|?*]', os.path.basename(output_path)):
            messagebox.showerror("Error", "Filename contains invalid characters (<>:\"|?*)")
            return False

        return True

    def run_script(self):
        if not self.validate_inputs():
            return

        script_name = self.script_mapping.get((self.country.get(), self.invoice_type.get()))
        if not script_name or not os.path.exists(script_name):
            messagebox.showerror("Error", f"Script not found: {script_name}")
            return

        output_path = self.output_file.get()
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            cmd = [
                'python',
                script_name,
                '--input', self.input_dir.get(),
                '--output', output_path
            ]
            
            self.status_label.config(text="Processing...", foreground="blue")
            self.master.update()

            result = subprocess.run(
                cmd,
                check=True,
                capture_output=True,
                text=True
            )

            if result.returncode == 0:
                messagebox.showinfo("Success", f"File generated:\n{output_path}")
                self.status_label.config(text="Completed", foreground="green")
            else:
                messagebox.showerror("Error", f"Script execution failed:\n{result.stderr}")
                self.status_label.config(text="Execution Error", foreground="red")

        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Processing failed: {e.stderr}")
            self.status_label.config(text="Process Error", foreground="red")
        except Exception as e:
            messagebox.showerror("Error", f"System error: {str(e)}")
            self.status_label.config(text="System Error", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()