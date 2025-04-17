import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os

class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Amazon Invoice Processor")

        # 配置国家、类型与脚本的映射
        self.script_mapping = {
            ('USA', 'Advertising'): 'script_usa_advertising.py',
            ('USA', 'Seller Fees'): 'script_us_seller_fees.py',
            ('USA', 'FBA Fulfillment'): 'script_us_fba_fulfillment.py',
            ('Canada', 'Advertising'): 'script_ca_advertising.py',
            ('Canada', 'Seller Fees'): 'script_can_seller_fees.py',
            ('Canada', 'FBA Fulfillment'): 'script_ca_fba_fulfillment.py',
            ('Mexico', 'Seller Fees'): 'script_mx_seller_fees.py',
            # 添加其他映射
        }

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 输入文件夹
        ttk.Label(self.master, text="PDF Folder:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.input_dir = tk.StringVar()
        ttk.Entry(self.master, textvariable=self.input_dir, width=40).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="Browse", command=self.select_input_dir).grid(row=0, column=2, padx=5, pady=5)

        # 输出文件夹
        ttk.Label(self.master, text="Output Path:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.output_dir = tk.StringVar()
        ttk.Entry(self.master, textvariable=self.output_dir, width=40).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="Browse", command=self.select_output_dir).grid(row=1, column=2, padx=5, pady=5)

        # 国家选择
        ttk.Label(self.master, text="Country:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.country = ttk.Combobox(self.master, values=['USA', 'Canada', 'Mexico'], state='readonly')
        self.country.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        # 类型选择
        ttk.Label(self.master, text="Invoice Type:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.invoice_type = ttk.Combobox(self.master, values=['Advertising', 'Seller Fees', 'FBA Fulfillment', 'FBA Fulfillment Non-Amazon'], state='readonly')
        self.invoice_type.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        # 运行按钮
        ttk.Button(self.master, text="Run", command=self.run_script).grid(row=4, column=1, pady=10)

    def select_input_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.input_dir.set(dir_path)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir.set(dir_path)

    def validate_inputs(self):
        return all([
            self.input_dir.get(),
            self.output_dir.get(),
            self.country.get(),
            self.invoice_type.get()
        ])

    def run_script(self):
        if not self.validate_inputs():
            messagebox.showerror("错误", "请填写所有字段")
            return

        script_name = self.script_mapping.get((self.country.get(), self.invoice_type.get()))
        if not script_name or not os.path.exists(script_name):
            messagebox.showerror("错误", "未找到对应的处理脚本")
            return

        cmd = ['python', script_name, '--input', self.input_dir.get(), '--output', self.output_dir.get()]
        try:
            subprocess.run(cmd, check=True)
            messagebox.showinfo("成功", "处理完成")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("错误", f"处理失败: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()