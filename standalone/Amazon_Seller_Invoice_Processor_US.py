import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime
import argparse
import warnings




def extract_invoice_info(text):
    """
    从PDF文本中提取发票号和日期
    返回格式: (invoice_number, invoice_date)
    """
    # 匹配发票号（英法双语）
    invoice_number = ""
    number_patterns = [
        r"Invoice\s*Number[:\/\s]*([A-Z0-9-]{7,})",          # 英文格式
        r"No\s*de\s*facture[:\/\s]*([A-Z00-9-]{7,})",       # 法文格式
        r"Facture\s*N°[:\/\s]*([A-Z0-9-]{7,})"              # 备用法文格式
    ]
    for pattern in number_patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            invoice_number = match.group(1).strip()
            break

    # 匹配发票日期（兼容多种格式）
    invoice_date = ""
    date_patterns = [
        r"Invoice\s*Date[:\/\s]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",    # 英文格式
        r"Date\s*de\s*facturation[:\/\s]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})",  # 法文格式
        r"Date\s*de\s*compte[:\/\s]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})"        # 备用法文格式
    ]
    for pattern in date_patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            raw_date = match.group(1)
            try:
                # 尝试解析日期格式
                for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y"):
                    try:
                        dt = datetime.strptime(raw_date, fmt)
                        invoice_date = dt.date().isoformat()
                        break
                    except:
                        continue
            except:
                invoice_date = raw_date  # 保留原始格式
            break

    return invoice_number, invoice_date

def extract_last_table_data(pdf_path):
    """
    增强版数据提取函数（支持发票信息）
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 第一步：提取整个PDF的发票信息
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            invoice_number, invoice_date = extract_invoice_info(full_text)

            # 第二步：处理表格数据
            last_valid_header = None
            collected_data = []
            
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    # 检测表头行
                    header_row = -1
                    for i, row in enumerate(table):
                        cell_check = [
                            any(keyword in str(cell).lower() 
                            for keyword in ["fees", "date", "location", "amount"])
                            for cell in row
                        ]
                        if any(cell_check):
                            header_row = i
                            break

                    # 更新表头信息
                    if header_row != -1:
                        header = [str(cell).strip().lower() for cell in table[header_row]]
                        last_valid_header = {
                            "header_row": header_row,
                            "col_indices": {
                                "location": next(
                                    (i for i, cell in enumerate(header) 
                                    if any(kw in cell for kw in ["province", "location", "region", ""])), -1),
                                "fees": next(
                                    (i for i, cell in enumerate(header) 
                                    if any(kw in cell for kw in ["fees", "amount"])), -1),
                                "gst/hst": next(
                                    (i for i, cell in enumerate(header) 
                                    if any(kw in cell for kw in ["gst/hst", "gst", "hst", "tax"])), -1),
                                "qst": next(
                                    (i for i, cell in enumerate(header) 
                                    if "qst" in cell), -1),
                                "pst": next(
                                    (i for i, cell in enumerate(header) 
                                    if any(kw in cell for kw in ["pst", "provincial sales tax"])), -1),
                                "total": next(
                                    (i for i, cell in enumerate(header) 
                                    if any(kw in cell for kw in ["total", "amount due"])), -1)
                            }
                        }

                    # 处理数据行
                    if last_valid_header:
                        header_row_idx = last_valid_header["header_row"]
                        col_indices = last_valid_header["col_indices"]
                        
                        for row in table[header_row_idx+1:]:
                            # 跳过汇总行
                            if any("total" in str(cell).lower() for cell in row):
                                continue
                            
                            # 构建数据条目
                            entry = {
                                "Invoice #": invoice_number,
                                "Date": invoice_date,
                                "Description": "Seller Fees/Frais de Vendeur",
                                "Location": safe_get(row, col_indices["location"]),
                                "Fees": clean_number(safe_get(row, col_indices["fees"])),
                                "GST/HST": clean_number(safe_get(row, col_indices["gst/hst"])),
                                "QST": clean_number(safe_get(row, col_indices["qst"])) if col_indices["qst"] != -1 else 0.0,
                                "PST": safe_get(row, col_indices["pst"]),
                                "Total": clean_number(safe_get(row, col_indices["total"])),
                                "Supplier": "Amazon.com Services LLC",
                                "Credit note?": "",
                                "Currency": "USD"
                            }
                            collected_data.append(entry)
            
            return collected_data

    except Exception as e:
        print(f"处理失败 {os.path.basename(pdf_path)}: {str(e)}")
    return None

def clean_number(val):
    """增强版数值清洗转换（支持括号负数和多种格式）"""
    try:
        # 预处理：移除所有空格、$符号和逗号
        cleaned = re.sub(r'[$\s,]', '', str(val))
        
        # 处理括号负数（支持格式如 ($203.10) 或 (1,234.56)）
        if re.match(r'^\(.*\)$', cleaned):
            cleaned = '-' + cleaned[1:-1]
            
        # 转换为浮点数
        return float(cleaned)
    except:
        return 0.0

def safe_get(row, index):
    """安全获取单元格内容"""
    try:
        return row[index] if index != -1 and index < len(row) else ""
    except IndexError:
        return ""

def batch_process_pdfs(pdf_folder, output_excel):
    """批量处理PDF文件夹"""
    all_data = []
    processed = 0
    failed_files = []
    
    if not os.path.exists(pdf_folder):
        print(f"错误：文件夹不存在 {pdf_folder}")
        return
    
    print(f"\n开始处理文件夹: {pdf_folder}")
    print("=" * 50)
    
    for filename in os.listdir(pdf_folder):
        if not filename.lower().endswith(".pdf"):
            continue
            
        pdf_path = os.path.join(pdf_folder, filename)
        print(f"处理中: {filename[:35]}...", end=" ")
        
        data = extract_last_table_data(pdf_path)
        if data:
            # 添加来源文件名
            for entry in data:
                entry["Source File"] = filename
            all_data.extend(data)
            processed += 1
            print("✓")
        else:
            failed_files.append(filename)
            print("×")
    
    if all_data:
        df = pd.DataFrame(all_data)
        
        # ========== 数据过滤 ==========
        if not df.empty:
            df['Location'] = df['Location'].astype(str).str.strip()
            is_date = df['Location'].str.contains(
                r'^\d{1,2}[/-]\d{1,2}[/-]\d{4}$',
                na=False
            )
            has_non_alpha = df['Location'].str.contains(
                r'[^A-Za-z\s]',
                na=False
            )
            is_empty = df['Location'].eq('')
            df = df[~(is_date | has_non_alpha | is_empty)]

        # ========== 新增功能 ==========
        # 省份缩写转换
        province_mapping = {
            'ALBERTA': 'AB',
            'ONTARIO': 'ON',
            'BRITISH COLUMBIA': 'BC',
            'QUEBEC': 'QC',
            'MANITOBA': 'MB',
            'SASKATCHEWAN': 'SK'
        }
        df['Location'] = df['Location'].str.upper().replace(province_mapping, regex=False)

        # 添加月份列
        df['month'] = pd.to_datetime(df['Date'], errors='coerce').dt.month
        df['month'] = df['month'].astype('Int64')  # 处理空值

        # 调整列顺序
        columns_order = [
            "Source File", "Date", "month", "Description",
            "Location", "Fees", "GST/HST", "QST", "PST", "Total", "Invoice #", 
            "Supplier", "Credit note?", "Currency"
        ]
        df = df[columns_order]

        # ========== 数据清洗 ==========
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
        num_cols = ["Fees", "GST/HST", "QST", "Total"]
        df[num_cols] = df[num_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

        # 保存Excel
        with pd.ExcelWriter(output_excel) as writer:
            df.to_excel(writer, index=False)
            print(f"\n成功处理 {processed} 个文件，提取 {len(df)} 条记录")
            print(f"结果文件: {os.path.abspath(output_excel)}")
        
        if failed_files:
            print("\n失败文件列表:")
            print("\n".join(f" - {f}" for f in failed_files))
    else:
        print("\n警告：未提取到有效数据")


if __name__ == "__main__":
    # 配置路径
    pdf_folder = r"C:\\Users\\Gilgamel\\Desktop\\StoreV-2021\\Amazon.com"
    output_excel = "C:\\Users\\Gilgamel\\Desktop\\StoreV-2021\\Amazon.com\\Seller_Fees_US.xlsx"


    
    # 执行处理
    batch_process_pdfs(pdf_folder, output_excel)
    
    # 自动打开结果文件（仅Windows）
    try:
        os.startfile(output_excel)
    except:
        print(f"请手动打开结果文件: {output_excel}")