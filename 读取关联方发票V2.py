import pdfplumber
import os
import pandas as pd
import re
import win32com.client
import time
from tkinter import filedialog, Tk
# from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment

start_time = time.time() ##记录开始时间

def convert_date_format(date_str, format_from, format_to='%Y-%m-%d'):
    """Convert a date string from one format to another format."""
    from datetime import datetime
    date_obj = datetime.strptime(date_str, format_from)
    return date_obj.strftime(format_to)

def select_folder():
    """Use tkinter to let the user select a folder and return the folder path."""
    root = Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory()
    root.destroy()  # Close the Tkinter instance
    return folder_path

input_folder = select_folder() # 输入文件夹的路径，其中包含 PDF 文件
output_folder = r"C:\Users\10042129\Desktop\Intercompany Test_2023\Intercompany Invoice_2023" # 输出文件夹的路径

# 确保输出文件夹存在
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取输入文件夹中的所有 PDF 文件，但跳过已读的文件
pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf') and "_已读" not in f]

# 对每个PDF文件提取文本和表格数据
for pdf_file in pdf_files:
    file_path = os.path.join(input_folder, pdf_file)

    # 打开 PDF 文件
    with pdfplumber.open(file_path) as pdf:
        invoice_no = ""
        csv_file_path = None

        # 查找 "packing list" 的页数
        packing_list_page_num = None
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if "PACKING LIST" in text:
                packing_list_page_num = page_num
                break

        # 如果找到 "packing list"，则提取之前的所有页面
        if packing_list_page_num is not None:
            for page_num in range(packing_list_page_num):
                page = pdf.pages[page_num]
                text = page.extract_text()

                # 在第一页找到 "INVOICE NO .:" 和"ISSUE DATE: "并提取其后的文本,如果发票以MM开头则日期格式为yyyy-mm-dd，保留
                # 反之为泰国发票，日期格式为dd/mm/yyyy,进一步处理

                if page_num == 0:
                    invoice_no_match = re.search(r"INVOICE NO .: (\S+)", text)
                    invoice_no = invoice_no_match.group(1) if invoice_no_match else "unknown"

                    issue_date_match = re.search(r"ISSUE DATE : (\S+)", text)
                    if issue_date_match:
                        extracted_date = issue_date_match.group(1)
                        if invoice_no.startswith('MM'):
                            issue_date = extracted_date  # Keep the original format
                        else:
                            # Convert from 'dd/mm/yyyy' to 'yyyy-mm-dd'
                            issue_date = convert_date_format(extracted_date, '%d/%m/%Y')
                    else:
                        issue_date = "unknown_date"

                    csv_file_name = f"{invoice_no}_{issue_date}.csv"
                    csv_file_path = os.path.join(output_folder, csv_file_name)

                    # 如果 CSV 文件已存在，删除它
                    # 避免重复追加
                    if os.path.exists(csv_file_path):
                        os.remove(csv_file_path)

                # 提取表格
                tables = page.extract_tables()
            # 检验是否存在表格
            # if tables:
                for table_data in tables:
                    # 将表格数据转换为 Pandas DataFrame
                    table_df = pd.DataFrame(table_data)

                    # 追加到 CSV 文件
                    mode = 'a' if os.path.exists(csv_file_path) else 'w'
                    table_df.to_csv(csv_file_path, index=False, mode=mode, header=not os.path.exists(csv_file_path))

            print(f"表格已保存到 {csv_file_path}")

    os.rename(file_path, os.path.join(input_folder, pdf_file.replace('.pdf', '_已读.pdf')))

            # else:
            #     print(f"No tables found on page {page_num} of {pdf_file}")



# 获取 Excel 应用对象
Excel = win32com.client.Dispatch('Excel.Application')

# 打开工作簿，并设置 UpdateLinks 参数为 3
# UpdateLinks 参数的值意思是：0 不更新任何引用，3 更新外部引用
wb = Excel.Workbooks.Open(r'C:\Users\10042129\Desktop\Intercompany Test_2023\Intercomany Invoice List.xlsx', UpdateLinks=3)

# 刷新所有数据查询
Excel.ActiveWorkbook.RefreshAll()

# 等待刷新完成
Excel.CalculateUntilAsyncQueriesDone()

# 保存并关闭工作簿
wb.Save()
wb.Close()

# 关闭 Excel 应用
Excel.Quit()


end_time = time.time()  # 记录程序结束时间

# 计算并打印程序运行时间
elapsed_time = end_time - start_time

print(f"The program took {elapsed_time} seconds to run.")