import os
import re
import pdfplumber
import pandas as pd
import win32com.client
import shutil
from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment
import time

start_time = time.time()  # 记录程序开始时间

####################### Please make sure all packages have been installed "pdfplumber, pandas, xlwings, pywin32"
####################### 在运行程序前，请确认安装所有库 "pdfplumber, pandas, xlwings, pywin32"

# 创建凭证
credentials = Credentials('dongqian.han@minthgroup.com', 'Minth@456!')

# 创建配置
config = Configuration(server='mail.minthgroup.com', credentials=credentials)

# 创建账户
account = Account('dongqian.han@minthgroup.com', config=config, autodiscover=False, access_type=DELEGATE)

# 获取文件夹
folder = account.inbox / 'test'


# 定义保存附件的文件夹路径
folder_T0P_path = r"C:\Users\10042129\Desktop\T0P"
folder_T0Q_path = r"C:\Users\10042129\Desktop\T0Q"

# 确保文件夹存在
os.makedirs(folder_T0P_path, exist_ok=True)
os.makedirs(folder_T0Q_path, exist_ok=True)

# 初始化计数器和附件信息列表
new_email_count = 0
attachments_info = []

#在文件夹中搜索未读邮件
for item in folder.filter(is_read=False):
    print(f"Processing an unread email with subject: {item.subject or 'No Subject'}")  # 打印邮件主题，如果没有主题则打印 'No Subject'

    # 如果邮件有附件
    if item.attachments:
        for attachment in item.attachments:
            # 如果附件是文件
            if isinstance(attachment, FileAttachment):
                # 根据附件名称选择文件夹
                if 'T0P' in attachment.name:
                    folder_name = folder_T0P_path
                elif 'T0Q' in attachment.name:
                    folder_name = folder_T0Q_path
                else:
                    continue

                filepath = os.path.join(folder_name, attachment.name)
                print(f"Saving attachment to: {filepath}")  # 打印保存路径

                # 将附件保存到文件
                with open(filepath, 'wb') as f:
                    f.write(attachment.content)

                # 记录附件信息
                attachments_info.append(attachment.name)
                new_email_count += 1  # 增加未读邮件计数

    # 标记邮件为已读（可选）
    item.is_read = True
    item.save(update_fields=['is_read'])


# 获取所有的PDF文件
folder_path = r"C:\Users\10042129\Desktop\T0Q" # 根据需要修改文件夹路径，100为特定字符，在前面加r修正
pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]

# 定义处理函数
def process_case1(text_after_keyword):
    # 使用正则表达式查找最后的数字
    match = re.search(r'(\d[\d\s]*\d)$', text_after_keyword)
    if match:
        number = match.group(1)

        # 如果数字是两位数，将其转换为小数
        if number.isdigit() and len(number) == 2:
            number = "0." + number

        # 否则，将数字中的空格替换为"."
        else:
            number = number.replace(" ", ".")

        return number
    else:
        return text_after_keyword

def process_case2(text_after_keyword):

    # 查找XXXXXXXXX后面的所有文字
    text_after_XXXXXXXXX = text_after_keyword.split('XXXXXXXXX', 1)[-1]
    text_elements = text_after_XXXXXXXXX.split()

    if len(text_elements) == 3:
        # Assuming that the first element is the thousands, the second is the hundreds and the third is the decimal part
        number = f"{text_elements[0]},{text_elements[1]}.{text_elements[2]}"
        return number

    return text_after_XXXXXXXXX

def process_case3(text_after_keyword):
    # 使用正则表达式查找XXXXXXXXX和Total Liability中间的所有数字
    matches = re.findall(r'XXXXXXXXX.*?(\d[\d\s]*\d).*?Total Liability', text_after_keyword)
    if matches:
        numbers = []
        for match in matches:
            # 将数字中的空格替换为"."
            number = match.replace(" ", ".")
            numbers.append(number)

        # 如果列表不为空，取出第一个元素
        if numbers:
            return numbers[0]
    else:
        return text_after_keyword


def process_Case4(text_after_keyword):

        # 使用正则表达式查找 "T0Q" 或 "T0P" 或"AZD"或“AMX" 和 "PayDate: xx/xx/xxxx"
        match_T0Q_or_T0P = re.search(r'(T0Q|T0P|AZD|AMX)', text_after_keyword) # 根据需要修改comnay code，|在正则表达式中代表或
        match_PayDate = re.search(r'(PayDate: \d{2}/\d{2}/\d{4})', text_after_keyword)

        result = ""
        if match_T0Q_or_T0P:
            result += match_T0Q_or_T0P.group(1)
        if match_PayDate:
            result += " " + match_PayDate.group(1)

        return result if result else text_after_keyword


# 定义关键字字典，其中关键字是键，处理函数是值
keywords = {
    "Federal Unemployment Tax": process_case1,
    "State Unemployment Insurance - ER": process_case1,
    "Total Taxes Debited": process_case2,
    "ADP Direct Deposit": process_case2,
    "Wage Garnishments": process_case3,
    "Company Code:": process_Case4
}
all_results = [] # 用于保存所有PDF文件的结果

#last_text_after_keyword = None # 新定义的变量

# 对每个PDF文件提取第一页的文本数据
for pdf_file in pdf_files:
    with pdfplumber.open(os.path.join(folder_path, pdf_file)) as pdf:
        # 提取第一页的文本
        text = pdf.pages[0].extract_text()

        results = {}  # 用于保存这个PDF文件的结果

        # 遍历每个关键字
        for keyword, process_function in keywords.items():
            # 查找包含关键字的行
            for line in text.split('\n'):
                if keyword in line:
                    # 提取关键字后面的所有文字
                    text_after_keyword = line.split(keyword)[-1].strip()
                    #last_text_after_keyword = text_after_keyword  # 存储 text_after_keyword

                    # 使用正确的处理函数处理文本
                    text_after_keyword = process_function(text_after_keyword)

                    results[keyword] = text_after_keyword

                    print(f'{keyword} {text_after_keyword}')

        all_results.append(results)


# 创建一个 DataFrame
df = pd.DataFrame(all_results)

# 写入 Excel 文件
df.to_csv(rf"C:\Users\10042129\Desktop\PayrollData.csv", index=False) # 代表将处理结果保存在以下文件，根据需要修改文件路径及名称



# 获取 Excel 应用对象
Excel = win32com.client.Dispatch('Excel.Application')

# 打开工作簿，并设置 UpdateLinks 参数为 3
# UpdateLinks 参数的值意思是：0 不更新任何引用，3 更新外部引用
wb = Excel.Workbooks.Open(r'C:\Users\10042129\Desktop\T0Q\Salary Payroll_Test_V1.xlsm', UpdateLinks=3)

# 刷新所有数据查询
Excel.ActiveWorkbook.RefreshAll()

# 等待刷新完成
Excel.CalculateUntilAsyncQueriesDone()

# 获取名为 "index" 的工作表
sheet = wb.Sheets['index']

# 获取 I26 单元格的值
cell_value = sheet.Range('I26').Value

# 判断单元格的值是否大于1
if cell_value > 1:
    print(f"\033[1;31mOops, rounding number is {cell_value}, please double check \033[0m")
else:
    print(f"\033[1;31mGood Job, rounding number is {cell_value}, less than 1 \033[0m")


# 运行宏
Excel.Application.Run('ImportDataToSalarySheet')
Excel.Application.Run('ExportSpecificSheetsToCSV')


# 保存并关闭工作簿
wb.Save()
wb.Close()

# 关闭 Excel 应用
Excel.Quit()



# 动态获取company code和paydate信息
#company_code = last_text_after_keyword.replace('/','').replace(':','_')
company_code = text_after_keyword.replace('/','').replace(':','_')

# 定义源文件夹和目标文件夹
source_folder_A = r'C:\Users\10042129\Desktop\Template'
source_folder_B = r'C:\Users\10042129\Desktop\T0Q'
destination_folder = rf"C:\Users\10042129\Desktop/{company_code}"

# 创建目标文件夹
os.makedirs(destination_folder, exist_ok=True)

# 获取源文件夹A和B下的所有文件和文件夹
entries_A = os.listdir(source_folder_A)
entries_B = os.listdir(source_folder_B)

# 针对导出的csv文件出现空行导致上传失败的问题进行进一步处理
for entry in entries_A:
    if entry.endswith('.csv'):
        file_path = os.path.join(source_folder_A, entry)

        # 读取 CSV 文件
        df = pd.read_csv(file_path)

        # 删除空行
        df.dropna(how='all', inplace=True)

        # 在处理完文件后保存修改
        df.to_csv(file_path, index=False)

# 创建目标文件夹
os.makedirs(destination_folder, exist_ok=True)

# # 对每个文件和文件夹进行剪切操作
# for entry in entries_A:
#     source_path = os.path.join(source_folder_A, entry)
#     destination_path = os.path.join(destination_folder, entry)
#     shutil.move(source_path, destination_path)
#
# for entry in entries_B:
#     source_path = os.path.join(source_folder_B, entry)
#     destination_path = os.path.join(destination_folder, entry)
#     shutil.move(source_path, destination_path)
#
def move_files_except_type(source_folder, destination_folder, file_type_to_skip):
    entries = os.listdir(source_folder)
    for entry in entries:
        # 检查文件扩展名是否不是指定的文件类型
        if not entry.endswith(file_type_to_skip):
            source_path = os.path.join(source_folder, entry)
            destination_path = os.path.join(destination_folder, entry)
            shutil.move(source_path, destination_path)

# 保留 .xlsm 文件，剪切其他所有文件
file_type_to_keep = '.xlsm'

move_files_except_type(source_folder_A, destination_folder, file_type_to_keep)
move_files_except_type(source_folder_B, destination_folder, file_type_to_keep)

print('All Done!!')

end_time = time.time()  # 记录程序结束时间

# 计算并打印程序运行时间
elapsed_time = end_time - start_time

print(f"The program took {elapsed_time} seconds to run.")