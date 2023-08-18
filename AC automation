import win32com.client
import time
import os
import tkinter as tk
from tkinter import filedialog

start_time = time.time()  # 记录程序开始时间
try:


    def unlock_sheets(wb, password):
        """解锁工作簿中的所有工作表"""
        for sheet in wb.Sheets:
            sheet.Unprotect(password)



    def process_sheet(source_sheet, target_sheet, sheet_name):
        try:
            def find_end_row(sheet, start_row, start_col):
                """查找具有数据的最后一行"""
                current_row = start_row
                while True:
                    if not any(sheet.Range(sheet.Cells(current_row, start_col), sheet.Cells(current_row, 17)).Value[0]):
                        return current_row - 1
                    current_row += 1
                return current_row

            if sheet_name == "AC08":
                source_ranges = [source_sheet.Range("B15:N28"), source_sheet.Range("B31:N40")]
            elif sheet_name == "AC10":
                source_ranges = [source_sheet.Range("D14:F27"), source_sheet.Range("I14:K27")]
            elif sheet_name == "AC40":
                account_items_cell = source_sheet.Range("A13:Z16").Find(What="=CF", LookIn=win32com.client.constants.xlFormulas)
                if not account_items_cell:
                    print(f"Unable to find formula '{formula_to_find}' in sheet {sheet_name} of source file.")
                    return
                start_row = account_items_cell.Row
                start_col = account_items_cell.Column
                end_col = 17  # Corresponds to column
                end_row = find_end_row(source_sheet, start_row, start_col)
                source_ranges = [source_sheet.Range(source_sheet.Cells(start_row, start_col), source_sheet.Cells(end_row, end_col))]

            else:
                if sheet_name in ["ST01", "ST02", "ST03"]:
                    formular_to_find = "=itemname"
                else:
                    formular_to_find = "=accItem"
                account_items_cell = source_sheet.Range("A13:Z16").Find(What=formular_to_find, LookIn=win32com.client.constants.xlFormulas)

                if not account_items_cell:
                    print(f"Unable to find 'Account items' in sheet {sheet_name} of source file.")
                    return
                start_row = account_items_cell.Row
                start_col = account_items_cell.Column
                end_col = 17  # Corresponds to column Q
                end_row = find_end_row(source_sheet, start_row, start_col)

                if end_row <= start_row:  # No data found
                    print(f"No data found in sheet {sheet_name}. Skipping {sheet_name}")
                    return

                if sheet_name in [ "AC07", "AC15", "AC22", "AC23", "AC26", "AC31", "AC36", "AC37", "AC38", "AC39", "AC41",
                                   "ST01", "ST02", "ST03", "ST05", "ST06"]:

                    # source_sheet.Columns("YY:BBB").Delete()


                    # Convert source sheet values in the range to values in FD column
                    source_sheet.Range(source_sheet.Cells(start_row, start_col), source_sheet.Cells(end_row, start_col)).Copy()
                    source_sheet.Range("ZZ" + str(start_row)).PasteSpecial(Paste=win32com.client.constants.xlPasteValues)

                    # source_range = source_sheet.Range(source_sheet.Cells(start_row, start_col),
                    #                                   source_sheet.Cells(end_row, start_col))
                    # print(f"Source Range for {sheet_name}: {source_range.Address}")

                    # Convert target sheet values in the range to values in FD column

                    if sheet_name in ["ST01", "ST02", "ST03"]:
                        formula_for_target = "=itemname"
                    else:
                        formula_for_target = "=accItem"

                    target_start_row = target_sheet.Range("A13:Z16").Find(What=formula_for_target,
                                                                            LookIn=win32com.client.constants.xlFormulas).Row
                    target_end_row = find_end_row(target_sheet, target_start_row, start_col)
                    target_sheet.Range(target_sheet.Cells(target_start_row, start_col),
                                       target_sheet.Cells(target_end_row, start_col)).Copy()
                    target_sheet.Range("ZZ" + str(target_start_row)).PasteSpecial(Paste=win32com.client.constants.xlPasteValues)

                    # Convert target sheet values in the range to values in FD column
                    target_start_row = target_sheet.Range("A13:Z16").Find(What=formula_for_target,
                                                                          LookIn=win32com.client.constants.xlFormulas).Row
                    target_end_row = find_end_row(target_sheet, target_start_row, start_col)
                    # target_range = target_sheet.Range(target_sheet.Cells(target_start_row, start_col),
                    #                                   target_sheet.Cells(target_end_row, start_col))
                    # print(f"Target Range for {sheet_name}: {target_range.Address}")

                    for row in range(start_row, end_row + 1):
                        value_to_search = source_sheet.Range("ZZ" + str(row)).Value
                        found_cell = target_sheet.Range("ZZ" + str(target_start_row), "ZZ" + str(target_end_row)).Find(
                            What=value_to_search)
                        if found_cell:
                            source_range = source_sheet.Range(source_sheet.Cells(row, start_col + 1),
                                                              source_sheet.Cells(row, end_col))
                            target_range = target_sheet.Range(target_sheet.Cells(found_cell.Row, start_col + 1),
                                                              target_sheet.Cells(found_cell.Row, end_col))
                            target_range.Value = source_range.Value
                        else:
                            print(f"{value_to_search}  not found in {sheet_name}")
                    return

                # if sheet_name in ["AC07", "AC15", "AC22", "AC23", "AC26", "AC31", "AC36", "AC37", "AC38", "AC39",
                #                   "AC41", "ST01", "ST02", "ST03", "ST05", "ST06"]:
                #     target_sheet.Columns("ZZ:ZZ").Delete()

                else:
                    #print(sheet_name, end_row)
                    source_ranges = [source_sheet.Range(source_sheet.Cells(start_row, start_col), source_sheet.Cells(end_row, end_col))]


                    # for src_range in source_ranges:
                    #     print(f"Source range for sheet {sheet_name}: {src_range.Address}")


            # Extract data and paste to target sheet for each range
            for source_range in source_ranges:
                data = source_range.Value
                target_range_start_cell = target_sheet.Cells(source_range.Row, source_range.Column)
                target_range = target_sheet.Range(target_range_start_cell, target_range_start_cell.Offset(source_range.Rows.Count , source_range.Columns.Count ))
                # print(f"Target range for sheet {sheet_name}: {target_range.Address}")
                target_range.Value = data
                print(f"{sheet_name} has been copied successfully")

                # # 使用pandas打印获取到的数据的预览
                # display_data_preview(data)
        except win32com.client.pywintypes.com_error as error:
            print(f"Error occurred in sheet {sheet_name}.")
            print(f"Caught exception: {error}")
            print(f"Description: {error.excepinfo[2]}")



    def main():
        # 初始化tkinter的主窗口但不显示
        root = tk.Tk()
        root.withdraw()

        # 弹出选择源文件的对话框
        source_file_path = filedialog.askopenfilename(title="Select the Source File")
        if not source_file_path:
            print("Source file not selected!")
            return

        # 弹出选择目标文件的对话框
        target_file_path = filedialog.askopenfilename(title="Select the Target File")
        if not target_file_path:
            print("Target file not selected!")
            return

        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = False  # Set to True for debugging purposes, so you can see the Excel instance

        # Open the source and target files
        source_wb = Excel.Workbooks.Open(source_file_path, UpdateLinks=False)
        unlock_sheets(source_wb, "abc@1")
        target_wb = Excel.Workbooks.Open(target_file_path, UpdateLinks=False)
        unlock_sheets(target_wb, "abc@1")


        # 复制源文件INDEX表单中B2单元格的值
        source_value = source_wb.Sheets("INDEX").Range("B2:B4").Value

        # 粘贴到目标文件的INDEX表单中B2单元格
        target_wb.Sheets("INDEX").Range("B2:B4").Value = source_value
        target_wb.Sheets("INDEX").Range("AA1").Value = 2

        # if sheet_name in ["AC07", "AC15", "AC22", "AC23", "AC26", "AC31", "AC36", "AC37", "AC38", "AC39",
        #                   "AC41", "ST01", "ST02", "ST03", "ST05", "ST06"]:
        #     target_sheet.Columns("ZZ:ZZ").Delete()
        #     return

        sheet_names = ["AC01", "AC02", "AC03","AC04", "AC05", "AC06", "AC07", "AC08", "AC09", "AC10", "AC12", "AC13", "AC14",
                       "AC15", "AC16","AC17", "AC18", "AC19","AC20", "AC21", "AC22", "AC23", "AC24","AC25","AC26", "AC27",
                       "AC28", "AC29", "AC30", "AC31", "AC32", "AC33", "AC35", "AC34", "AC36", "AC37", "AC38", "AC39",
                       "AC40", "AC41", "ST01", "ST02", "ST03", "ST05", "ST06"]
        for sheet_name in sheet_names:
            process_sheet(source_wb.Sheets(sheet_name), target_wb.Sheets(sheet_name), sheet_name)


        sheets_to_delete_from = ["AC07", "AC15", "AC22", "AC23", "AC26", "AC31", "AC36", "AC37", "AC38", "AC39",
                                "AC41", "ST01", "ST02", "ST03", "ST05", "ST06"]  # 添加您想要删除列的所有工作表的名称
        for sheet_name in sheets_to_delete_from:
            target_wb.Sheets(sheet_name).Columns("ZZ:ZZ").Delete()

        # 打印目标文件check工作表中E5单元格的内容
        check_content = target_wb.Sheets("check").Range("E5").Value
        print(f"check result is {check_content}")

        # 重命名目标文件
        index_sheet = target_wb.Sheets("Index")
        new_name = str(index_sheet.Range("B2").Value) + str(index_sheet.Range("B3").Value) + str(
            index_sheet.Range("B4").Value)
        new_file_path = os.path.join(os.path.dirname(target_file_path), new_name + ".xlsb")



        # Save and close the workbooks
        target_wb.Save()
        source_wb.Close(False)
        target_wb.Close()
        Excel.Quit()

        # 重命名文件
        os.rename(target_file_path, new_file_path)

    if __name__ == "__main__":
        main()



    end_time = time.time()  # 记录程序结束时间
    elapsed_time = end_time - start_time  # 计算并打印程序运行时间
    print(f"The program took {elapsed_time} seconds to run.")

except win32com.client.pywintypes.com_error as error:
    print(f"Caught exception: {error}")
    print(f"Description: {error.excepinfo[2]}")


