import openpyxl
import os

dirs = os.listdir(".")
filename = ""

print("文件夹下所有文件列表:", dirs)
for i in dirs:
    if i[:-5] == ".xlsx" or i[:-4] == ".xls":
        filename = i
        break

    filename = i
print("基站文件找到了："+("./"+filename))

wb = openpyxl.load_workbook("./"+filename)
# 显示所有表名
print("工作表列表:", wb.sheetnames)

filejztxt = open("jztxt", "w+")
# 遍历所有表
for sheet in wb:
    sheettitle = sheet.title
    print("正在处理：", "="*10, sheet.title, "="*10, "工作表最大列数:",
          wb[sheet.title].max_column, "="*10, "工作表最大行数:", wb[sheet.title].max_row, "="*10)
    #print(sheet.max_column,"="*10,sheet[1])
    for cellx in sheet[1]:
        if cellx.value =="CellID":
            print("值",cellx.value,end = "   ")
            print("数字列标",cellx.column,end = "\n")
            cellIDcolumn = cellx.column
            print(cellx.column)
    '''       print("值",cellx.value,end = "   ")
        
        print("字母列标",cellx.column_letter,end = "   ")
        print("行号",cellx.row,end = "   ")
        print("坐标",cellx.coordinate,end = "\n")'''
