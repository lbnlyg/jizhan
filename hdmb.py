import openpyxl
import os

dirs=os.listdir(".")
filename=""

print("文件夹下所有文件列表:",dirs)
for i in dirs:
    if i[:-5]==".xlsx" or i[:-4]==".xls":
        filename=i
        break

    filename=i
print("基站文件找到了："+("./"+filename) )   

wb=openpyxl.load_workbook("./"+filename) 
# 显示所有表名
print("工作表列表:",wb.sheetnames)

filejztxt=open("jztxt","w+")
# 遍历所有表
for sheet in  wb:
    sheettitle=sheet.title
    print("正在处理：","="*10,sheet.title,"="*10,"工作表最大列数:",wb[sheet.title].max_column,"="*10,"工作表最大行数:",wb[sheet.title].max_row,"="*10)
    print(sheet.max_column)
    