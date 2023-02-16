import openpyxl
import os

dirs=os.listdir(".")
filename=""

print(dirs)
for i in dirs:
    if i[:-5]==".xlsx" or i[:-4]==".xls":
        filename=i
        break

    filename=i
print("找到了："+("./"+filename) )   

wb=openpyxl.load_workbook("./"+filename) 
print(wb.sheetnames)
