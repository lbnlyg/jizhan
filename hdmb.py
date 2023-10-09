import openpyxl
import os
#基站模板管理
dirs = os.listdir(".")
filename = "未找到基站文件"

print("文件夹下所有文件列表:", dirs)
for i in dirs:
    #print("i:"+i+"   ++++"+i[-5:] , i[-4:] )
    if i[-5:] == ".xlsx" or i[-4:] == ".xls":
        filename = i
        break

if filename=="未找到基站文件":
    print("没找到基站文件") 
    exit()
print("基站文件找到了："+("./"+filename))

wb = openpyxl.load_workbook("./"+filename)
# 显示所有表名
print("工作表列表:", wb.sheetnames)

filejztxt = open("jztxt.txt", "w+")
filejztxt.write('cellID'+"\n")
# 遍历所有表
for sheet in wb:
    sheettitle = sheet.title
    print("正在处理：", "="*10, sheet.title, "="*10, "工作表最大列数:",
          wb[sheet.title].max_column, "="*10, "工作表最大行数:", wb[sheet.title].max_row, "="*10)
    #print(sheet.max_column,"="*10,sheet[1])
    cellIDcolumn_letter="-1"
    for cellx in sheet[1]:
        if cellx.value =="CellID":
            print("值",cellx.value,end = "   ")
            print("数字列标",cellx.column,end = "\n")
            cellIDcolumn_letter = cellx.column_letter
            break
    if cellIDcolumn_letter=="-1":
        print("文件：",filename,"不是基站文件")
        exit()
    print("值",cellIDcolumn_letter,end = " \n  ")
    for jizhanID in sheet[cellIDcolumn_letter]:
        if jizhanID.value =="CellID" :
            continue
        filejztxt.write(str(jizhanID.value)+"\n")
        #print("基站ID:",jizhanID.value)
filejztxt.close

        