import os
dirs = os.listdir(".")

#在本地目录下搜索所有xlsx,xls表， 找到的第一个文件
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
print('Filename:', filename) 
print("基站文件为：  "+(filename))