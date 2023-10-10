from tkinter import *
import tkinter.filedialog  # 注意次数要将文件对话框导入
# 定义一个处理文件的相关函数
def askfile():
    # 从本地选择一个文件，并返回文件的目录
    filename = tkinter.filedialog.askopenfilename()
    if filename != '':
         lb.config(text= filename)
    else:
         lb.config(text='您没有选择任何文件')
root = Tk()
root.config(bg='#87CEEB')
root.title("C语言中文网")
root.geometry('400x200+300+300')

btn=Button(root,text='选择文件',relief=RAISED,command=askfile)
btn.grid(row=0,column=0)
lb = Label(root,text='',bg='#87CEEB')
lb.grid(row=0,column=1,padx=5)
# 显示窗口
root.mainloop()