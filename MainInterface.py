# -*- coding: utf-8 -*-
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from ExcelDealFunction import exceldealfunc
import os
import pandas as pd
import numpy as np
import threading

class MyMainFace(object):
    def __init__(self):
        self.root = Tk()
        self.root.title('Excel合併工具')

        Label(self.root,text = "表1檔案路徑：").grid(row = 0, column = 0,sticky="w")
        self.path1 = StringVar()
        Entry(self.root, width=40,textvariable = self.path1,state='readonly').grid(row = 1, column = 0)
        Button(self.root, text = "選擇路徑", command = self.selectPath1).grid(row = 1, column = 1)

        self.num1 = IntVar()
        Label(self.root,text = "選擇比較列：").grid(row = 2, column = 0,sticky="w")

        self.lb1 = Listbox(self.root, width=50, height = 6)
        self.lb1.grid(row = 3, column = 0,sticky="w",columnspan = 2)
       # Button(self.root, text = "確認列(預設為0)", command = self.selectColumn1).grid(row = 1, column = 7)
        self.lb1.bind ( "<ButtonRelease>" , lambda e:self.selectColumn1())
        self.scrollbar1 = Scrollbar(self.root)
        self.scrollbar1.grid(column = 1,row=3 ,sticky='NSE')
        self.lb1.config(yscrollcommand = self.scrollbar1.set)
        self.scrollbar1.config(command=self.lb1.yview) 
    
        Label(self.root,text = "表2檔案路徑：").grid(row = 0, column = 2,sticky="w")
        self.path2 = StringVar()
        Entry(self.root,width=40,textvariable = self.path2,state='readonly').grid(row = 1, column = 2,)
        Button(self.root, text = "選擇路徑", command = self.selectPath2).grid(row = 1, column = 3)
        
        self.num2 = IntVar()
        Label(self.root,text = "選擇比較列：").grid(row = 2, column = 2,sticky="w")
        
        self.lb2 = Listbox(self.root, width=50, height = 6)
        self.lb2.grid(row = 3, column = 2,sticky="w",columnspan = 2)
       # Button(self.root, text = "確認列(預設為0)", command = self.selectColumn2).grid(row = 2, column = 7)
        self.lb2.bind ( "<ButtonRelease>" , lambda e:self.selectColumn2()
                        )
        self.scrollbar2 = Scrollbar(self.root)
        self.scrollbar2.grid(column = 3,row=3 ,sticky='NSE')
        self.lb2.config(yscrollcommand = self.scrollbar2.set)
        self.scrollbar2.config(command=self.lb2.yview)

        # 请选择生成表格路径

        Label(self.root,text = "表格存放路徑：").grid(row = 4, column = 0,sticky="w")
        self.path3 = StringVar()
        Entry(self.root,width=40,textvariable = self.path3,state='readonly').grid(row = 5, column = 0,columnspan=1)
        Button(self.root, text = "選擇路徑", command = self.selectPath3).grid(row = 5, column = 1)

        # 请输入生成表格名称

        Label(self.root,text = "生成表格的表名：").grid(row = 6, column = 0,sticky="w")
        self.name1 = StringVar()
        Entry(self.root,width=50, textvariable = self.name1).grid(row = 7, column = 0,sticky="w",columnspan=2)

        Label(self.root,text = "準確度(0~100)：").grid(row = 8, column = 0,sticky="w")
        self.num3 = StringVar()
        Entry(self.root,width=50, textvariable = self.num3).grid(row = 9, column = 0,sticky="w",columnspan=2)

        self.labeltxt=StringVar()
        self.labeltxt.set("aa")
        Label(self.root,textvariable = self.labeltxt,foreground="red").grid(row = 5, column = 2,columnspan=2,rowspan=4)

        self.lb3 = Listbox(self.root, width=50, height = 6)
        self.lb3.grid(row = 5, column = 2,columnspan = 2,rowspan=4)


        self.progress = Progressbar(self.root, orient = HORIZONTAL, length = 100, mode = 'determinate') 
        self.progress.grid(row = 9,column = 2,columnspan=2,ipadx=100)
        
        self.var = StringVar()
        self.var.set("開始")
        self.button =  Button(self.root,textvariable = self.var,command = lambda: threading.Thread(target=self.start).start(), width = 49)
        self.button.grid(row = 10,column = 0,pady=5,ipady=3,columnspan=2,)

        self.var2 = StringVar()
        self.var2.set("重置")
        self.button2 =  Button(self.root,textvariable = self.var2,command = self.reset, width = 49)
        self.button2.grid(row = 10,column = 2,pady=5,ipady=3,columnspan=2)
        
        self.root.mainloop()

    def start(self):
        """
            函数名：start(self)
            函数功能：开始按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        self.lb3.delete(0,'end')
        if self.path1.get():
            filename1=self.path1.get()
        else:
            self.lb3.insert(END,"請選擇好表1")
            return

        if self.path2.get():
            filename2=self.path2.get()
        else:
            self.lb3.insert(END,"請選擇好表2")
            return

        if self.path3.get():
            filename=self.path3.get()
        else:
            self.lb3.insert(END,"請選擇好路徑")
            return
        
        if self.name1.get():
            filename3=filename+"/"+self.name1.get()+".xlsx"
        else:
            self.lb3.insert(END,"請選擇好表名")
            return

        if self.num3.get():
            num3=int(self.num3.get())
        else:
            self.lb3.insert(END,"請選擇好準確度")
            return

        num1 = self.num1.get()
        num2 = self.num2.get()
        
        self.button.config(state="disable") # 关闭按钮1功能
       # self.root.withdraw()

        exceldealfunc(filename1,filename2,filename3,num1,num2,num3,self)
        #self.root.deiconify()

    def reset(self):
        """
            函数名：reset(self)
            函数功能：重置按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        self.progress['value'] = 0
        self.lb3.delete(0,'end')
        self.button.config(state="active") # 激活按钮1

        #fill_line = self.canvas.create_rectangle(2,2,0,27,width = 0,fill = "white") 
        self.var.set("开始")
        self.labeltxt.set(" ")
       # self.canvas.coords(fill_line, (0, 0, 181, 30))
        self.root.update()

    def selectPath1(self):
        """
            函数名：selectPath1(self)
            函数功能：选择路径1按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        path_ = askopenfilename(filetypes = [('Excel', '*.xls*')])
        self.path1.set(path_)
        self.num1.set(0)
        sheet = pd.read_excel(path_)
        self.lb1.delete(0,'end')
        for item in sheet.columns:
            self.lb1.insert(END,item)
        self.lb1.itemconfig(self.num1.get(), {'fg': 'red'})
            
    def selectPath2(self):
        """
            函数名：selectPath2(self)
            函数功能：选择路径2按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        path_ = askopenfilename(filetypes = [('Excel', '*.xls*')])
        self.path2.set(path_)
        self.num2.set(0)
        sheet = pd.read_excel(path_)
        self.lb2.delete(0,'end')
        for item in sheet.columns:
            self.lb2.insert(END,item)
        self.lb2.itemconfig(self.num2.get(), {'fg': 'red'})

    def selectPath3(self):
        """
            函数名：selectPath3(self)
            函数功能：选择路径3按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        path_ = askdirectory()
        self.path3.set(path_)

    def selectColumn1(self):
        self.lb1.itemconfig(self.num1.get(), {'fg': 'black'})
        value = self.lb1.get(self.lb1.curselection())
        sheet = pd.read_excel(self.path1.get())
        self.num1.set(sheet.columns.get_loc(value))
        self.lb1.itemconfig(self.num1.get(), {'fg': 'red'})
        
    def selectColumn2(self):
        self.lb2.itemconfig(self.num2.get(), {'fg': 'black'})
        value = self.lb2.get(self.lb2.curselection())
        sheet = pd.read_excel(self.path2.get())
        self.num2.set(sheet.columns.get_loc(value))
        self.lb2.itemconfig(self.num2.get(), {'fg': 'red'})
##############################程序入口#########################################
if __name__=="__main__":
    MyMainFace()
