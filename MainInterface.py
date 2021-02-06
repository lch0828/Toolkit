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
        self.root.title('Excel Merging Tool')

        Label(self.root,text = "Excel file 1: ").grid(row = 0, column = 0,sticky="w")
        self.path1 = StringVar()
        Entry(self.root, width=40,textvariable = self.path1,state='readonly').grid(row = 1, column = 0)
        Button(self.root, text = "Choose file", command = self.selectPath1).grid(row = 1, column = 1)

        self.num1 = IntVar()
        Label(self.root,text = "Choose row: ").grid(row = 2, column = 0,sticky="w")

        self.lb1 = Listbox(self.root, width=50, height = 6)
        self.lb1.grid(row = 3, column = 0,sticky="w",columnspan = 2)
        self.lb1.bind ( "<ButtonRelease>" , lambda e:self.selectColumn1())
        self.scrollbar1 = Scrollbar(self.root)
        self.scrollbar1.grid(column = 1,row=3 ,sticky='NSE')
        self.lb1.config(yscrollcommand = self.scrollbar1.set)
        self.scrollbar1.config(command=self.lb1.yview) 
    
        Label(self.root,text = "Excel file 2: ").grid(row = 0, column = 2,sticky="w")
        self.path2 = StringVar()
        Entry(self.root,width=40,textvariable = self.path2,state='readonly').grid(row = 1, column = 2,)
        Button(self.root, text = "Choose file", command = self.selectPath2).grid(row = 1, column = 3)
        
        self.num2 = IntVar()
        Label(self.root,text = "Choose row:").grid(row = 2, column = 2,sticky="w")
        
        self.lb2 = Listbox(self.root, width=50, height = 6)
        self.lb2.grid(row = 3, column = 2,sticky="w",columnspan = 2)
        self.lb2.bind ( "<ButtonRelease>" , lambda e:self.selectColumn2()
                        )
        self.scrollbar2 = Scrollbar(self.root)
        self.scrollbar2.grid(column = 3,row=3 ,sticky='NSE')
        self.lb2.config(yscrollcommand = self.scrollbar2.set)
        self.scrollbar2.config(command=self.lb2.yview)

        # File path

        Label(self.root,text = "Choose save path").grid(row = 4, column = 0,sticky="w")
        self.path3 = StringVar()
        Entry(self.root,width=40,textvariable = self.path3,state='readonly').grid(row = 5, column = 0,columnspan=1)
        Button(self.root, text = "Choose path", command = self.selectPath3).grid(row = 5, column = 1)

        # File name

        Label(self.root,text = "Choose new file name").grid(row = 6, column = 0,sticky="w")
        self.name1 = StringVar()
        Entry(self.root,width=50, textvariable = self.name1).grid(row = 7, column = 0,sticky="w",columnspan=2)

        Label(self.root,text = "Choose accuracy (0~100): ").grid(row = 8, column = 0,sticky="w")
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
        self.var.set("Start")
        self.button =  Button(self.root,textvariable = self.var,command = lambda: threading.Thread(target=self.start).start(), width = 49)
        self.button.grid(row = 10,column = 0,pady=5,ipady=3,columnspan=2,)

        self.var2 = StringVar()
        self.var2.set("Reset")
        self.button2 =  Button(self.root,textvariable = self.var2,command = self.reset, width = 49)
        self.button2.grid(row = 10,column = 2,pady=5,ipady=3,columnspan=2)
        
        self.root.mainloop()

    def start(self):
        """
            Func: start button function
                input: MyMainFace type
                output: none
        """
        self.lb3.delete(0,'end')
        if self.path1.get():
            filename1=self.path1.get()
        else:
            self.lb3.insert(END,"Please choose file 1")
            return

        if self.path2.get():
            filename2=self.path2.get()
        else:
            self.lb3.insert(END,"Please chhose file 2")
            return

        if self.path3.get():
            filename=self.path3.get()
        else:
            self.lb3.insert(END,"Please choose save path")
            return
        
        if self.name1.get():
            filename3=filename+"/"+self.name1.get()+".xlsx"
        else:
            self.lb3.insert(END,"Please choose file name")
            return

        if self.num3.get():
            num3=int(self.num3.get())
        else:
            self.lb3.insert(END,"Please choose accuracy")
            return

        num1 = self.num1.get()
        num2 = self.num2.get()
        
        self.button.config(state="disable") # disable button 1
       # self.root.withdraw()

        exceldealfunc(filename1,filename2,filename3,num1,num2,num3,self)
        #self.root.deiconify()

    def reset(self):
        """
            Func: reset button
                input: MyMainFace type
                output: none
        """
        self.progress['value'] = 0
        self.lb3.delete(0,'end')
        self.button.config(state="active") # activate button1

        #fill_line = self.canvas.create_rectangle(2,2,0,27,width = 0,fill = "white") 
        self.var.set("Start")
        self.labeltxt.set(" ")
       # self.canvas.coords(fill_line, (0, 0, 181, 30))
        self.root.update()

    def selectPath1(self):
        """
            Func: SelectPath1 button
                input: MyMainFace type
                output: none
        """
        path_ = askopenfilename(filetypes = [('Excel', '.xlsx')])
        self.path1.set(path_)
        self.num1.set(0)
        sheet = pd.read_excel(path_)
        self.lb1.delete(0,'end')
        for item in sheet.columns:
            self.lb1.insert(END,item)
        self.lb1.itemconfig(self.num1.get(), {'fg': 'red'})
            
    def selectPath2(self):
        """
            Func: SelectPath2 button
                input: MyMainFace type
                output: none
        """
        path_ = askopenfilename(filetypes = [('Excel', '.xlsx')])
        self.path2.set(path_)
        self.num2.set(0)
        sheet = pd.read_excel(path_)
        self.lb2.delete(0,'end')
        for item in sheet.columns:
            self.lb2.insert(END,item)
        self.lb2.itemconfig(self.num2.get(), {'fg': 'red'})

    def selectPath3(self):
        """
            Func: SelectPath3 button
                input: MyMainFace type
                output: none
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
##############################Program entry#########################################
if __name__=="__main__":
    MyMainFace()
