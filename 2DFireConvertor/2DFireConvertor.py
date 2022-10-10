import string
import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog as fd

import sys
import os
import pyodbc 
import numpy as np
import polars as pl
import pandas as pd

#Global variable for File path
FilePath=''


#---Read From template

def excelProgress():
    print("Start convert to ZiiPOS")
    
    # Python not allow to use 2D as the variable 
    _2DFireExcel = pd.read_excel(FilePath, index_col=None,dtype = str)
    _2DFExcelData=_2DFireExcel.astype("string")
        
    print(_2DFExcelData) 

    ZiiPOSMenuDataTemplate =pd.ExcelFile('template.xlsx')
    menuItemData = pd.read_excel(ZiiPOSMenuDataTemplate, "MenuItem")
    categoryData = pd.read_excel(ZiiPOSMenuDataTemplate, "Category")
  
    print(menuItemData)
    print(categoryData)







#-------------------------------------------  Gui Setting -------------------------------------------------------------------------------



class App:
    def __init__(self, root):
        #setting title
        root.title("2DFire Convertor")
        #setting window size
        width=536
        height=403
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_115=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_115["font"] = ft
        GLabel_115["fg"] = "#333333"
        GLabel_115["justify"] = "center"
        GLabel_115["text"] = "Select your File"
        GLabel_115.place(x=50,y=50,width=91,height=42)

        GLine_FilePath=tk.Entry(root)
        GLine_FilePath["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        GLine_FilePath["font"] = ft
        GLine_FilePath["fg"] = "#333333"
        GLine_FilePath["justify"] = "left"
        GLine_FilePath["text"] = FilePath
        GLine_FilePath.place(x=50,y=120,width=313,height=30)

        GButton_FileSelector=tk.Button(root)
        GButton_FileSelector["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        GButton_FileSelector["font"] = ft
        GButton_FileSelector["fg"] = "#000000"
        GButton_FileSelector["justify"] = "center"
        GButton_FileSelector["text"] = "select file"
        GButton_FileSelector.place(x=390,y=120,width=78,height=30)
        GButton_FileSelector["command"] = self.GButton_FileSelector_command

        GButton_StartButton=tk.Button(root)
        GButton_StartButton["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        GButton_StartButton["font"] = ft
        GButton_StartButton["fg"] = "#000000"
        GButton_StartButton["justify"] = "center"
        GButton_StartButton["text"] = "start"
        GButton_StartButton.place(x=50,y=290,width=115,height=31)
        GButton_StartButton["command"] = self.GButton_StartButton_command

    def GButton_FileSelector_command(self):
        print("command")
        filename = fd.askopenfilename()
        global FilePath 
        FilePath = filename
        GLine_FilePath=tk.Entry(root)
        GLine_FilePath["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        GLine_FilePath["font"] = ft
        GLine_FilePath["fg"] = "#333333"
        GLine_FilePath["justify"] = "left"
        GLine_FilePath.insert(0,filename)
        GLine_FilePath.place(x=50,y=120,width=313,height=30)
        print(FilePath)



    def GButton_StartButton_command(self):
        print("Start Process")
        print("File Path: " + FilePath)
        if(FilePath==""):
            print("Please Selec")
        else:
            excelProgress()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
