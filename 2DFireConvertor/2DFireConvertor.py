import string
import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog as fd
from tkinter import messagebox

import sys
import os
from unicodedata import category
import pyodbc 
import numpy as np
import pandas as pd

#Global variable for File path
FilePath=""





 

#---process 2DF categories
def _2DFCategoriesProgress(_2DFCategoryList,templateCategory):
    categoryDATA = pd.DataFrame()
    indexNumber=0

    for indexNumber in range(len(_2DFCategoryList)):
        temp = templateCategory
        if indexNumber < 10:
            id = "0"+str(indexNumber)
            cateID = "C"+ id
        else:
            id = str(indexNumber)
            cateID = "C"+ id
       
        
        categoryName=_2DFCategoryList.iloc[indexNumber]["分类名称"]      
        temp["Category"] = categoryName
        temp["Category1"] = categoryName
        temp["CultureCategory"] = categoryName
        temp["Code"] = cateID
        temp["OrderIndex"] = '00'
              
        categoryDATA = pd.concat([categoryDATA, temp],ignore_index=True)
   


    return categoryDATA

                # Code                             
            # MenuGroupCode 
            # Category                
            # Category1                         
            # Category2  
            # Category3  
            # Enable          
            # Notes      
            # ShowOnMainMenu  
            # ShowOnPOSMenu      
            # ShowOnPhoneOrderMenu           
            # ShowOnSelfOrderMenu
            # OnlineOrderCategory
            # OrderIndex                        
            # CourseCode 
            # ShowOnDineInMenu
            # ShowOnTakeawayMenu 
            # ShowOnQuickSaleMenu
            # ShowOnDeliveryMenu 
            # ShowOnPickupMenu
            # MinimumChoiceQty
            # MaximumChoiceQty
            # BorderColor
            # OnlineDisplayName1
            # OnlineDisplayName2
            # CultureCategory         
            # MenuGroupList     
            # OnlineStatus   
            # QRCodeStatus  




#---process 2DF Menu Items
def _2DFMenuItemProgress(_2DFMenuItems,templateMenuItem):
    MenuItemDATA = pd.DataFrame()
    indexNumber=0

    for indexNumber in range(len(_2DFMenuItems)):
        temp = templateMenuItem
        if indexNumber < 10:
            id = "00"+str(indexNumber)
            itemID = "M"+ id
        elif indexNumber >= 10 and indexNumber < 100:
            id = "0"+str(indexNumber)
            itemID = "M"+ id

        else:
            id = str(indexNumber)
            itemID = "M"+ id
       
        
        menuItemName=_2DFMenuItems.iloc[indexNumber]["商品名称"]
        menuItemName2=_2DFMenuItems.iloc[indexNumber]["双语名称"]   
        categoryName=_2DFMenuItems.iloc[indexNumber]["分类名称"]
        menuitemPrince=_2DFMenuItems.iloc[indexNumber]["单价(元)"]  

        temp["ItemCode"] = itemID   
        temp["Category"] = categoryName
        temp["Description1"] = menuItemName
        temp["Description2"] = menuItemName2
        temp["Price"] = menuitemPrince  
        temp["CultureDescription"] = menuItemName
        temp["PrinterPort"] = '0'
        

        MenuItemDATA = pd.concat([MenuItemDATA, temp],ignore_index=True)

    return MenuItemDATA



#---Read From template

def excelProgress():
    print("Start convert to ZiiPOS")
    
    # Python not allow to use 2D as the variable 
    _2DFireExcel = pd.read_excel(FilePath, index_col=None, dtype = str)
    _2DFExcelData=_2DFireExcel.astype("string")
    _2DFCategoryList = _2DFExcelData.drop_duplicates(subset=["分类名称"])
        
    #print(_2DFCategoryList) 
    
    ZiiPOSMenuDataTemplate = pd.ExcelFile("template.xlsx")
    overviewData = pd.read_excel(ZiiPOSMenuDataTemplate, "Overview",index_col=None, dtype = str)
    menuGroupTableData = pd.read_excel(ZiiPOSMenuDataTemplate, "MenuGroupTable",index_col=None, dtype = str)
    itemGroupTableData = pd.read_excel(ZiiPOSMenuDataTemplate, "ItemGroupTable",index_col=None, dtype = str)
    courseData = pd.read_excel(ZiiPOSMenuDataTemplate, "Course",index_col=None, dtype = str)
    categoryData = pd.read_excel(ZiiPOSMenuDataTemplate, "Category",index_col=None, dtype = str)
    presetNoteGroupData = pd.read_excel(ZiiPOSMenuDataTemplate, "PresetNoteGroup",index_col=None, dtype = str)
    menuItemData = pd.read_excel(ZiiPOSMenuDataTemplate, "MenuItem",index_col=None, dtype = str)
    menuItemRelationData = pd.read_excel(ZiiPOSMenuDataTemplate, "MenuItemRelation",index_col=None, dtype = str)
    subMenuLinkHeadData = pd.read_excel(ZiiPOSMenuDataTemplate, "SubMenuLinkHead",index_col=None, dtype = str)
    subMenuLinkDetailData = pd.read_excel(ZiiPOSMenuDataTemplate, "SubMenuLinkDetail",index_col=None, dtype = str)
    subItemGroupData = pd.read_excel(ZiiPOSMenuDataTemplate, "SubItemGroup",index_col=None, dtype = str)
    instructionLinkGroupData = pd.read_excel(ZiiPOSMenuDataTemplate, "InstructionLinkGroup",index_col=None, dtype = str)
    instructionLinkData = pd.read_excel(ZiiPOSMenuDataTemplate, "InstructionLink",index_col=None, dtype = str)

  


    tempCategory = categoryData.iloc[0]
    tempCategory = tempCategory.to_frame()
    tempCategory = tempCategory.transpose()
    
    categoryData = _2DFCategoriesProgress(_2DFCategoryList,tempCategory)
   

    tempMenuItem = menuItemData.iloc[0]
    
    tempMenuItem = tempMenuItem.to_frame()
    tempMenuItem = tempMenuItem.transpose()
    #print(tempMenuItem)

    preProcessMenuItem = menuItemData
    menuItemData = _2DFMenuItemProgress(_2DFExcelData,tempMenuItem)

    try:
        with pd.ExcelWriter("output.xlsx") as writer:
            overviewData.to_excel(writer, sheet_name="Overview",index = False, header=True)
            menuGroupTableData.to_excel(writer, sheet_name="MenuGroupTable",index = False, header=True)
            itemGroupTableData.to_excel(writer, sheet_name="ItemGroupTable",index = False, header=True)
            courseData.to_excel(writer, sheet_name="Course",index = False, header=True)
            categoryData.to_excel(writer, sheet_name="Category",index = False, header=True)
            presetNoteGroupData.to_excel(writer, sheet_name="PresetNoteGroup",index = False, header=True)
            menuItemData.to_excel(writer, sheet_name="MenuItem",index = False, header=True)
            menuItemRelationData.to_excel(writer, sheet_name="MenuItemRelation",index = False, header=True)
            subMenuLinkHeadData.to_excel(writer, sheet_name="SubMenuLinkHead",index = False, header=True)
            subMenuLinkDetailData.to_excel(writer, sheet_name="SubMenuLinkDetail",index = False, header=True)
            subItemGroupData.to_excel(writer, sheet_name="SubItemGroup",index = False, header=True)
            instructionLinkGroupData.to_excel(writer, sheet_name="InstructionLinkGroup",index = False, header=True)
            instructionLinkData.to_excel(writer, sheet_name="InstructionLink",index = False, header=True)

          
        print("Done")
        messagebox.showinfo("showinfo", "2DF Excel conversion process was completed, please find the 'output.xlsx' excel file in the same folder")

    except ValueError:
         print("Oops!  Failed to export to Excel file")
         messagebox.showerror("Error", "Oops!  Failed to export to Excel file, please close your excel program, and try again")


    



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
        alignstr = "%dx%d+%d+%d" % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_115=tk.Label(root)
        ft = tkFont.Font(family="Times",size=10)
        GLabel_115["font"] = ft
        GLabel_115["fg"] = "#333333"
        GLabel_115["justify"] = "center"
        GLabel_115["text"] = "Select your File"
        GLabel_115.place(x=50,y=50,width=91,height=42)

        GLine_FilePath=tk.Entry(root)
        GLine_FilePath["borderwidth"] = "1px"
        ft = tkFont.Font(family="Times",size=10)
        GLine_FilePath["font"] = ft
        GLine_FilePath["fg"] = "#333333"
        GLine_FilePath["justify"] = "left"
        GLine_FilePath["text"] = FilePath
        GLine_FilePath.place(x=50,y=120,width=313,height=30)

        GButton_FileSelector=tk.Button(root)
        GButton_FileSelector["bg"] = "#f0f0f0"
        ft = tkFont.Font(family="Times",size=10)
        GButton_FileSelector["font"] = ft
        GButton_FileSelector["fg"] = "#000000"
        GButton_FileSelector["justify"] = "center"
        GButton_FileSelector["text"] = "select file"
        GButton_FileSelector.place(x=390,y=120,width=78,height=30)
        GButton_FileSelector["command"] = self.GButton_FileSelector_command

        GButton_StartButton=tk.Button(root)
        GButton_StartButton["bg"] = "#f0f0f0"
        ft = tkFont.Font(family="Times",size=10)
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
        ft = tkFont.Font(family="Times",size=10)
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
