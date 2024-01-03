import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox
import sys
import os
import pyodbc 
import pandas as pd
import numpy as np
import xlsxwriter
import xlrd
import wget
import ssl
from tqdm import tqdm


ssl._create_default_https_context = ssl._create_unverified_context


def closesystem():
    sys.exit()

def convertToDDAExcel(SourceData,DDATemplate):
    DDAExcelDataFrame = pd.DataFrame()

    TemplateData=DDATemplate.iloc[0]
    TemplateData["WholesalePrice1(Inc GST)"] =0
    TemplateData["WholesalePrice2(Inc GST)"] =0

    z=0
    for z in tqdm(range(len(SourceData))):
    #for z in range(len(SourceData)):
        
        #print(z ," / ",len(SourceData) ," (2)" )
        if pd.notna(SourceData.iloc[z]["StockID"]):
            TemplateData=DDATemplate.iloc[0]
            TemplateData["ProductCode(15)"] = SourceData.iloc[z]["StockID"]
            TemplateData["Description1(100)"] = SourceData.iloc[z]["Description1"]
            TemplateData["Description2(100)"] = SourceData.iloc[z]["Description2"]
            
            TemplateData["Category(25)"] = SourceData.iloc[z]["DepartmentName"]

            TemplateData["SalesPrice1(Inc GST)"] = SourceData.iloc[z]["Price"]
            #TemplateData["WholesalePrice1(Inc GST)"] = SourceData.iloc[z]["F_WPrice"]
            
            TemplateData["LastOrderPrice(Ex GST)"] = SourceData.iloc[z]["ItemCost"]

            # BC1 = str(SourceData.iloc[z]["barcode1"])
            # BC2 = str(SourceData.iloc[z]["barcode2"])
            # BC3 = str(SourceData.iloc[z]["barcode3"])
            # BC4 = str(SourceData.iloc[z]["barcode4"])
            # BC5 = str(SourceData.iloc[z]["barcode5"])
            # BC6 = str(SourceData.iloc[z]["barcode6"])
            
            TemplateData["Barcode1(30)"] = str(SourceData.iloc[z]["barcode1"]).replace(" ", "")
            TemplateData["Barcode2(30)"] = str(SourceData.iloc[z]["barcode2"]).replace(" ", "")
            TemplateData["Barcode3(30)"] = str(SourceData.iloc[z]["barcode3"]).replace(" ", "")
            TemplateData["Barcode4(30)"] = str(SourceData.iloc[z]["barcode4"]).replace(" ", "")
            TemplateData["Barcode5(30)"] = str(SourceData.iloc[z]["barcode5"]).replace(" ", "")
            TemplateData["Barcode6(30)"] = str(SourceData.iloc[z]["barcode6"]).replace(" ", "")
            

            TemplateData["GSTRate"] = SourceData.iloc[z]["GSTRate"]
            TemplateData["Measurement (Pack)"] = SourceData.iloc[z]["PackSize"]
            
        
            if SourceData.iloc[z]["Scale"] == "Y":
                TemplateData["Scaleable"] = 1
            else:
                TemplateData["Scaleable"] = 0    

        
      

            TemplateData = TemplateData.to_frame()
            TemplateData = TemplateData.transpose()
                

            DDAExcelDataFrame = pd.concat([DDAExcelDataFrame, TemplateData],ignore_index=True)

  
    
    return DDAExcelDataFrame




         


def processProductWithBarCode(connect_string):
    PassSQLServerConnection = pyodbc.connect(connect_string)

  

    
    
    productListQuery = "select ItemDetails.*,barcode1,barcode2,barcode3,barcode4,barcode5,barcode6,y.Description as DepartmentName, NormalPrice.Price, CostCompare.ItemCost from ItemDetails left join (Select StockID,  MAX(CASE when a.rowNum=1 THEN Barcode else '' end) barcode1,  MAX(CASE when a.rowNum=2 THEN Barcode else '' end) barcode2,  MAX(CASE when a.rowNum=3 THEN Barcode else '' end) barcode3,  MAX(CASE when a.rowNum=4 THEN Barcode else '' end) barcode4,  MAX(CASE when a.rowNum=5 THEN Barcode else '' end) barcode5,  MAX(CASE when a.rowNum=6 THEN Barcode else '' end) barcode6   from( select *,ROW_NUMBER( ) OVER ( PARTITION BY StockId ORDER BY Barcode  ) AS rowNum from Barcode )a  GROUP BY StockID) x on ItemDetails.StockID =x.StockID  left join (Select * from   Department) y on ItemDetails.Department =y.DepartmentID left join (select * from NormalPrice) NormalPrice on NormalPrice.StockID=ItemDetails.StockID left join (select * from CostCompare)CostCompare on CostCompare.StockID = ItemDetails.StockID"


   

    productList = pd.read_sql_query(productListQuery, PassSQLServerConnection)
  

    result=productList
    

    print("total rows: ")
    print(len(result))
    
    directory='C:\\Ziitech'
    if not os.path.exists(directory):
        os.makedirs(directory)
        
    Export_file="C:\\Ziitech\\export_data.xlsx"
    print("Export to new excel")
    result.to_excel(Export_file, index = True, header=True,engine='xlsxwriter')
    print("Stage 1 Process completed")

    print("Stage 2 Start, converting Data to DDA Formate.....")
    DDAExcel = pd.DataFrame()
    
    DDADownloadTemplate_file="C:\\Ziitech\\ItemImportFormat.xls"
    if not os.path.exists(DDADownloadTemplate_file):
           
        try:
            downloadURL="https://download.ziicloud.com/programs/ziiposclassic/ItemImportFormat.xls"
                
            wget.download(downloadURL, DDADownloadTemplate_file)
        except wget.Error as ex:
            print("Download Files error")
        

    DDADataTemplate = pd.read_excel(DDADownloadTemplate_file, index_col=None,dtype = str)
     #DDAExcel=DDADataTemplete.astype(str)
    DDAExcel =DDADataTemplate
    
    DDAExcelFinal = convertToDDAExcel(result,DDAExcel)
    
    FinalDDA_file="C:\\Ziitech\\OutPut.xls"
    DDAExcelFinal.to_excel(FinalDDA_file, index = False, header=True,engine='xlsxwriter')
    messagebox.showinfo(title="Process Completed",message="Data Process Completed, Please check C:\\Ziitech Folder")
    
   




    
def ConnectionTest(connect_string):
    connectionTestResult = 0
   
    PassSQLServerConnection = pyodbc.connect(connect_string)

    print(connect_string)
    try:
        PassSQLServerConnection = pyodbc.connect(connect_string)
        print("{c} is working".format(c=connect_string))
        PassSQLServerConnection.close()
        connectionTestResult = 1
    except pyodbc.Error as ex:
        #print("{c} is not working".format(c=connect_string))
        messagebox.showerror(title="Error", message="{c} is not working")

    return connectionTestResult


def inforProcess(DBSource,DBName):
    connectionTestResult=0
    #connect_string = 'DRIVER={SQL Server}; SERVER='+DBSource+'; DATABASE='+DBName+'; UID='+DBUsername+'; PWD='+ DBPassword
    connect_string = 'DRIVER={SQL Server}; SERVER='+DBSource+'; DATABASE='+DBName+'; Trusted_Connection=yes;'
    
    
    
    if DBName=="":
        messagebox.showerror(title="Error", message="DB Name Field is Empty!!")
        connectionTestResult = 0
    else:
        connectionTestResult=ConnectionTest(connect_string)

    if connectionTestResult==1:
        print("next")
        processProductWithBarCode(connect_string)

    else:
        print("error")







class App:
    def __init__(self, root):
        #setting title
        root.title("EGB Office Converter")
        #setting window size
        width=600
        height=500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_DB_Source=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_DB_Source["font"] = ft
        GLabel_DB_Source["fg"] = "#333333"
        GLabel_DB_Source["justify"] = "left"
        GLabel_DB_Source["text"] = "DB Connection"
        GLabel_DB_Source.place(x=50,y=90,width=90,height=30)

        DBSource_Box=tk.Entry(root)
        DBSource_Box["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        DBSource_Box["font"] = ft
        DBSource_Box["fg"] = "#333333"
        DBSource_Box["justify"] = "left"
        DBSource_Box.insert(0,'localhost\\SQLEGB')
        #DBSource_Box["text"] = "localhost\\SQLEGB"
        
        DBSource_Box.place(x=190,y=90,width=275,height=30)

        # DB_UserName_Label=tk.Label(root)
        # ft = tkFont.Font(family='Times',size=10)
        # DB_UserName_Label["font"] = ft
        # DB_UserName_Label["fg"] = "#333333"
        # DB_UserName_Label["justify"] = "left"
        # DB_UserName_Label["text"] = "User Name"
        # DB_UserName_Label.place(x=50,y=140,width=90,height=30)

        # DB_UserName_Box=tk.Entry(root)
        # DB_UserName_Box["borderwidth"] = "1px"
        # ft = tkFont.Font(family='Times',size=10)
        # DB_UserName_Box["font"] = ft
        # DB_UserName_Box["fg"] = "#333333"
        # DB_UserName_Box["justify"] = "left"
        # #DB_UserName_Box["text"] = "sa"
        # DB_UserName_Box.insert(0,'sa')
        # DB_UserName_Box.place(x=190,y=140,width=275,height=30)

        # DB_Password_Label=tk.Label(root)
        # ft = tkFont.Font(family='Times',size=10)
        # DB_Password_Label["font"] = ft
        # DB_Password_Label["fg"] = "#333333"
        # DB_Password_Label["justify"] = "left"
        # DB_Password_Label["text"] = "Password"
        # DB_Password_Label.place(x=50,y=200,width=90,height=25)

        # DB_Password_Box=tk.Entry(root)
        # DB_Password_Box["borderwidth"] = "1px"
        
        # ft = tkFont.Font(family='Times',size=10)
        # DB_Password_Box["font"] = ft
        # DB_Password_Box["fg"] = "#333333"
        # DB_Password_Box["justify"] = "left"
        # #DB_Password_Box["text"] = "0000"
        # DB_Password_Box.insert(0,'0000')
        # DB_Password_Box.place(x=190,y=200,width=275,height=30)
        # DB_Password_Box["show"] = "*"

        
        
        DB_Name_Label=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        DB_Name_Label["font"] = ft
        DB_Name_Label["fg"] = "#333333"
        DB_Name_Label["justify"] = "left"
        DB_Name_Label["text"] = "DB Name"
        DB_Name_Label.place(x=50,y=260,width=90,height=25)

        DB_Name_Box=tk.Entry(root)
        DB_Name_Box["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        DB_Name_Box["font"] = ft
        DB_Name_Box["fg"] = "#333333"
        DB_Name_Box["justify"] = "left"    
        DB_Name_Box.place(x=190,y=260,width=275,height=30)
        DB_Name_Box.insert(0,'EGBOFFICE')
 









         #-----------------Functions---------------------------------
        def getDBSource():
            result=DBSource_Box.get()
            return result
           
            
        def getDBUsername():
            result=DB_UserName_Box.get()
            return result
      
        def getDBPassword():
            result=DB_Password_Box.get()
            return result
        
        def getDBName():
            result=DB_Name_Box.get()
            return result
      
        
        def StartConversionProcess():
            DBSource=getDBSource()
            # username=getDBUsername()
            # password=getDBPassword()
            databaseName=getDBName()
            #inforProcess(DBSource,DBUsername,DBPassword,DBName):
            inforProcess(DBSource,databaseName)
            

        def testDBSource():
            DBSource=getDBSource()
            # username=getDBUsername()
            # password=getDBPassword()
            databaseName=getDBName()
            #connect_string = 'DRIVER={SQL Server}; SERVER='+DBSource+'; DATABASE='+databaseName+'; UID='+username+'; PWD='+ password
            connect_string = 'DRIVER={SQL Server}; SERVER='+DBSource+'; DATABASE='+databaseName+'; Trusted_Connection=yes;'
            PassSQLServerConnection = pyodbc.connect(connect_string)
            if databaseName=="":
                messagebox.showerror(title="Error", message="DB Name Field is Empty!!")
            else:
                try:
                    PassSQLServerConnection = pyodbc.connect(connect_string)
                    print("{c} is working".format(c=connect_string))
                    PassSQLServerConnection.close()
                except pyodbc.Error as ex:
               
                    print("{c} is not working".format(c=connect_string))
                    messagebox.showerror(title="Error", message="{c} is not working")
          
            




            
            


            
            
        
        
        

            















        



            
#--------------Button Actions-------------------------
        Star_Button=tk.Button(root)
        Star_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        Star_Button["font"] = ft
        Star_Button["fg"] = "#000000"
        Star_Button["justify"] = "center"
        Star_Button["text"] = "Start"
        Star_Button.place(x=70,y=390,width=90,height=45)
        Star_Button["command"] = StartConversionProcess

        TEST_Button=tk.Button(root)
        TEST_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        TEST_Button["font"] = ft
        TEST_Button["fg"] = "#000000"
        TEST_Button["justify"] = "center"
        TEST_Button["text"] = "Test DB"
        TEST_Button.place(x=250,y=390,width=90,height=45)
        TEST_Button["command"] = testDBSource

        Close_Button=tk.Button(root)
        Close_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        Close_Button["font"] = ft
        Close_Button["fg"] = "#000000"
        Close_Button["justify"] = "center"
        Close_Button["text"] = "Close"
        Close_Button.place(x=420,y=390,width=90,height=45)
        Close_Button["command"] = closesystem
       
        




#----------------Not in use--------------------------------
    def Star_Button_command(self):
        print("Star_Button_command")
    def TEST_Button_command(self):
        print("command")
    def Close_Button_command(self):
        print("Exit")
        exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
