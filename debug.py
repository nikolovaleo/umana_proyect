import pandas as pd


import os
#from os import system, name, getcwd
#import win32com.client as win32
from win32com.client.gencache import EnsureDispatch
from pathlib import Path



excel = EnsureDispatch('Excel.Application')
excel.Visible = True


f_path = (os.getcwd() + '\\data\\4-17HPS_LIT_030A.xlsx')








addin_path = r'C:\Data\Excel Addins\OvationUtils\Addin\OvationUtils-F07d.xla'
excel.AddIns.Add(addin_path).Installed = False
excel.Workbooks.Open(addin_path)
excel.AddIns.Add(addin_path).Installed = True

#excel.RegisterXLL('C:/Data/Excel Addins/OvationUtils/Addin/OvationUtils-F07d.xla')
#excel.COMAddIns("OvationUtils-F07d.ExcelAddIn").Connect = True
wb1_raw = excel.Workbooks.Open(f_path)


#wb1_raw.RefreshAll()
#wb1_raw.SaveAs()
#wb1_raw.Close(True)
