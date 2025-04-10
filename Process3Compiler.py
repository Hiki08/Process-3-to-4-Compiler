# %%
from pathlib import Path
import shutil
import glob
import pandas as pd
import os
import numpy as np
import math
import openpyxl
import datetime
import time
from openpyxl.styles import Font
import pyttsx3
from python_calamine import CalamineWorkbook
import xlrd

# #GUI
# from tkinter import *
# import tkinter as tk
# from tkinter import ttk

# #Fixing Blur UI
# from ctypes import windll

# %%
dateToday = ""
currentDateToday = ""
process3Data = ""
process3Data2 = ""
mergedProcess3Data = ""
isFileExist = ""
readCount = 0

# %%
def GetDateToday():
    global dateToday

    dateToday = datetime.datetime.today()
    dateToday = dateToday.strftime('%y%m%d')

# %%
def ReadProcess3Csv():
    global process3Data
    global process3Data2

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    process3Directory = (r'\\192.168.2.10\csv\csv\VT3')
    os.chdir(process3Directory)
    
    process3Data = pd.read_csv(f'log000_3_1.csv', encoding='latin1', skiprows=0)
    # process3Data = pd.read_csv(f'log000_3_240814.csv', encoding='latin1', skiprows=0)

    process3Directory2 = (r'\\192.168.2.10\csv\csv\VT3')
    os.chdir(process3Directory2)
    
    process3Data2 = pd.read_csv(f'log000_3_2.csv', encoding='latin1', skiprows=0)
    # process3Data2 = pd.read_csv(f'log000_3_2_240814.csv', encoding='latin1', skiprows=0)

# %%
def MergedProcess3():
    global process3Data
    global process3Data2
    global mergedProcess3Data

    mergedProcess3Data = pd.concat([
        process3Data["DATA No"],
        process3Data["DATE"],
        process3Data["TIME"],
        process3Data["Process 3 Model Code"],
        process3Data["Process 3 S/N"],
        process3Data["Process 3 ID"],
        process3Data["Process 3 NAME"],
        process3Data["Process 3 Regular/Contractual"],
        process3Data["Process 3 Frame Gasket"],
        process3Data["Process 3 Frame Gasket Lot No"],
        process3Data["Process 3 Casing Block"],
        process3Data["Process 3 Casing Block Lot No"],
        process3Data["Process 3 Casing Gasket"],
        process3Data["Process 3 Casing Gasket Lot No"],
        process3Data["Process 3 M4x16 Screw 1"],
        process3Data["Process 3 M4x16 Screw 1 Lot No"],
        process3Data["Process 3 M4x16 Screw 2"],
        process3Data["Process 3 M4x16 Screw 2 Lot No"],
        process3Data["Process 3 Ball Cushion"],
        process3Data["Process 3 Ball Cushion Lot No"],
        process3Data["Process 3 Frame Cover"],
        process3Data["Process 3 Frame Cover Lot No"],
        process3Data["Process 3 Partition Board"],
        process3Data["Process 3 Partition Board Lot No"],

        process3Data2["Process 3 Built In Tube 1"],
        process3Data2["Process 3 Built In Tube 1 Lot No"],
        process3Data2["Process 3 Built In Tube 2"],
        process3Data2["Process 3 Built In Tube 2 Lot No"],
        process3Data2["Process 3 Head Cover"],
        process3Data2["Process 3 Head Cover Lot No"],
        process3Data2["Process 3 Casing Packing"],
        process3Data2["Process 3 Casing Packing Lot No"],
        process3Data2["Process 3 M4x12 Screw"],
        process3Data2["Process 3 M4x12 Screw Lot No"],
        process3Data2["Process 3 Csb L"],
        process3Data2["Process 3 Csb L Lot No"],
        process3Data2["Process 3 Csb R"],
        process3Data2["Process 3 Csb R Lot No"],
        process3Data2["Process 3 Head Packing"],
        process3Data2["Process 3 Head Packing Lot No"],
        process3Data2["Process 3 ST"],
        process3Data2["Process 3 Actual Time"],
        process3Data2["Process 3 NG Cause"],
        process3Data2["Process 3 Repaired Action"]

        
    ], axis=1, ignore_index=True)

    mergedProcess3Data.columns = [
        "DATA No",
        "DATE",
        "TIME",
        "Process 3 Model Code",
        "Process 3 S/N",
        "Process 3 ID",
        "Process 3 NAME",
        "Process 3 Regular/Contractual",
        "Process 3 Frame Gasket",
        "Process 3 Frame Gasket Lot No",
        "Process 3 Casing Block",
        "Process 3 Casing Block Lot No",
        "Process 3 Casing Gasket",
        "Process 3 Casing Gasket Lot No",
        "Process 3 M4x16 Screw 1",
        "Process 3 M4x16 Screw 1 Lot No",
        "Process 3 M4x16 Screw 2",
        "Process 3 M4x16 Screw 2 Lot No",
        "Process 3 Ball Cushion",
        "Process 3 Ball Cushion Lot No",
        "Process 3 Frame Cover",
        "Process 3 Frame Cover Lot No",
        "Process 3 Partition Board",
        "Process 3 Partition Board Lot No",
        "Process 3 Built In Tube 1",
        "Process 3 Built In Tube 1 Lot No",
        "Process 3 Built In Tube 2",
        "Process 3 Built In Tube 2 Lot No",
        "Process 3 Head Cover",
        "Process 3 Head Cover Lot No",
        "Process 3 Casing Packing",
        "Process 3 Casing Packing Lot No",
        "Process 3 M4x12 Screw",
        "Process 3 M4x12 Screw Lot No",
        "Process 3 Csb L",
        "Process 3 Csb L Lot No",
        "Process 3 Csb R",
        "Process 3 Csb R Lot No",
        "Process 3 Head Packing",
        "Process 3 Head Packing Lot No",
        "Process 3 ST",
        "Process 3 Actual Time",
        "Process 3 NG Cause",
        "Process 3 Repaired Action"
    ]

# %%
def WriteCsv(excelData):
    fileDirectory = (r'\\192.168.2.10\csv\csv\VT3')
    # fileDirectory = (r'C:\Users\c.raniel\Documents')
    os.chdir(fileDirectory)
    print(os.getcwd())

    print("Creating New File")
    #Create Excel File
    newValue = pd.concat([excelData], axis = 0, ignore_index = True)
    wireFrame = newValue
    wireFrame.to_csv(f"log000_3.csv", index = False)

# %%
GetDateToday()

try:
    process3OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT3\log000_3_1.csv')
    isFileExist = True
except:
    print("No File Exist")
    isFileExist = False
    
while True:
    GetDateToday()
    print("Reading Process 3")
    try:
        process3CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT3\log000_3_1.csv')
    except:
        print

    try:
        if process3CurrentFile != process3OrigFile:
            ReadProcess3Csv()
            MergedProcess3()
            WriteCsv(mergedProcess3Data)
            print("Cnanges Detected")
            process3OrigFile = process3CurrentFile
    except:
        print

    if not isFileExist:
        try:
            process3OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT3\log000_3_1.csv')
            isFileExist = True
        except:
            print("No File Exist")
            isFileExist = False

    readCount += 1
    if readCount >= 10:
        os.system('cls')
        readCount = 0
    time.sleep(1)


