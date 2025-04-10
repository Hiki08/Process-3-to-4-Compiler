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
process4Data = ""
process4Data2 = ""
mergedProcess4Data = ""
isFileExist = ""
readCount = 0

# %%
def GetDateToday():
    global dateToday

    dateToday = datetime.datetime.today()
    dateToday = dateToday.strftime('%y%m%d')

# %%
def ReadProcess4Csv():
    global process4Data
    global process4Data2

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    process4Directory = (r'\\192.168.2.10\csv\csv\VT4')
    os.chdir(process4Directory)
    
    process4Data = pd.read_csv(f'log000_4_1.csv', encoding='latin1', skiprows=0)
    # process3Data = pd.read_csv(f'log000_3_240814.csv', encoding='latin1', skiprows=0)

    process4Directory2 = (r'\\192.168.2.10\csv\csv\VT4')
    os.chdir(process4Directory2)
    
    process4Data2 = pd.read_csv(f'log000_4_2.csv', encoding='latin1', skiprows=0)
    # process3Data2 = pd.read_csv(f'log000_3_2_240814.csv', encoding='latin1', skiprows=0)

# %%
def MergedProcess4():
    global process4Data
    global process4Data2
    global mergedProcess4Data

    mergedProcess4Data = pd.concat([
        process4Data["DATA No"],
        process4Data["DATE"],
        process4Data["TIME"],
        process4Data["Process 4 Model Code"],
        process4Data["Process 4 S/N"],
        process4Data["Process 4 ID"],
        process4Data["Process 4 NAME"],
        process4Data["Process 4 Regular/Contractual"],
        process4Data["Process 4 Tank"],
        process4Data["Process 4 Tank Lot No"],
        process4Data["Process 4 Upper Housing"],
        process4Data["Process 4 Upper Housing Lot No"],
        process4Data["Process 4 Cord Hook"],
        process4Data["Process 4 Cord Hook Lot No"],
        process4Data["Process 4 M4x16 Screw"],
        process4Data["Process 4 M4x16 Screw Lot No"],
        process4Data["Process 4 Tank Gasket"],
        process4Data["Process 4 Tank Gasket Lot No"],
        process4Data["Process 4 Tank Cover"],
        process4Data["Process 4 Tank Cover Lot No"],
        process4Data["Process 4 Housing Gasket"],
        process4Data["Process 4 Housing Gasket Lot No"],
        process4Data["Process 4 M4x40 Screw"],
        process4Data["Process 4 M4x40 Screw Lot No"],

        process4Data2["Process 4 PartitionGasket"],
        process4Data2["Process 4 PartitionGasket Lot No"],
        process4Data2["Process 4 M4x12 Screw"],
        process4Data2["Process 4 M4x12 Screw Lot No"],
        process4Data2["Process 4 Muffler"],
        process4Data2["Process 4 Muffler Lot No"],
        process4Data2["Process 4 Muffler Gasket"],
        process4Data2["Process 4 Muffler Gasket Lot No"],
        process4Data2["Process 4 VCR"],
        process4Data2["Process 4 VCR Lot No"],
        process4Data2["Process 4 ST"],
        process4Data2["Process 4 Actual Time"],
        process4Data2["Process 4 NG Cause"],
        process4Data2["Process 4 Repaired Action"]
    ], axis=1, ignore_index=True)

    mergedProcess4Data.columns = [
        "DATA No",
        "DATE",
        "TIME",
        "Process 4 Model Code",
        "Process 4 S/N",
        "Process 4 ID",
        "Process 4 NAME",
        "Process 4 Regular/Contractual",
        "Process 4 Tank",
        "Process 4 Tank Lot No",
        "Process 4 Upper Housing",
        "Process 4 Upper Housing Lot No",
        "Process 4 Cord Hook",
        "Process 4 Cord Hook Lot No",
        "Process 4 M4x16 Screw",
        "Process 4 M4x16 Screw Lot No",
        "Process 4 Tank Gasket",
        "Process 4 Tank Gasket Lot No",
        "Process 4 Tank Cover",
        "Process 4 Tank Cover Lot No",
        "Process 4 Housing Gasket",
        "Process 4 Housing Gasket Lot No",
        "Process 4 M4x40 Screw",
        "Process 4 M4x40 Screw Lot No",
        "Process 4 PartitionGasket",
        "Process 4 PartitionGasket Lot No",
        "Process 4 M4x12 Screw",
        "Process 4 M4x12 Screw Lot No",
        "Process 4 Muffler",
        "Process 4 Muffler Lot No",
        "Process 4 Muffler Gasket",
        "Process 4 Muffler Gasket Lot No",
        "Process 4 VCR",
        "Process 4 VCR Lot No",
        "Process 4 ST",
        "Process 4 Actual Time",
        "Process 4 NG Cause",
        "Process 4 Repaired Action"
    ]

# %%
def WriteCsv(excelData):
    fileDirectory = (r'\\192.168.2.10\csv\csv\VT4')
    # fileDirectory = (r'C:\Users\c.raniel\Documents')
    os.chdir(fileDirectory)
    print(os.getcwd())

    print("Creating New File")
    #Create Excel File
    newValue = pd.concat([excelData], axis = 0, ignore_index = True)
    wireFrame = newValue
    wireFrame.to_csv(f"log000_4.csv", index = False)

# %%
GetDateToday()

try:
    process4OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT4\log000_4_1.csv')
    isFileExist = True
except:
    print("No File Exist")
    isFileExist = False
    
while True:
    GetDateToday()
    print("Reading Process 4")
    try:
        process4CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT4\log000_4_1.csv')
    except:
        print

    try:
        if process4CurrentFile != process4OrigFile:
            ReadProcess4Csv()
            MergedProcess4()
            WriteCsv(mergedProcess4Data)
            print("Cnanges Detected")
            process4OrigFile = process4CurrentFile
    except:
        print

    if not isFileExist:
        try:
            process4OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT4\log000_4_1.csv')
            isFileExist = True
        except:
            print("No File Exist")
            isFileExist = False

    readCount += 1
    if readCount >= 10:
        os.system('cls')
        readCount = 0
    time.sleep(1)


