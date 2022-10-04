#!/usr/bin/env python3
import batterystats as bs
import readexcel as re
import tkinter as tk
from tkinter.filedialog import askopenfilename
import openpyxl
import xlwt
import xlrd

#GUI using Tkinter

window = tk.Tk()

window.withdraw()

try:
    user_input = askopenfilename()
    
    #Calling main functions
    #bs.stats(user_input)

    re.readexcel(user_input)
except:
    print("Inaccurate Filepath or Canceled Process")