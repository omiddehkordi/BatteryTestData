import pandas as pd
import numpy as np
import xlsxwriter as xl
import openpyxl
import xlwt
import xlrd

def toexcel(writer, brick, num):
    num_str = str(num)
    brick.to_excel(writer, sheet_name= "Battery " + num_str, index= False)
    workbook = writer.book
    worksheet = writer.sheets['Battery ' + num_str]
    worksheet.set_column(0, 10, 20)
