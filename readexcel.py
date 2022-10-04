import batterystats as bs
import pandas as pd
import numpy as np
import xlsxwriter as xl
import openpyxl
import xlwt
import xlrd

def readexcel(excel):

    testdata =  pd.read_excel(excel, usecols="C:K", dtype={'module_sn' : str, 'cathode_bond_R'  : np.float64, 'anode_bond_R' : np.float64})
    #reorganize data by row and column first, not pass/fail
    testdata.sort_values(by=['cell_row', 'cell_column'], inplace= True)
    testdata.dropna(subset=['module_sn'], inplace= True)

    td = []
    tempcol = 0
    temprow = 0
    num_battery = 0
    n = 1
    bn = 1

    for x in testdata.index:
        if testdata.loc[x, "cell_row"] == 1:
            if testdata.loc[x, "cell_column"] % 2 == 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 7:
            if testdata.loc[x, "cell_column"] % 2 == 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 13:
            if testdata.loc[x, "cell_column"] % 2 == 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 19:
            if testdata.loc[x, "cell_column"] % 2 == 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 25:
            if testdata.loc[x, "cell_column"] % 2 == 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 4:
            if testdata.loc[x, "cell_column"] % 2 != 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 10:
            if testdata.loc[x, "cell_column"] % 2 != 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 16:
            if testdata.loc[x, "cell_column"] % 2 != 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 22:
            if testdata.loc[x, "cell_column"] % 2 != 0:
                testdata.loc[x, "module_sn"] = 'RIM'
        elif testdata.loc[x, "cell_row"] == 28:  
            if testdata.loc[x, "cell_column"] % 2 != 0:
                testdata.loc[x, "module_sn"] = 'RIM'

        if testdata.loc[x, "cell_column"] == tempcol and testdata.loc[x, "cell_row"] == temprow:
            num_battery = num_battery + 1
    for f in testdata.index:
        if n <= num_battery:
            td.append(n)
            n = n + 1
        else:
            n = 1
            td.append(n)
            n = n + 1

    testdata.insert(0, "battery_num", td)

    with pd.ExcelWriter('stats.xlsx', engine= 'xlsxwriter') as writer:
        for i in range(num_battery):
        
            #Need to save rest of values other than first into testdata and make testdata1 the drop
            #duplicates that gets passed to bs
            testdata1 = testdata.drop_duplicates(subset=['cell_row', 'cell_column'], keep = 'first')
            bs.stats(writer, testdata1, bn)
            testdata.drop(testdata[testdata['battery_num'] == bn].index, inplace = True)
            bn = bn + 1
        