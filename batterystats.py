import pandas as pd
import numpy as np
import xlsxwriter as xl
import writeexcel as we
import openpyxl
import xlwt
import xlrd


def stats(writer, testdata, bn):
    
    brick = pd.DataFrame(columns = ['cell_row_range', 'cell_column_range', 'halfbrick_num', 'color', 'cat_or_an', 'min', 'max', 'average', 'std_dev'])

    col_it = 0
    iterat = 0
    column = 0
    hb_num = 1

    row = []

    lowlbcat = []
    lowlban = []
    lowgreencat = []
    lowgreenan = []
    highlbcat = []
    highlban = []
    highgreencat = []
    highgreenan = []

    lower = True

    for x in testdata.index:
        if lower:
            if iterat % 2 == 0 and testdata.loc[x, "module_sn"] != 'RIM':
                if testdata.loc[x, "cathode_bond_test_fp"] == 1:
                    lowlbcat.append(testdata.loc[x, "cathode_bond_R"])
                if testdata.loc[x, "anode_bond_test_fp"] == 1:
                    lowlban.append(testdata.loc[x, "anode_bond_R"])
            elif iterat % 2 != 0 and testdata.loc[x, "module_sn"] != 'RIM':
                if testdata.loc[x, "cathode_bond_test_fp"] == 1:
                    lowgreencat.append(testdata.loc[x, "cathode_bond_R"])
                if testdata.loc[x, "anode_bond_test_fp"] == 1:
                    lowgreenan.append(testdata.loc[x, "anode_bond_R"])
        else:
            if iterat % 2 == 0 and testdata.loc[x, "module_sn"] != 'RIM':
                if testdata.loc[x, "cathode_bond_test_fp"] == 1:
                    highlbcat.append(testdata.loc[x, "cathode_bond_R"])
                if testdata.loc[x, "anode_bond_test_fp"] == 1:
                    highlban.append(testdata.loc[x, "anode_bond_R"])
            elif iterat % 2 != 0 and testdata.loc[x, "module_sn"] != 'RIM':
                if testdata.loc[x, "cathode_bond_test_fp"] == 1:
                    highgreencat.append(testdata.loc[x, "cathode_bond_R"])
                if testdata.loc[x, "anode_bond_test_fp"] == 1:
                    highgreenan.append(testdata.loc[x, "anode_bond_R"])

        iterat = iterat + 1
        col_it = col_it + 1
        if iterat % 5 == 0:
            if lower:
                lower = False
            else:
                lower = True
                row.append(column)
                column = column + 1
                

        if col_it % 30 == 0:
            if len(lowlbcat) > 0:
                lowlbcatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'cat (+)', "min" : round(min(lowlbcat), 6), "max" : round(max(lowlbcat), 6), "average" : round(np.mean(lowlbcat), 6), "std_dev" : round(np.std(lowlbcat), 6)}, index =[0])
            else:
                lowlbcatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'cat (+)'}, index =[0])
            if len(lowlban) > 0:
                lowlbanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'an (-)', "min" : round(min(lowlban), 6), "max" : round(max(lowlban), 6), "average" : round(np.mean(lowlban),  6), "std_dev" : round(np.std(lowlban),  6)}, index =[0])
            else:
                lowlbanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'an (-)'}, index =[0])
            if len(lowgreencat) > 0:
                lowgreencatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'cat (+)', "min" : round(min(lowgreencat), 6), "max" : round(max(lowgreencat), 6), "average" : round(np.mean(lowgreencat), 6), "std_dev" : round(np.std(lowgreencat), 6)}, index =[0])
            else:
                lowgreencatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'cat (+)'}, index =[0])
            if len(lowgreenan) > 0:
                lowgreenanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'an (-)', "min" : round(min(lowgreenan), 6), "max" : round(max(lowgreenan), 6), "average" : round(np.mean(lowgreenan), 6), "std_dev" : round(np.std(lowgreenan), 6)}, index =[0])
            else:
                lowgreenanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '0-4', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'an (-)'}, index =[0])

            if len(lowlbcatstats) > 0:
                brick = brick.append(lowlbcatstats)
            if len(lowlbanstats) > 0:
                brick = brick.append(lowlbanstats)
            if len(lowgreencatstats) > 0:
                brick = brick.append(lowgreencatstats)
            if len(lowgreenanstats) > 0:
                brick = brick.append(lowgreenanstats)

            hb_num = hb_num + 1

            if len(highlbcat) > 0:
                highlbcatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'cat (+)', "min" : round(min(highlbcat), 6), "max" :round(max(highlbcat), 6), "average" : round(np.mean(highlbcat),  6), "std_dev" : round(np.std(highlbcat), 6)}, index =[0])
            else:
                highlbcatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'cat (+)'}, index =[0])
            if len(highlban) > 0:
                highlbanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'an (-)', "min" : round(min(highlban), 6), "max" : round(max(highlban), 6), "average" : round(np.mean(highlban), 6), "std_dev" : round(np.std(highlban), 6)}, index =[0])
            else:
                highlbanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'blue', "cat_or_an" : 'an (-)'}, index =[0])
            if len(highgreencat) > 0:
                highgreencatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'cat (+)', "min" : round(min(highgreencat), 6), "max" : round(max(highgreencat), 6), "average" : round(np.mean(highgreencat), 6), "std_dev" : round(np.std(highgreencat), 6)}, index =[0])
            else:
                highgreencatstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'cat (+)'}, index =[0])
            if len(highgreenan) > 0:    
                highgreenanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'an (-)', "min" : round(min(highgreenan), 6), "max" : round(max(highgreenan), 6), "average" : round(np.mean(highgreenan), 6), "std_dev" : round(np.std(highgreenan), 6)}, index =[0])
            else:
                highgreenanstats = pd.DataFrame({"cell_row_range" : str(row[0]) + '-' + str(row[-1]), "cell_column_range" : '5-9', "halfbrick_num" : hb_num, "color" : 'green', "cat_or_an" : 'an (-)'}, index =[0])

            if len(highlbcatstats) > 0:
                brick = brick.append(highlbcatstats)
            if len(highlbanstats) > 0:
                brick = brick.append(highlbanstats)
            if len(highgreencatstats) > 0:
                brick = brick.append(highgreencatstats)
            if len(highgreenanstats) > 0:
                brick = brick.append(highgreenanstats)

            hb_num = hb_num + 1

                
            row.clear()
            lowlbcat.clear()
            lowlban.clear()
            lowgreencat.clear()
            lowgreenan.clear()
            highlbcat.clear()
            highlban.clear()
            highgreencat.clear()
            highgreenan.clear() 
   
    we.toexcel(writer, brick, bn)