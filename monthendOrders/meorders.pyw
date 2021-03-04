##this was a script that I converted to a exe for the person
##in charge of stock orders. they would recieve a excel sheet
##for all orders for all branches from all suppliers.
##This creates importable .csv files split by branch and supplier
##this is still in use daily.

import sys
import pandas as pd
import pyautogui
import os

##get file name from user
filename = pyautogui.prompt(text='File Name?', title='Transfers', default='')

file = pd.read_excel(rf"{filename}.xlsx")

file.columns = file.columns.str.replace(" ", "_")

##get list of suppliers and branches
branches = []
suppliers = []
for sup, br in zip(file['Supplier'], file['Branch']):
    if br not in branches:
        branches.append(br)
    if sup not in suppliers:
        suppliers.append(sup)

print(branches)

##branches are in a format of '001' ect, this changes the interger values to string
## would have been better to use onverters in read_excel method
for branch in branches:
    print(len(str(branch)))
    if len(str(branch)) == 1:
       branchStr = '00'+str(branch)
    elif len(str(branch)) == 2:
       branchStr = '0'+ str(branch)
    elif len(str(branch)) == 3:
       branchStr = str(branch)
    print(f"{branchStr}")
    os.mkdir(f"{branchStr}")

##create .csv files split by branch, supplier in importable format
for branch in branches:
    for suplier in suppliers:
        filter = file[(file.Branch == branch) & (file.Supplier == suplier)]
        if len(filter) != 0:
            if len(str(branch)) == 1:
               branchStr = '00'+str(branch)
            elif len(str(branch)) == 2:
               branchStr = '0'+ str(branch)
            elif len(str(branch)) == 3:
               branchStr = str(branch)
            filecsv = open(rf"{branchStr}\ {suplier}.csv", "w")
            filecsv.write("Code, Ord_Quan,\n")


            for row in range(len(filter)):
                filecsv.write(rf"{filter.iloc[row].Code},{filter.iloc[row].Ord_Qty},"+"\n")

##alert user wher done
pyautogui.alert(text='Done', title='Filter Completed', button='OK')
