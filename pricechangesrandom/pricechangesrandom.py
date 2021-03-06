##This is used to produce a SQL script that in turn will export results
##in .csv format for each branch containing a list of stock
##codes whos prices have been changed.

import sys
import pandas as pd
import openpyxl
import pyautogui
import os

##get file name from user
filename = pyautogui.prompt(text='File Name?', title='Transfers', default='')

##get list of product codes from specified file
basefile = pd.read_csv(rf"C:\Users\User\Desktop\{filename}.CSV", usecols=[0], names=["Code"])
print(basefile)
print(str(basefile.Code[0]))
codes = ""
for code in range(len(basefile)):
    print(basefile.Code[code])
    codes += "'"+str(basefile.Code[code])+"',"
print(codes)

##create new file with SQL script
file = open('priceupdaterandom.csv', "w")

branches = ['001', '002', '003', '004', '006', '007', '009', '010']
for branch in branches:
    file.write(rf'SELECT CODE, DESCRIPT, ONHAND, SELLPINC1 INTO "C:\IQRetail\IQEnterprise\Exports\FWRTABLE" FROM "C:\IQRetail\IQEnterprise\{branch}\Stock.dat" WHERE CODE IN (')
    file.write(codes[:-1])
    file.write(rf') AND ONHAND > 0; EXPORT TABLE "C:\IQRetail\IQEnterprise\Exports\FWRTABLE" TO "C:\IQRetail\IQEnterprise\Exports\{branch}.csv" WITH HEADERS; ')
file.close()
