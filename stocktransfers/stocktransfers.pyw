##This script is needed to filter stock transfer details and create new
##sheets only relevant to specific branches.
##It also creates txt files with the list of stock needed for
##a additional check compatible with SQL.

import sys
import pandas as pd
import openpyxl
import pyautogui
import os

##get file and folder details from user
filename = pyautogui.prompt(text='File Name?', title='Transfers', default='')
name = pyautogui.prompt(text='Foldername Name?', title='Transfers', default='')
name=name.replace("/", "")


file = pd.read_excel(rf"C:\Users\User\Desktop\{filename}.xlsx", converters={'Branch':str})

##create list of relevant branches
branches = []
for desc, cd, bc, br, ac in zip(file['Description'], file['Code'], file['Barcode'], file['Branch'], file['Action']):
    if ac <0 and br not in branches:
        branches.append(br)
#print(branches)

##loop through branches and get relevant stock details
for branch in branches:

##    get list of stock items
    x = []
    for desc, cd, bc, br, ac in zip(file['Description'], file['Code'], file['Barcode'], file['Branch'], file['Action']):
        if br == branch and ac < 0:
            #print(desc, cd, bc, br, ac)
            x.append(desc)

##    prep list for dataframe with stock details for current branch
    df = []
    for desc, cd, bc, br, ac in zip(file['Description'], file['Code'], file['Barcode'], file['Branch'], file['Action']):
        for y in x:
            if desc == y:
                if len(str(br)) == 1:
                    br = '00'+str(br)
                if len(str(br)) == 2:
                    br = '0'+ str(br)
                df.append([desc, str(cd), str(bc), br, ac])

##create dataframe from prepared list and sort
    df1 = pd.DataFrame(df, columns=['Description', 'Code', 'Barcode', 'Branch', 'Action'])
    df1 = df1.sort_values(by=['Description', 'Branch'])

    if len(str(branch)) == 1:
        branchStr = '00'+str(branch)
    if len(str(branch)) == 2:
        branchStr = '0'+ str(branch)
    if len(str(branch)) == 3:
        branchStr = str(branch)

##    check if directory exists and save dataframe to excel
    folders  = os.listdir("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\")
    if name in folders:
        df1.to_excel("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx", index=False)
        wb = openpyxl.load_workbook("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx")
        sheet = wb.active
        sheet.column_dimensions['A'].width = 50
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 30
        wb.save("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx")
    else:
##        if directory does not exist, create it, then save dataframe to excel
        os.mkdir("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name)
        df1.to_excel("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx", index=False)
        wb = openpyxl.load_workbook("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx")
        sheet = wb.active
        sheet.column_dimensions['A'].width = 50
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 30
        wb.save("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx")

##Here I get the stock codes from the new sheet and save it to
##txt file as string that IQ Retail SQL will accpet
    data = pd.read_excel("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".xlsx", converters={'Branch':str, 'Code':str})

    IQ_String = ""

    for code, br in zip(data['Code'], data['Branch']):
        code = str(code)
        print(code, br, branch)
        if branch == br:
            print(code, br)
            if "'" in code:
                code = code.replace("'", "''")
            IQ_String += "'"+code+"', "

    # for code, br in zip(data['Code'], data['Branch']):
    #     code = str(code)
    #     print(code, br)
    #     if branch == 'BTS':
    #         print(code, br)
    #         if "'" in code:
    #             code = code.replace("'", "''")
    #         IQ_String += "'"+code+"', "
    #         continue
    #     elif str(branch)[-1] == '0':
    #         print('010')
    #         if int(br) == 10:
    #             print(code, br)
    #             if "'" in code:
    #                 code = code.replace("'", "''")
    #             IQ_String += "'"+code+"', "
    #     elif str(branch)[-1] == '8':
    #         if int(br) == int(branch):
    #             print(code, br)
    #             if "'" in code:
    #                 code = code.replace("'", "''")
    #             IQ_String += "'"+code+"', "
    #     else:
    #         if int(br) == int(str(branch)[-1]):
    #             print(code, br)
    #             if "'" in code:
    #                 code = code.replace("'", "''")
    #             IQ_String += "'"+code+"', "

    
    print(IQ_String[:-2])
    newfile = open("C:\\Users\\User\\Desktop\\Python IQ\\TRANSFERCHECK\\"+name+"\\"+branchStr+".txt", "a")
    newfile.write(IQ_String[:-2])
    newfile.close()
