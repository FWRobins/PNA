##I am in control of the staff acccounts at PNA. this script generates
##.csv files that are importable into our IQ Retail system.

import pyautogui as pg
import pandas as pd

##read data from excel file
file = pd.read_excel("staf.xlsx", converters={"ALTERNATIVE_NAME":str})
#print(file)

##create headers required by IQ Retail debtors journals module
headers = '"Account","Name","Date","Reference","Notes","OrderNumber","Branch","Code","Vatrate","Ldgr Account","Rep","Amount",\n'
#print(headers)

Dr000 = [headers]
Dr001 = [headers]
Dr002 = [headers]
Dr003 = [headers]
Dr004 = [headers]
Dr005 = [headers]
Dr006 = [headers]
Dr007 = [headers]
Dr009 = [headers]
Dr010 = [headers]
DrCOM = [headers]

##generate strings and insert into relevant branch list
print(len(file))
for i in range(len(file)):
    if file['ALTERNATIVE_NAME'][i]=='000':
        Dr000.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"000","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='001':
        Dr001.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"001","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='002':
        Dr002.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"002","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='003':
        Dr003.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"003","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='004':
        Dr004.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"004","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='005':
        Dr005.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"005","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='006':
        Dr006.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"006","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='007':
        Dr007.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"007","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='009':
        Dr009.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"009","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='010':
        Dr010.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"010","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')
    if file['ALTERNATIVE_NAME'][i]=='COM':
        DrCOM.append(rf'"{file["ACCOUNT"][i]}","{file["NAME"][i]}",25/02/2021,"Salary Deduction",,,"COM","JC","0","1600.000.000.00","37","{file["TOTAL"][i]}",'+'\n')




print(Dr000)


##write generated strings to relevant .csv files
file1 = open("Dr000.csv", "w")
file1.writelines(Dr000)
file1.close()
file1 = open("Dr001.csv", "w")
file1.writelines(Dr001)
file1.close()
file1 = open("Dr002.csv", "w")
file1.writelines(Dr002)
file1.close()
file1 = open("Dr003.csv", "w")
file1.writelines(Dr003)
file1.close()
file1 = open("Dr004.csv", "w")
file1.writelines(Dr004)
file1.close()
file1 = open("Dr005.csv", "w")
file1.writelines(Dr005)
file1.close()
file1 = open("Dr006.csv", "w")
file1.writelines(Dr006)
file1.close()
file1 = open("Dr007.csv", "w")
file1.writelines(Dr007)
file1.close()
file1 = open("Dr009.csv", "w")
file1.writelines(Dr009)
file1.close()
file1 = open("Dr010.csv", "w")
file1.writelines(Dr010)
file1.close()
file1 = open("DrCOM.csv", "w")
file1.writelines(DrCOM)
file1.close()

