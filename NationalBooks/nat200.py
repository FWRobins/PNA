##This creates a SQL querie to be executed on out IQ Retail database
##of the top 200 national best selling boooks

import pandas as pd

##get ISBNs from excel file
basefile = pd.read_excel('nat.xlsx', usecols=[4], names=["ISBN"])

##get top 200 ISBNs
isbns = ""
for isbn in range(200):
    isbns += "'"+str(basefile.ISBN[isbn])+"',"


file = open('nat.csv', "w")

##create SQL querie for each brach DB file
branches = ['000','001', '002', '003', '004', '005', '006', '007', '009', '010', '888', 'BTS', 'COM']
for branch in branches:
    file.write(rf'update "C:\IQRetail\IQEnterprise\{branch}\STOCK.dat" set ABCclass=1 where code in (')
    file.write(isbns[:-1])
    file.write(') commit interval 5; ')
file.close()

input("press enter to close.")
