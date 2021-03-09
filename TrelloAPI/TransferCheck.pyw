##This script gets Trello cards sent do 'Done' list
##in the last four days and matches it to folders.
##if there is a match it will produce a list of these card
##so that there can be check if the work was truely done

import pandas as pd
import requests
import json
import datetime
from easygui import *
import pyautogui
import os
import keyboard
import time
import win32com.client as win32

##get list of folders in current directory
folders = os.listdir()


##folders_dict = dict()
##for file in folders:
##    folders_dict.update({file:datetime.datetime.fromtimestamp(os.path.getctime(file)).strftime("%Y-%m-%d")})

##list of branches and emails
Strand = ['001', 'strand@pnastrandgroup.co.za']
Colours = ['002', 'colours@pnastrandgroup.co.za']
Swes = ['003', 'swes@pnastrandgroup.co.za']
Eikestad = ['004', 'eikestad@pnastrandgroup.co.za']
Stellenbosch = ['005', "stelsquare@pnastrandgroup.co.za"]
Commercial = ['COM', 'sales@pnastrandgroup.co.za']
Commercial2 = ['BTS', 'sales2@pnastrandgroup.co.za']
Sanctuary = ['006', 'sanctuary@pnastrandgroup.co.za']
Hermanus = ['007', 'whalecoast@pnastrandgroup.co.za']
Gordonsbay = ['009', 'gordonsbay@pnastrandgroup.co.za']
Swellendam = ['010', 'swellendam@pnastrandgroup.co.za']
Warehouse = ['888', 'warehouse@pnastrandgroup.co.za']


BranchList = {"strand@pnastrandgroup.co.za":Strand, "colours@pnastrandgroup.co.za":Colours, \
              "swes@pnastrandgroup.co.za":Swes, "eikestad@pnastrandgroup.co.za":Eikestad, "stelsquare@pnastrandgroup.co.za":Stellenbosch, \
              "sales@pnastrandgroup.co.za":Commercial, "sales2@pnastrandgroup.co.za":Commercial2, \
              "sanctuary@pnastrandgroup.co.za":Sanctuary, "whalecoast@pnastrandgroup.co.za":Hermanus, \
               "gordonsbay@pnastrandgroup.co.za":Gordonsbay, "swellendam@pnastrandgroup.co.za":Swellendam, \
              "warehouse@pnastrandgroup.co.za":Warehouse}

key1 = # removed for security purposes
token1 = # removed for security purposes
keytok = {'key':key1, 'token':token1}

##dictionary of user IDs
Members = {"warehouse@pnastrandgroup.co.za":"58da5eac5891d87ac1c2fe58",
           "strand@pnastrandgroup.co.za":"550bf4bc2094ac8b5df163dc",
           "swes@pnastrandgroup.co.za":"550fd08951e882750fa5a557",
           "sanctuary@pnastrandgroup.co.za":"57f36c796b25693e34822da8",
           "gordonsbay@pnastrandgroup.co.za":"5a283ea3be6c95d3fb2348ae",
           "sales@pnastrandgroup.co.za":"561f84f36d45d07ffc3b4dec",
           "colours@pnastrandgroup.co.za":"550aa531b66b3a41f9cbeb50",
           "whalecoast@pnastrandgroup.co.za":"5a1d98407eb93a0ddf5a0867",
           "eikestad@pnastrandgroup.co.za":"5df8704ca18b9f4d9a31c958",
           "sales2@pnastrandgroup.co.za":"5c3711bcd127c4444bd74167",
           "swellendam@pnastrandgroup.co.za":"5c98854c0ca5cd64c112fd47",
           "gary@pnastrandgroup.co.za":"5cde4964111d3a8fb753b4b1",
           "it@pnastrandgroup.co.za":"5aa2605a7d8fc2c6292b5e51",
           "liana@pnastrandgroup.co.za":"5b6db96e26f22863cc775dd7",
           "stelsquare@pnastrandgroup.co.za":"60068017690baa8596c41612"}

##dictionary of list IDs
myList = {"To Do":"55042f9dd8eec487931646ea",
          "Doing":"55042fa1c1beafd66acdf60f",
          "Done":"58163ee99eb2ae47a2be5e85"}

memList = []

##get main data from trello Activities log
response2 = requests.get("https://api.trello.com/1/boards/dRBJIaTF/actions/?limit=1000&key="+key1+"&token="+token1)
data2 = response2.json()
print(data2[2])
# print(data2[4])


####loop through data and find cards marked as done in the past 4 days
for x in range(len(data2)):
##    if card is updated still active eg not achived
    if data2[x]["type"] == 'updateCard' and "closed" not in data2[x]["data"]["card"]:
##        card can be found in two places depending on the mood of Trello
##        this find the correct path
        try:
            data2[x]["data"]["list"]["name"]
            y = 'list'
        except:
            data2[x]["data"]["listAfter"]["name"]
            y = 'listAfter'

##        if card is marked as 'done'
        if data2[x]["data"][y]["name"] == 'Done':
            if datetime.datetime.strptime(data2[x]["date"][:10], "%Y-%m-%d").date() >= datetime.date.today()-datetime.timedelta(4):
##                check if card is a folder in directory
                if data2[x]["data"]["card"]["name"] in folders:
##                    add to memlist
                    if [data2[x]["date"][:10],
                    data2[x]["data"]["card"]["name"],
                    list(Members.keys())[list(Members.values()).index(data2[x]["memberCreator"]["id"])]] not in memList:
                        memList.append([data2[x]["date"][:10],
                        data2[x]["data"]["card"]["name"],
                        list(Members.keys())[list(Members.values()).index(data2[x]["memberCreator"]["id"])]])

##prep details from memlist for UI
choices = []
choice = {}
for x in range(len(memList)):
    choices.append(memList[x])
    choice.update({str(memList[x]) : memList[x]})
#print(choice)

##UI to show list
Branches = multchoicebox("choose", title= "hello", choices=choices)
# print(Branches)




