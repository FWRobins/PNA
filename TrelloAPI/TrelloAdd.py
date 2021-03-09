##this script is used to add multiple cards to the Trello board

import requests
import json
import datetime
from easygui import *
import pyautogui

##request relevant branched from user
Branches = multchoicebox("choose", title= "hello", choices=
                       ["strand@pnastrandgroup.co.za", "colours@pnastrandgroup.co.za", 
                        "swes@pnastrandgroup.co.za", "eikestad@pnastrandgroup.co.za",
                        "sales@pnastrandgroup.co.za","sales2@pnastrandgroup.co.za",
                        "sanctuary@pnastrandgroup.co.za", "whalecoast@pnastrandgroup.co.za", 
                        "gbay@pnastrandgroup.co.za", "warehouse@pnastrandgroup.co.za", "swellendam@pnastrandgroup.co.za"])

key1 = # removed for security purposes

token1 = # removed for security purposes

##dictionary of branch IDs
Members = {"warehouse@pnastrandgroup.co.za":"58da5eac5891d87ac1c2fe58",
           "strand@pnastrandgroup.co.za":"550bf4bc2094ac8b5df163dc",
           "swes@pnastrandgroup.co.za":"550fd08951e882750fa5a557",
           "sanctuary@pnastrandgroup.co.za":"57f36c796b25693e34822da8",
           "gbay@pnastrandgroup.co.za":"5a283ea3be6c95d3fb2348ae",
           "sales@pnastrandgroup.co.za":"561f84f36d45d07ffc3b4dec",
           "colours@pnastrandgroup.co.za":"550aa531b66b3a41f9cbeb50",
           "whalecoast@pnastrandgroup.co.za":"5a1d98407eb93a0ddf5a0867",
           "eikestad@pnastrandgroup.co.za":"550bdfdcab34abe24a3e15f2",
           "sales2@pnastrandgroup.co.za":"5c3711bcd127c4444bd74167",
           "swellendam@pnastrandgroup.co.za":"5c98854c0ca5cd64c112fd47"}

##dictionary of list IDs
myList = {"To Do":"55042f9dd8eec487931646ea",
          "Doing":"55042fa1c1beafd66acdf60f",
          "Done":"58163ee99eb2ae47a2be5e85"}

##request trallo card name and due date from user
Name = pyautogui.prompt(text='Transfer name?', title='Card Name', default='')
dueDate = pyautogui.prompt(text='Due Date?', title='Due Date', default='YYYY-MM-DD')

url = "https://api.trello.com/1/cards"

##main loop to add cards with given details for selected branches
for x in Branches:
    querystring = {"name":Name,"due":dueDate,"idList":myList["To Do"],"idMembers":Members[str(x)],"keepFromSource":"all","key":key1,"token":token1}
    response = requests.request("POST", url, params=querystring)
