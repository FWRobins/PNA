##This UI Automation was used to delete old stock items in our IQ Retail
##system once a month for each branch.
##This is no longer in use


import pyautogui
import os
import time
import keyboard
import datetime
from datetime import date, timedelta
from easygui import *

myFilter= "(ONHAND = 0) AND (DATE_LM NOT BETWEEN '2016-09-01' AND '2018-09-01') AND ((LINECOLOURTYPE = 8) OR (DESCRIPT LIKE 'ZZ%') OR (DESCRIPT LIKE '##%'))"

Strand = ['001', 'strand@pnastrand.co.za']
Colours = ['002', 'colours@pnastrand.co.za']
Swes = ['003', 'swes@pnastrand.co.za']
Eikestad = ['004', 'eikestad@pnastrand.co.za']
Commercial = ['005', 'sales@pnastrand.co.za']
Sanctuary = ['006', 'sanctuary@pnastrand.co.za']
Hermanus = ['007', 'whalecoast@pnastrand.co.za']
Gordonsbay = ['009', 'gordonsbay@pnastrand.co.za']
BTS = ['010', 'swes@pnastrand.co.za']
Warehouse = ['888', 'swes@pnastrand.co.za']
BranchList = {"Strand":Strand, "Colours":Colours, "Swes":Swes, "Eikestad":Eikestad, \
              "Commercial":Commercial, "Sanctuary":Sanctuary, \
              "Hermanus":Hermanus, "Gordonsbay":Gordonsbay, "BTS":BTS, "Warehouse":Warehouse}

Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Commercial", "Sanctuary", "Hermanus", \
                        "Gordonsbay", "BTS", "Warehouse"])






for Branch in Branches:
##    Branches = Branch    
    keyboard.send('windows')
    time.sleep(2)
    pyautogui.click(321, 748)       #rdp tasbar
    time.sleep(2)
    keyboard.send('windows')
    time.sleep(2)
    pyautogui.click(191, 755)       #iq in rdp taskbar
    time.sleep(2)
    while True:
        img = pyautogui.locateOnScreen('Utilities.png', region=(330, 10, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    while True:
        img = pyautogui.locateOnScreen('SelectCompany.png', region=(5, 40, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    time.sleep(2)
    keyboard.write(BranchList[Branch][0], 0.25)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(5)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.click(246, 26)        #stock
    time.sleep(2)
    pyautogui.click(56,72)          #maintenance
    time.sleep(2)
    pyautogui.click(585, 726)       #filter
    time.sleep(0.5)
    pyautogui.click(547, 227)       #advanced filter
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(1)
    keyboard.write(myFilter, 0.3)
    time.sleep(1)
    keyboard.send('tab')
    time.sleep(1)
    keyboard.write('a', 0.25)
    time.sleep(5)
    pyautogui.click(1051, 727)      #delete
    time.sleep(2)
    pyautogui.click(1059, 684)      #all
    time.sleep(2)
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.moveTo(10 ,10)
    while True:
        img = pyautogui.locateOnScreen('accept.png', region=(1040, 655, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    while True:
        img = pyautogui.locateOnScreen('delete.png', region=(550, 335, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(2)
    keyboard.send('esc')
    time.sleep(2)
    keyboard.send('esc')
    time.sleep(2)
    pyautogui.click(1317, 10)       #minimize IQ
    time.sleep(2)
    pyautogui.click(882, 12)        #minimize RDP
    time.sleep(2)
    pyautogui.click(882, 12)        #select desktop
    time.sleep(2)
    continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
    if continueScript == 'No':
        quit()
    else:
        continue
    

