##this UI Automation script was used to import new prices
##for stock on our IQ Retail system.
##This was used in conjunction with 'PriceUpdateEmail.pyw'
##This is no longer in use and has since been replaced
##with SQL

import pyautogui
import keyboard
import time
import os
from easygui import *

##get file name from user
file = pyautogui.prompt(text='Import file name?', title='import', default='')

##setup dictionary of branch details
Strand = ['001', 'strand@pnastrandgroup.co.za']
Colours = ['002', 'colours@pnastrandgroup.co.za']
Swes = ['003', 'swes@pnastrandgroup.co.za']
Eikestad = ['004', 'eikestad@pnastrandgroup.co.za']
Commercial = ['005', 'sales@pnastrandgroup.co.za']
Sanctuary = ['006', 'sanctuary@pnastrandgroup.co.za']
Hermanus = ['007', 'whalecoast@pnastrandgroup.co.za']
Gordonsbay = ['009', 'gordonsbay@pnastrandgroup.co.za']
Swellendam = ['010', 'swellendam@pnastrandgroup.co.za']
Warehouse = ['888', 'swes@pnastrandgroup.co.za']
BTS = ['BTS', 'sales@pnastransgroup.co.za']
BranchList = {"Strand":Strand, "Colours":Colours, "Swes":Swes, "Eikestad":Eikestad, \
              "Commercial":Commercial, "Sanctuary":Sanctuary, \
              "Hermanus":Hermanus, "Gordonsbay":Gordonsbay, \
              "Swellendam":Swellendam, "Warehouse":Warehouse, "BTS":BTS}

##ask user what branches to loop through
Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Commercial", "Sanctuary", "Hermanus", \
                        "Gordonsbay", "Swellendam", "Warehouse", "BTS"])

##request wait time from user for import time
wait = pyautogui.prompt(text='wait time?', title='import', default='')


    

##The rest is just the main loop for the UI Automation
for Branch in Branches:
    keyboard.send('windows')
    time.sleep(2)
    pyautogui.click(321, 748)       #rdp tasbar
    time.sleep(2)
    keyboard.send('windows')
    time.sleep(2)
    pyautogui.click(191, 755)       #iq in rdp taskbar
    time.sleep(2)
    Branches = Branch
    time.sleep(2)
    pyautogui.click(366, 34)        #utilitys
    pyautogui.click(30, 80)         #change company
    time.sleep(2)
    keyboard.write(BranchList[Branch][0], 0.25)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(5)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.click(246, 26)
    time.sleep(2)
    pyautogui.click(697, 71)
    time.sleep(2)
    pyautogui.click(724, 277)
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.click(236, 706)
    time.sleep(1)
    pyautogui.click(243, 650)
    time.sleep(1)
    pyautogui.click(453, 247)
    time.sleep(1)    
    keyboard.write(file+".csv", 0.25)
    time.sleep(1)
    keyboard.send('enter')
    time.sleep(int(wait))
    time.sleep(5)
    keyboard.press('alt')
    time.sleep(0.5)
    keyboard.press('p')
    time.sleep(2)
    keyboard.release('p')
    keyboard.release('alt')
    time.sleep(1)
    keyboard.send('enter')
    time.sleep(2)
    keyboard.press('alt')
    time.sleep(0.5)
    keyboard.press('a')
    time.sleep(2)
    keyboard.release('a')
    keyboard.release('alt')
    time.sleep(int(wait))
    keyboard.send('enter')
    time.sleep(1)
    keyboard.send('esc')
    time.sleep(2)
    pyautogui.click(1317, 10)           #minimize IQ
    time.sleep(2)
    pyautogui.click(882, 12)            #minimize RDP
    time.sleep(2)
##    print(Branch[2])
    continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
    if continueScript == 'No':
        quit()
    else:
        continue
