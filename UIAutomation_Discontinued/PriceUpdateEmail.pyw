##This UI Automation script was used to get stock
##that had price changes and are onhand at all branches
##and send list to the branch to update their price labels
##This is no longer in use and has since been replaced
##with another python script and SQL

import pyautogui
import datetime
import keyboard
import time
import os
from easygui import *

##get stock info from user
supCode = pyautogui.prompt(text='Please Enter Supplier Code', title='Supplier Code', default='')
modDate = pyautogui.prompt(text='Please Enter Modification Date', title='Modified Date', default=datetime.datetime.now().strftime('%Y-%m-%d'))

##setup filter string for later use
myFilter = "(REGULAR_SU LIKE '"+supCode+"') AND (ONHAND <> 0) AND (MODIFIED > '"+modDate+"') AND (USER = 37)"

##this was used on 2 PCs, rdp min location was at diff positions
if os.environ['computername'] == 'USER-PC':
    rdpmin = (910, 12)
else:
    rdpmin = (882, 12)

##setup dictionary of branches
Strand = ['001', 'strand@pnastrandgroup.co.za']
Colours = ['002', 'colours@pnastrandgroup.co.za']
Swes = ['003', 'swes@pnastrandgroup.co.za']
Eikestad = ['004', 'eikestad@pnastrandgroup.co.za']
Sanctuary = ['006', 'sanctuary@pnastrandgroup.co.za']
Hermanus = ['007', 'whalecoast@pnastrandgroup.co.za']
Gordonsbay = ['009', 'gordonsbay@pnastrandgroup.co.za']
Swellendam = ['010', 'swellendam@pnastrandgroup.co.za']
BranchList = {"Strand":Strand, "Colours":Colours, "Swes":Swes, "Eikestad":Eikestad, \
              "Sanctuary":Sanctuary, "Hermanus":Hermanus, "Gordonsbay":Gordonsbay, "Swellendam":Swellendam}

##UI promt so user can choose relevant branches for main loop
Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Sanctuary", "Hermanus", "Gordonsbay", "Swellendam"])


#get stock info from user for email details
supFilter = pyautogui.prompt(text='Please supplier name', title='Supplier Name', default='')
DueDate = pyautogui.prompt(text='Please enter the Due Date', title='Due Date', default='')


    

##The rest is just the script for the UI Atomation to loop per branch chosen
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
    pyautogui.click(366, 34)        #utilitys
    time.sleep(2)
    pyautogui.click(30, 80)         #change company
    time.sleep(2)
    keyboard.write(BranchList[Branch][0], 0.25)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(5)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.click(250, 28)        #stock
    time.sleep(0.5)
    pyautogui.click(279, 77)         #stock enquiries
    time.sleep(3)
    pyautogui.click(272, 731)       #filter
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
    time.sleep(10)
    pyautogui.click(1221, 725)       #preview
    time.sleep(30)
    pyautogui.click(112, 34)        #email
    time.sleep(2)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write(BranchList[Branch][1], 0.25)
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write(supFilter+" Price Changes", 0.25)
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)    
    keyboard.write('Goeie Dag', 0.3)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.write('Sien asb aangeheg price changes.', 0.3)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.write('Due '+DueDate, 0.3)    
    time.sleep(5)
    pyautogui.click(333, 667)
    time.sleep(3)
    keyboard.send('enter')
    time.sleep(5)
    keyboard.send('esc')
    time.sleep(2)
    keyboard.send('esc')
    time.sleep(2)
    pyautogui.click(1317, 10)           #minimize IQ
    time.sleep(2)
    pyautogui.click(rdpmin)            #minimize RDP
    time.sleep(2)
##    print(Branch[2])
    continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
    if continueScript == 'No':
        quit()
    else:
        continue
    

