##This UI Automation was used to check for old Transfers
##and email list to branches that need to follow up.
##This is no longer in use and has since been
##replaced by SQL and Fast Repots

import pyautogui
import os
import time
import keyboard
import datetime
from datetime import date, timedelta
from easygui import *

daten = (datetime.date.today() - timedelta(3)).strftime('%Y-%m-%d')
date7 = (datetime.date.today() - timedelta(10)).strftime('%Y-%m-%d')



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

Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Sanctuary", "Hermanus", "Gordonsbay", "Swellendam"])






for Branch in Branches:

    myFilter = "(IN_DOCUMENT = NULL) AND (IN_BRANCH LIKE '"+BranchList[Branch][0]+"') \
    AND (OUT_DATE < '"+daten+"') \
    AND (OUT_TOTAL > 0) \
    AND (OUT_BRANCH NOT LIKE '888') \
    AND ((OUT_BRANCH NOT LIKE '007') OR (OUT_DATE < '"+date7+"')) \
    AND ((OUT_BRANCH NOT LIKE '010') OR (OUT_DATE < '"+date7+"')) \
    AND NOT ORDERNUMBER LIKE '%XMAS%'"
    
    if BranchList[Branch][0] == '007' or BranchList[Branch][0] == '010':
        myFilter = "(IN_DOCUMENT = NULL) AND (IN_BRANCH LIKE '"+BranchList[Branch][0]+"') \
        AND (OUT_DATE < '"+date7+"') \
        AND (OUT_TOTAL > 0) \
        AND (OUT_BRANCH NOT LIKE '888') \
        AND ((OUT_BRANCH NOT LIKE '007') OR (OUT_DATE < '"+date7+"')) \
        AND ((OUT_BRANCH NOT LIKE '010') OR (OUT_DATE < '"+date7+"')) \
        AND NOT ORDERNUMBER LIKE '%XMAS%'"

    
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
    pyautogui.click(253, 31)        #stock
    time.sleep(2)
    pyautogui.click(697, 71)        #utilities
    time.sleep(2)
    pyautogui.click(741, 214)       #transfers
    time.sleep(2)
    pyautogui.click(305, 697)       #filter
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
    time.sleep(1)
    pyautogui.click(882, 12)            #minimize RDP
    time.sleep(2)
    ask = pyautogui.confirm(text='Email or skip to next branch?', \
                            title='EMAIL OR SKIP', buttons=['PRINT', 'NEXT'])
    if ask == 'NEXT':
        keyboard.send('windows')
        time.sleep(2)
        pyautogui.click(321, 748)       #rdp tasbar
        time.sleep(2)
        keyboard.send('esc')
        time.sleep(2)
        pyautogui.click(1317, 10)           #minimize IQ
        time.sleep(2)
        pyautogui.click(882, 12)            #minimize RDP
        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
            continue
    else:
        keyboard.send('windows')
        time.sleep(2)
        pyautogui.click(321, 748)       #rdp tasbar
        time.sleep(2)
        pyautogui.click(476, 162)       #summary
        time.sleep(1)
        pyautogui.click(1113, 699)       #report
        time.sleep(1)
        pyautogui.click(1085, 642)       #previev
        time.sleep(5)
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
        keyboard.write("liana@pnastrandgroup.co.za", 0.25)
        time.sleep(2)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.send('tab')
        time.sleep(0.5)
        keyboard.send('tab')
        time.sleep(0.5)
        keyboard.write("OUTSTANDING TRANSFERS OLDER THAN 4/10 DAYS", 0.25)
        time.sleep(0.5)
        keyboard.send('tab')
        time.sleep(0.5)    
        keyboard.write('Goeie Dag', 0.3)
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.write('Sien asb transfers wat ouer is as 4/10 dae. Volg asb die transfers op so gou moontlik.', 0.3)
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.write('Groete', 0.3)    
        time.sleep(5)
        pyautogui.click(333, 667)
        time.sleep(3)
        keyboard.send('enter')
        time.sleep(5)
        pyautogui.click(17, 38)
        time.sleep(1)
        keyboard.send('enter')
        time.sleep(7)
        keyboard.send('esc')
        time.sleep(2)
        keyboard.send('esc')
        time.sleep(2)
        pyautogui.click(1317, 10)           #minimize IQ
        time.sleep(2)
        pyautogui.click(882, 12)            #minimize RDP
        time.sleep(2)

        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
            continue
    

##(IN_DOCUMENT = NULL) AND (IN_BRANCH LIKE '006')
##    AND (OUT_DATE < '2019-04-08') AND (OUT_TOTAL > 0) AND
##    (OUT_BRANCH NOT LIKE '888') AND ((OUT_BRANCH NOT LIKE '007') OR (OUT_DATE < '2019-04-01'))
##    AND NOT ORDERNUMBER LIKE '%XMAS%'
