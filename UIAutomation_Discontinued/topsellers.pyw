##This UI Automation was used to get and send a report
##to my manager in our IQ Retail System
##This is no longer in use and has since
##been replaced with SQL and Fast Reports

import pyautogui
import os
import time
import keyboard
import datetime
from datetime import date, timedelta
from easygui import *

##this was used on 2 PCs, rdp min location was at diff positions
if os.environ['computername'] == 'USER-PC':
    rdpmin = (910, 12)
else:
    rdpmin = (882, 12)

##setup branch details
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

##request what branches to loop through from user
Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Sanctuary", "Hermanus", "Gordonsbay", "Swellendam"])



##The rest is just the UI Automation scripts main loop
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
    pyautogui.click(30, 80)         #change company
    time.sleep(2)
    keyboard.write(BranchList[Branch][0], 0.25)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(5)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.click(366, 34)        #utilitys
    pyautogui.click(567, 80)        #reports
    time.sleep(2)
    keyboard.write('stock_S', 0.15)
    time.sleep(2)
    keyboard.press('alt')
    time.sleep(0.5)
    keyboard.press('p')
    time.sleep(2)
    keyboard.release('p')
    keyboard.release('alt')
    time.sleep(5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.write('02', 0.15)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('f9')
    time.sleep(0.5)
    keyboard.write('0279', 0.15)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write((datetime.date.today() - timedelta(30)).strftime('%d%m%y'), 0.25)
    keyboard.send('tab')
    keyboard.write((datetime.date.today() - timedelta(1)).strftime('%d%m%y'), 0.25)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write('q', 0.15)
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write('o', 0.15)
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(2)
    pyautogui.click(112, 34)
    time.sleep(2)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write('GARY@pnastrandgroup.co.za', 0.15)
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.write('chris@pnastrandgroup.co.za', 0.15)
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
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
    keyboard.write(BranchList[Branch][0]+"  TOP VERKOPERS AFGELOPE 30 DAE", 0.25)
    time.sleep(0.5)
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

#National book sale email

keyboard.send('windows')
time.sleep(2)
pyautogui.click(321, 748)       #rdp tasbar
time.sleep(2)
keyboard.send('windows')
time.sleep(2)
pyautogui.click(191, 755)       #iq in rdp taskbar
time.sleep(2)
pyautogui.click(366, 34)        #utilitys
pyautogui.click(30, 80)         #change company
time.sleep(2)
keyboard.write('000', 0.25)
time.sleep(0.5)
keyboard.send('enter')
time.sleep(5)
keyboard.send('enter')
time.sleep(2)
pyautogui.click(366, 34)        #utilitys
pyautogui.click(567, 80)        #reports
time.sleep(2)
keyboard.write('nat', 0.15)
time.sleep(2)
keyboard.press('alt')
time.sleep(0.5)
keyboard.press('p')
time.sleep(2)
keyboard.release('p')
keyboard.release('alt')
time.sleep(20)
pyautogui.click(112, 34)
time.sleep(2)
keyboard.send('tab')
time.sleep(0.5)
keyboard.write('GARY@pnastrandgroup.co.za', 0.15)
time.sleep(2)
keyboard.send('enter')
time.sleep(0.5)
keyboard.send('enter')
time.sleep(0.5)
keyboard.write('chris@pnastrandgroup.co.za', 0.15)
time.sleep(2)
keyboard.send('enter')
time.sleep(0.5)
keyboard.send('enter')
time.sleep(0.5)
keyboard.write('liana@pnastrandgroup.co.za', 0.15)
time.sleep(2)
keyboard.send('enter')
time.sleep(0.5)
keyboard.send('enter')
time.sleep(0.5)
keyboard.write('it@pnastrandgroup.co.za', 0.15)
time.sleep(2)
keyboard.send('enter')
time.sleep(0.5)
keyboard.send('enter')
time.sleep(0.5)
for branch in BranchList.values():
    keyboard.write(branch[1], 0.25)
    time.sleep(2)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
keyboard.send('tab')
time.sleep(0.5)
keyboard.send('tab')
time.sleep(0.5)
keyboard.write("NATIONAL TOP 100 BOOK SALES LAST 30 DAYS", 0.25)
time.sleep(0.5)
pyautogui.click(333, 667)
time.sleep(10)
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
