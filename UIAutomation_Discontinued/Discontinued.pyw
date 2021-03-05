##this was used fir UI automation on server through RDP connection,
##this has since been replaced by SQL scripts and FastReports
##purpose was to get data from IQ Retail of dicontinued stock and email to branch

import pyautogui
import keyboard
import time
import os
#import clipboard
from easygui import *

##setup two filters depending on branch size
filterBig = "((ONHAND <> 0) AND (LINECOLOURTYPE IN (6, 7, 8, 9))) OR ((DESCRIPT LIKE 'ZZ %') AND (ONHAND <> 0))"
filterSmall = "((LINECOLOURTYPE IN (5, 6, 7, 8, 9)) AND (ONHAND <> 0)) OR ((DESCRIPT LIKE 'ZZ %') AND (ONHAND <> 0))"

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

##UI for user to choose branches to run through
Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Sanctuary", "Hermanus", "Gordonsbay", "Swellendam"])





##UI automation loop for branches chosen
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
    if BranchList[Branch][0] == '002' or BranchList[Branch][0] == '004' or BranchList[Branch][0] == '006' or BranchList[Branch][0] == '007':
        keyboard.write(filterBig, 0.3)
    else:
        keyboard.write(filterSmall, 0.3)
    time.sleep(1)
    keyboard.send('tab')
    time.sleep(1)
    keyboard.write('a', 0.25)
    time.sleep(10)
    pyautogui.click(1221, 725)       #preview
    time.sleep(10)
    0, 20
    while True:
        img = pyautogui.locateOnScreen('reporticons.png', region=(0, 20, 250, 150))
        time.sleep(2)
        if img != None:
            break
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
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write("Discontinued Stok Onhand", 0.25)
    time.sleep(0.5)
    keyboard.send('tab')
    time.sleep(0.5)
    keyboard.write('Goeie Dag', 0.3)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.write('Sien asb aangeheg lys van discontinued stock onhand.', 0.3)
    time.sleep(0.5)
    keyboard.send('enter')
    time.sleep(0.5)
    keyboard.write('Maak asb seker dat wanneer hierdie items uitverkoop is, dat julle die pryse van die rak afhaal en nuwe produkte op sit.', 0.3)
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
##    pyautogui.click(17, 38)
##    time.sleep(1)
##    keyboard.send('enter')
##    time.sleep(12)
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
