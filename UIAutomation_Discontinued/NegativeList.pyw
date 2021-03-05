##This UI was made go get a list of negative stock and email
##it to the branch to recount their stock and adjust as needed
##This is no longer in use and has been replaced by
##a combination of SQL and Fast Reports

import pyautogui
import keyboard
import time
import os
#import clipboard
from easygui import *
#(REGULAR_SU LIKE 'PYR001') AND (ONHAND <> 0) AND (MODIFIED BETWEEN '2018-06-13' AND '2018-06-14') AND (USER = 37)


Strand = ['001', 'strand@pnastrandgroup.co.za']
Colours = ['002', 'colours@pnastrandgroup.co.za']
Swes = ['003', 'swes@pnastrandgroup.co.za']
Eikestad = ['004', 'eikestad@pnastrandgroup.co.za']
Comm = ['005', 'sales@pnastrandgroup.co.za']
Sanctuary = ['006', 'sanctuary@pnastrandgroup.co.za']
Hermanus = ['007', 'whalecoast@pnastrandgroup.co.za']
Gordonsbay = ['009', 'gordonsbay@pnastrandgroup.co.za']
Swellendam = ['010', 'swellendam@pnastrandgroup.co.za']
BranchList = {"Strand":Strand, "Colours":Colours, "Swes":Swes, "Eikestad":Eikestad, \
              "Comm":Comm, "Sanctuary":Sanctuary, "Hermanus":Hermanus, "Gordonsbay":Gordonsbay, "Swellendam":Swellendam}

Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Comm", "Sanctuary", "Hermanus", "Gordonsbay", "Swellendam"])


myFilter = '(ONHAND < 0)'
##supFilter = pyautogui.prompt(text='Please supplier name', title='Supplier Name', default='')
##DueDate = pyautogui.prompt(text='Please enter the Due Date', title='Due Date', default='')





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
    time.sleep(10)
##    pyautogui.click(514, 31)
##    time.sleep(1)
##    keyboard.send('tab')
##    time.sleep(1)
##    keyboard.send('tab')
##    time.sleep(1)
##    keyboard.send('ctrl+x')
##    time.sleep(1)
##    pyautogui.hotkey('ctrl','x')
##    keyboard.press('ctrl+x')
##    time.sleep(1)
##    keyboard.release('ctrl+x')
##    time.sleep(1)
##    supFilter = int(clipboard.paste())
##    supFilter = pyautogui.prompt(text='How Many Pages?', title='Pages', default='')
    nsec = 0
    while True:
        img = pyautogui.locateOnScreen('1pages.png', region=(985, 85, 250, 150))
        time.sleep(5)
        if img != None:
            supFilter = 1
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
        time.sleep (1)
        nsec += 1
        if nsec > 1:
            supFilter = 2
            print ("timeout") # If the result display can not be obtained until 5 s or more have elapsed after the start of the processing operation, it ends with timeout
            break
    time.sleep(3)
    if int(supFilter) < 2:
        time.sleep(5)
        keyboard.send('esc')
        time.sleep(2)
        keyboard.send('esc')
        time.sleep(2)
        pyautogui.click(1317, 10)           #minimize IQ
        time.sleep(2)
        pyautogui.click(882, 12)            #minimize RDP
        time.sleep(2)
##        print(Branch[2])
        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
            continue
    else:
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
        keyboard.write("Negative List Greater than 1 Page", 0.25)
        time.sleep(0.5)
        keyboard.send('tab')
        time.sleep(0.5)
        keyboard.write('Goeie Dag', 0.3)
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.write('Sien asb aangeheg Negative lys.', 0.3)
        time.sleep(0.5)
        keyboard.send('enter')
        time.sleep(0.5)
        keyboard.write('Hierdie lys word aangestuur omdat dit meer as 1 bladsy bevat wat buite PNA regulasies val.', 0.3)
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
        time.sleep(12)
        keyboard.send('esc')
        time.sleep(2)
        keyboard.send('esc')
        time.sleep(2)
        pyautogui.click(1317, 10)           #minimize IQ
        time.sleep(2)
        pyautogui.click(882, 12)            #minimize RDP
        time.sleep(2)
##        print(Branch[2])
        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
            continue
