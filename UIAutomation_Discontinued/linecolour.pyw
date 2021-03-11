##This UI automation was used to change stock 'linecolours'
##via RDP on our IQ Retail system.
##This is no longer in use and has since been replaced with SQL

import pyautogui
import keyboard
import time
import os
from easygui import *
import pandas as pd

##this was used on 2 PCs, rdp min location was at diff positions
if os.environ['computername'] == 'USER-PC':
    rdpmin = (910, 12)
else:
    rdpmin = (882, 12)

##get file from user
file = pyautogui.prompt(text='Import file name?', title='import', default='')

file = pd.read_excel(rf"C:\Users\User\Desktop\{file}.xlsx", header=None, usecols=[1, 4, 5])



##The rest is just the UI automtion script taking over
wait = 10

df = []
for code, br, ac, in zip(file[0], file[1], file[2]):
    df.append([code, br, ac])
print(str(df[0][0])[1])
if str(df[0][0]) == 'nan':
    del(df[0])
branches = []
for x in df:
    if x[1] not in branches:
        branches.append(x[1])
#print(branches)
for branch in branches:
    print(int(branch))

    newdf = []
    for line in df:
        if line[1] == branch:
            newdf.append([line[0], line[2]])
    newdf = pd.DataFrame(newdf, columns=['code','linecolourtype'])
    filename = str(int(branch))
    print(filename)
    if len(filename) == 1:
        filename = "00"+filename
    elif len(filename) == 2:
        filename = "0"+filename
    newdf.to_csv(rf"C:\Users\User\Desktop\{filename}.csv", index=False)

comps = {1:'001', 2:"002", 3:"003", 4:"004", 5:"005", 6:"006", 7:"007", 9:"009", 10:"010"}

pyautogui.confirm(text='please copy file to server', title='Continue', buttons=['Yes',])

for branch in branches:
    keyboard.send('windows')
    time.sleep(2)
    pyautogui.click(321, 748)       #rdp tasbar
    time.sleep(2)
    keyboard.send('windows')
    time.sleep(2)
    pyautogui.click(191, 755)       #iq in rdp taskbar
    time.sleep(2)
    #Branches = Branch
    time.sleep(2)
    pyautogui.click(366, 34)        #utilitys
    pyautogui.click(30, 80)         #change company
    time.sleep(2)
    keyboard.write(comps[branch], 0.25)
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
    keyboard.write(comps[branch]+".csv", 0.25)
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
    pyautogui.click(rdpmin)            #minimize RDP
    time.sleep(2)
##    print(Branch[2])
    continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
    if continueScript == 'No':
        quit()
    else:
        continue
