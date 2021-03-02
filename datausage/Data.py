##when we used limited data, this was a test to see if more data will be needed before the month ends

import pandas as pd
import wget
import datetime
from datetime import date, timedelta
from easygui import *
import pyautogui
import csv
import os

##get amount topped up
topup = pyautogui.prompt(text='Was there a topup? please enter amount', title='Topup', default='')


##get data from url
daten = (datetime.date.today() - timedelta(0)).strftime('%Y-%m-%d')
url = 'https://www.youradsl.co.za/calls/internalaction_usagedownload.php?username=pnas001@itblue.co.za&year=2018&month=8'
filename = wget.download(url)
filename
da = pd.read_csv('pnas001@itblue.co.za_usage_' + daten + '.csv')
count = 0
total = 0
avg = 0
est = 0
for usage in zip(da['TotalUsage_GB']):
    usage1 = usage[0]
    count += 1
    total += float(usage1)

avg = total/(count-1)
est = (30 - (count-1))*avg
db = pd.read_csv('cap.csv')
cap = db['cap']
balance = int(cap[0]) + int(topup) - total

print ("Days past: " + str(count-1))
print ("Monthly total: " + str(total))
print ("Avg per Day: " + str(avg))
print ("Estemate needed: " + str(est))

result = "Days past: " + str(count-1) + '\n' \
          "Monthly total: " + str(total) + '\n' \
          "Avg per Day: " + str(avg) + '\n' \
          "Estemate needed: " + str(est) + '\n' \
          "Balance: " + str(balance)

pyautogui.confirm(text= result, title='Resuts')

with open('cap.csv', 'w', newline='') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=' ')
    spamwriter.writerow(['cap'])
    spamwriter.writerow([cap[0]+int(topup)])

os.remove('pnas001@itblue.co.za_usage_' + daten + '.csv')

quit()

    
