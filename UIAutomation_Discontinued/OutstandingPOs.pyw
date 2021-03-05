##This UI Automation was used to get a list of outstanding
##purchase orders and email it the the persons responsible for
##stock ordering
##This is no longer in use and has since been
##replaced by SQL and Fast Repots


import pyautogui
import os
import time
import keyboard
import datetime
from datetime import date, timedelta
from easygui import *
import win32com.client as win32
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mimetypes
import email.mime.application
import smtplib

if os.environ['computername'] == 'USER-PC':
    rdpmin = (910, 12)
else:
    rdpmin = (882, 12)

def convertxls(filename):
    fname = "C:\\Users\\User\\Desktop\\"+filename+".xls"
    # excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()


def send_email(address):
    from_email= # removed for security purposes
    from_password= # removed for security purposes
    to_email=address
    subject='OUTSTANDING PURCHASE ORDERS'

    message="""\
    <html>
        <head></head>
        <body>
            <p>Goeie Dag</p>
            <p>Sien asb aangeheg ou purchase orders wat opgevolg moet word.</p>
            <p>Groete</p>
            <image src='https://i.imgur.com/EiRbB2H.jpg'></image>
            </body>
    </html>
    """

    msg = MIMEMultipart()
    txt=MIMEText(message, 'html')
    msg["Subject"]=subject
    msg["To"]=to_email
    msg["From"]=from_email
    msg["Cc"]=from_email
    msg.attach(txt)


    for Branch in Branches:
        if os.path.isfile(rf"C:\Users\User\Desktop\{BranchList[Branch][0]}.xlsx"):
            filename = f"{BranchList[Branch][0]}.xlsx"
            fo = open(rf"C:\Users\User\Desktop\{BranchList[Branch][0]}.xlsx", 'rb')
            file = email.mime.application.MIMEApplication(fo.read(),_subtype="xlsx")
            fo.close()
            file.add_header('Content-Disposition','attachment',filename=filename)
            msg.attach(file)

    emacc=smtplib.SMTP('smtp-mail.outlook.com', 587)
    emacc.ehlo()
    emacc.starttls()
    emacc.login(from_email, from_password)
    emacc.send_message(msg)

Strand = ['001', 'strand@pnastrand.co.za']
Colours = ['002', 'colours@pnastrand.co.za']
Swes = ['003', 'swes@pnastrand.co.za']
Eikestad = ['004', 'eikestad@pnastrand.co.za']
Commercial = ['005', 'sales@pnastrand.co.za']
Sanctuary = ['006', 'sanctuary@pnastrand.co.za']
Hermanus = ['007', 'whalecoast@pnastrand.co.za']
Gordonsbay = ['009', 'gordonsbay@pnastrand.co.za']
Warehouse = ['888', 'warehouse@pnastrand.co.za']
BranchList = {"Strand":Strand, "Colours":Colours, "Swes":Swes, "Eikestad":Eikestad, \
              "Commercial":Commercial, "Sanctuary":Sanctuary, "Hermanus":Hermanus, "Gordonsbay":Gordonsbay,\
              "Warehouse":Warehouse}

Branches = multchoicebox("choose", title= "hello", choices= \
                       ["Strand", "Colours", "Swes", "Eikestad", \
                        "Commercial", "Sanctuary", "Hermanus", \
                        "Gordonsbay", "Warehouse"])


##myFilter = pyautogui.prompt(text='Enter Filter Formula:', title='Filter', default='')
filterDate= (datetime.date.today() - timedelta(30)).strftime('%Y-%m-%d')
myFilter = "(ORDDATE < '"+filterDate+"') AND (ORDERNUM NOT LIKE '%SUB%') AND (INVDATE = '1899-12-30')"


# send_email('it@pnastrandgroup.co.za')


for Branch in Branches:
##    Branches = Branch
    time.sleep(1)
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
    while True:
        img = pyautogui.locateOnScreen('Processing.png', region=(0, 0, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    while True:
        img = pyautogui.locateOnScreen('PurchaseOrders.png', region=(270, 30, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    time.sleep(3)
    pyautogui.click(295, 734)       #filter
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
    pyautogui.click(177, 728)       #Export
    time.sleep(1)
    pyautogui.click(175, 671)       #Export XLS
    time.sleep(1)


    nsec = 0
    while nsec < 3:
        img = pyautogui.locateOnScreen('noexport.png', region=(610, 344, 250, 250))
        time.sleep(3)
        if img != None:
            keyboard.send('enter')
            time.sleep(1)
            keyboard.send('esc')
            time.sleep(1)
            keyboard.send('enter')
            time.sleep(1)
            pyautogui.click(1317, 10)       #minimize IQ
            time.sleep(2)
            pyautogui.click(rdpmin)        #minimize RDP
            time.sleep(2)
            continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
            if continueScript == 'No':
                quit()
            break
        time.sleep (1)
        nsec += 1
    if nsec > 2:
##        time.sleep(1)
##        keyboard.send('windows')
##        time.sleep(2)
##        pyautogui.click(321, 748)       #rdp tasbar
        time.sleep(1)

        keyboard.write('d')
        time.sleep(1)
        keyboard.write('d')
        time.sleep(1)
        if BranchList[Branch][0] == '005' or BranchList[Branch][0] == '888':
            keyboard.write('d')
            time.sleep(1)
            keyboard.write('d')
            time.sleep(1)
            keyboard.write('d')
            time.sleep(1)
            keyboard.write('d')
            time.sleep(1)
        pyautogui.click(564, 296)       #desktop
        time.sleep(1)
        time.sleep(1)
        keyboard.send('enter')
        time.sleep(1)
        pyautogui.click(915, 593)       #Accept
        time.sleep(3)
        keyboard.send('right')
        time.sleep(1)
        keyboard.send('enter')
        time.sleep(1)
        keyboard.send('esc')
        time.sleep(1)
        keyboard.send('enter')
        time.sleep(1)
        pyautogui.click(1317, 10)       #minimize IQ
        time.sleep(2)
        while True:
            img2 = pyautogui.locateOnScreen('File.png', region=(0, 630, 250, 150))
            if img2 != None:
                x, y = pyautogui.center(img2)
                pyautogui.rightClick(x, y)
                break
        while True:
            img = pyautogui.locateOnScreen('Rename.png', region=(50, 560, 250, 150))
            if img != None:
                x, y = pyautogui.center(img)
                pyautogui.click(x, y)
                break
        keyboard.write(BranchList[Branch][0], 0.4)
        time.sleep(1)
        keyboard.send('enter')
        x, y = pyautogui.center(img2)
        pyautogui.rightClick(x, y)
        time.sleep(1)
        while True:
            img = pyautogui.locateOnScreen('Copy.png', region=(45, 550, 250, 150))
            if img != None:
                x, y = pyautogui.center(img)
                pyautogui.click(x, y)
                break
        time.sleep(1)
        while True:
            img = pyautogui.locateOnScreen('MinimizeRDP.png', region=(845, 0, 250, 150))
            if img != None:
                x, y = pyautogui.center(img)
                pyautogui.click(x, y)
                break
        time.sleep(2)
        pyautogui.click(882, 12)        #minimize RDP
        time.sleep(2)
        pyautogui.click(rdpmin)        #select desktop
        time.sleep(2)
        keyboard.press('ctrl+v')
        time.sleep(2)
        keyboard.release('ctrl+v')
        time.sleep(10)
        convertxls(BranchList[Branch][0])
        time.sleep(5)

        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
            keyboard.send('windows')
            time.sleep(2)
            pyautogui.click(321, 748)       #rdp tasbar
            time.sleep(2)
            pyautogui.click(321, 748)       #rdp tasbar
            time.sleep(2)
            pyautogui.rightClick(35, 687)
            time.sleep(1)
            pyautogui.click(78, 635)        #delete?
            ##keyboard.send('delete')
            time.sleep(1)
            keyboard.send('enter')
            time.sleep(2)
            pyautogui.click(rdpmin)
            nsec = 0
            continue

send_email('orders@pnastrandgroup.co.za, corrine@pnastrandgroup.co.za')

#
# keyboard.send('windows')
# time.sleep(1)
# pyautogui.click(223, 749)
# while True:
#     img = pyautogui.locateOnScreen('ThunderBirdNew.png', region=(90, 50, 250, 150))
#     if img != None:
#         x, y = pyautogui.center(img)
#         pyautogui.click(x, y)
#         break
# time.sleep(5)
# keyboard.write('orders@pnastrandgroup.co.za', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# ##keyboard.write('gary@pnastrandgroup.co.za', 0.25)
# ##keyboard.send('enter')
# ##time.sleep(1)
# keyboard.write('art2@pnastrandgroup.co.za', 0.25)
# time.sleep(1)
# keyboard.send('tab')
#
# keyboard.write('OUTSTANDING PURCHASE ORDERS', 0.25)
# ##keyboard.write((datetime.date.today() - timedelta(1)).strftime('%d-%m-%y'), 0.25)
# keyboard.send('tab')
# keyboard.write('Goeie Dag', 0.25)
# keyboard.send('enter')
# keyboard.send('enter')
# keyboard.write('Sien asb aangeheg ou purchase orders wat opgevolg moet word.', 0.25)
# keyboard.send('enter')
# ##keyboard.write('Ek weet dit is vroeg, maar ek is volgende week op verlof.', 0.25)
# ##keyboard.send('enter')
# keyboard.send('enter')
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('001.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('002.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('003.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('004.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('005.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('006.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('007.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('009.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+shift+a')
# time.sleep(1)
# time.sleep(1)
# keyboard.write('888.xlsx', 0.25)
# keyboard.send('enter')
# time.sleep(1)
# keyboard.send('ctrl+enter')
# time.sleep(1)
# keyboard.send('enter')
# time.sleep(3)
# pyautogui.doubleClick(1250, 12)       #minimize email
# time.sleep(2)
