##This UI Automation script was used to get old sales orders
##that have not been converted to invoices an mail the list
##to the branch to follow up on the stock or the customer.
##This is no longer in use and has since been replaced
##by SQL and Fast Reports

import pyautogui
import os
import time
import keyboard
import datetime
from datetime import date, timedelta
from easygui import *
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mimetypes
import email.mime.application
import smtplib
import win32com.client as win32


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


##myFilter = pyautogui.prompt(text='Enter Filter Formula:', title='Filter', default='')
fdate = datetime.datetime.now().strftime('%Y-%m-%d')
if int(fdate[5:7]) < 10:
    myFilter = "(ORDDATE < '"+fdate[:5]+'0'+str(int(fdate[5:7])-1)+'-01'+"')"
else:
    myFilter = "(ORDDATE < '"+fdate[:5]+str(int(fdate[5:7])-1)+'-01'+"')"
#myFilter = "(ORDDATE < '2019-04-01')"

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
    from_email='it@pnastrandgroup.co.za'
    from_password='DDA.xhp.136'
    to_email=address
    subject='OLD SALES ORDERS'

    message="""\
    <html>
        <head></head>
        <body>
            <p>Goeie Dag</p>
            <p>Volg op of delete asb die sales orders aangeheg.</p>
            <p>Groete</p>
            <image src='https://i.imgur.com/iRJ0Pe8.png'></image>
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

    filename = BranchList[Branch][0]+".xlsx"
    fo = open("C:\\Users\\User\\Desktop\\"+BranchList[Branch][0]+".xlsx", 'rb')
    file = email.mime.application.MIMEApplication(fo.read(),_subtype="xlsx")
    fo.close()
    file.add_header('Content-Disposition','attachment',filename=filename)
    msg.attach(file)

    emacc=smtplib.SMTP('smtp-mail.outlook.com', 587)
    emacc.ehlo()
    emacc.starttls()
    emacc.login(from_email, from_password)
    emacc.send_message(msg)



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
    while True:
        img = pyautogui.locateOnScreen('Processing.png', region=(0, 0, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    while True:
        img = pyautogui.locateOnScreen('SalesOrders.png', region=(220, 30, 250, 150))
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
    pyautogui.click(882, 12)        #minimize RDP
    continueScript = pyautogui.confirm(text='Do you want to email, continue to next branch or stop?', title='Continue or Stop', buttons=['Next Branch', 'Email', 'Stop'])
    if continueScript == 'Stop':
        quit()
    elif continueScript == 'Next Branch':
        time.sleep(1)
        keyboard.send('windows')
        time.sleep(2)
        pyautogui.click(321, 748)       #rdp tasbar
        time.sleep(1)
        keyboard.send('esc')
        time.sleep(1)
        keyboard.send('enter')
        time.sleep(1)
        pyautogui.click(1317, 10)       #minimize IQ
        time.sleep(2)
        pyautogui.click(882, 12)        #minimize RDP
        time.sleep(2)
##        Branch = eval(Branch[2])
        continue
    elif continueScript == 'Email':
        time.sleep(1)
        keyboard.send('windows')
        time.sleep(2)
        pyautogui.click(321, 748)       #rdp tasbar
        time.sleep(1)
        pyautogui.click(177, 728)       #Export
        time.sleep(1)
        pyautogui.click(175, 650)       #Export XLS
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
        keyboard.write(BranchList[Branch][0], 0.25)
        time.sleep(0.5)
        keyboard.send('enter')
        x, y = pyautogui.center(img2)
        pyautogui.rightClick(x, y)
        while True:
            img = pyautogui.locateOnScreen('Copy.png', region=(45, 550, 250, 150))
            if img != None:
                x, y = pyautogui.center(img)
                pyautogui.click(x, y)
                break
        while True:
            img = pyautogui.locateOnScreen('MinimizeRDP.png', region=(845, 0, 250, 150))
            if img != None:
                x, y = pyautogui.center(img)
                pyautogui.click(x, y)
                break
        time.sleep(2)
        pyautogui.click(882, 12)        #minimize RDP
        time.sleep(2)
        pyautogui.click(882, 12)        #select desktop
        time.sleep(2)
        keyboard.press('ctrl+v')
        time.sleep(2)
        keyboard.release('ctrl+v')
        time.sleep(7)
        convertxls(BranchList[Branch][0])
        # os.startfile('C:\\Users\\User\\Desktop\\'+BranchList[Branch][0]+'.xls')
        # time.sleep(5)
        # keyboard.send('enter')
        # time.sleep(1)
        # pyautogui.click(15, 212)        #select all
        # pyautogui.click(15, 212)        #select all
        # time.sleep(1)
        # pyautogui.click(1097, 124)      #format?
        # time.sleep(1)
        # pyautogui.click(1097, 246)      #autofit colomn
        # time.sleep(1)
        # time.sleep(3)
        # pyautogui.click(1340, 16)
        # time.sleep(3)
        # keyboard.send('enter')
        # time.sleep(2)
        # keyboard.send('enter')
        time.sleep(2)
        send_email(BranchList[Branch][1])
        # keyboard.send('windows')
        # time.sleep(1)
        # pyautogui.click(223, 749)
        # while True:
        #     img = pyautogui.locateOnScreen('ThunderBirdNew.png', region=(90, 50, 250, 150))
        #     if img != None:
        #         x, y = pyautogui.center(img)
        #         pyautogui.click(x, y)
        #         break
        # time.sleep(2)
        # keyboard.write(BranchList[Branch][1], 0.25)
        # keyboard.send('tab')
        # keyboard.write('OLD SALES ORDERS ', 0.25)
        # ##keyboard.write((datetime.date.today() - timedelta(1)).strftime('%d-%m-%y'), 0.25)
        # keyboard.send('tab')
        # keyboard.write('Goeie Dag', 0.25)
        # keyboard.send('enter')
        # keyboard.send('enter')
        # keyboard.write('Volg op of delete asb die sales orders aangeheg.', 0.25)
        # keyboard.send('enter')
        # keyboard.send('enter')
        # keyboard.send('ctrl+shift+a')
        # time.sleep(1)
        # time.sleep(1)
        # keyboard.write(BranchList[Branch][0]+'.xlsx', 0.25)
        # keyboard.send('enter')
        # time.sleep(1)
        # keyboard.send('tab')
        # time.sleep(1)
        # keyboard.send('ctrl+enter')
        # time.sleep(1)
        # keyboard.send('enter')
        # time.sleep(3)
        # pyautogui.doubleClick(1250, 12)       #minimize email
        time.sleep(2)
##        print(Branch[2])
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
            pyautogui.click(78, 635)
            ##keyboard.send('delete')
            time.sleep(1)
            keyboard.send('enter')
            time.sleep(2)
            pyautogui.click(882, 12)
            continue
