##This UI Automation was used to check price overrides at each
##branch for the previouse day against 3 days ago
##to see if they have updated their stock price labels as needed
##This is no longer in use and has since been replced
##by SQL and Fast Repots

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

# def id():
#  return os.getpid()
#id = os.getpid()

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

def send_email(address):
    from_email= # removed for security purposes
    from_password= # removed for security purposes
    to_email=address
    subject='PRICE OVERRIDES ' + (datetime.date.today() - timedelta(1)).strftime('%d-%m-%y')

    message="""\
    <html>
        <head></head>
        <body>
            <p>Goeie Dag</p>
            <p>Sien asb aangeheg overrides wat opgevolg moet word.</p>
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

    filename = "PRICE OVERRIDES.pdf"
    fo = open(r"C:\Users\User\Desktop\PRICE OVERRIDES FWR.pdf", 'rb')
    file = email.mime.application.MIMEApplication(fo.read(),_subtype="pdf")
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
        img = pyautogui.locateOnScreen('Utilities.png', region=(330, 10, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    while True:
        img = pyautogui.locateOnScreen('UserReports.png', region=(530, 45, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    time.sleep(2.5)
    keyboard.write('price overrides fwr', 0.25)
    time.sleep(2)
    while True:
        img = pyautogui.locateOnScreen('POV2.png', region=(600, 195, 250, 150))
        if img != None:
                time.sleep(1)
                while True:
                    img = pyautogui.locateOnScreen('Preview.png', region=(1060, 660, 250, 150))
                    if img != None:
                        x, y = pyautogui.center(img)
                        pyautogui.click(x, y)
                        break
        break
    time.sleep(5)
    #keyboard.write((datetime.date.today() - timedelta(1)).strftime('%d%m%y'), 0.25)
    keyboard.send('tab')
    #time.sleep(1)
    #keyboard.write((datetime.date.today() - timedelta(1)).strftime('%d%m%y'), 0.25)
    keyboard.send('tab')
    keyboard.send('enter')
##    while True:
##        img = pyautogui.locateOnScreen('ReportIcons.png', region=(0, 15, 250, 150))
##        if img != None:
##            break
##    while True:
##        img = pyautogui.locateOnScreen('ReportSave.png', region=(45, 15, 250, 150))
##        if img != None:
##            x, y = pyautogui.center(img)
##            pyautogui.click(x, y)
##            break
    while True:
        img = pyautogui.locateOnScreen('ReportSavePDF.png', region=(66, 15, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
##    while True:
##        img = pyautogui.locateOnScreen('WYSIWYG.png', region=(670, 350, 250, 150))
##        img2 = pyautogui.locateOnScreen('notWYSIWYG.png', region=(670, 350, 250, 150))
##        if img != None:
##            x, y = pyautogui.center(img)
##            pyautogui.click(x, y)
##            break
##        elif img2 != None:
##            break
    time.sleep(2)
    keyboard.send('enter')
    while True:
        img = pyautogui.locateOnScreen('DesktopLocation.png', region=(410, 225, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.click(x, y)
            break
    keyboard.send('enter')
    time.sleep(2)
    keyboard.send('esc')
    time.sleep(2)
    keyboard.send('esc')
    time.sleep(2)
    pyautogui.click(1317, 10)       #minimize IQ
    time.sleep(2)
    while True:
        img = pyautogui.locateOnScreen('PriceOverrideIcon.png', region=(0, 650, 250, 150))
        if img != None:
            x, y = pyautogui.center(img)
            pyautogui.rightClick(x, y)
            break
    while True:
        img = pyautogui.locateOnScreen('Copy.png', region=(50, 560, 250, 150))
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
    if os.path.isfile("C:\\Users\\User\\Desktop\\PRICE OVERRIDES FWR.pdf"):
        os.remove("C:\\Users\\User\\Desktop\\PRICE OVERRIDES FWR.pdf")
    time.sleep(1)
    keyboard.press('ctrl+v')
    time.sleep(2)
    keyboard.release('ctrl+v')
    time.sleep(7)
    os.startfile('C:\\Users\\User\\Desktop\\PRICE OVERRIDES FWR.pdf')
    time.sleep(10)
    ask = pyautogui.confirm(text='Please review the list, How many pages must I print?', title='PRINT', buttons=['0', '1', '2'])
    time.sleep(3)
    if ask == "0":
        time.sleep(1)
        pyautogui.click(1340, 16)
        time.sleep(3)
##        keyboard.send('enter')
##        time.sleep(1)
##        print(Branch[2])
        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
##            Branch = eval(Branch[2])
            continue
    else:
        if ask =="2":
            keyboard.send('ctrl+s')
            time.sleep(3)
            keyboard.send('enter')
            time.sleep(3)
            keyboard.send('y')
            time.sleep(3)
        time.sleep(1)
        keyboard.send('ctrl+p')
##        while True:
##            img = pyautogui.locateOnScreen('Print.png', region=(155, 120, 250, 150))
##            if img != None:
##                x, y = pyautogui.center(img)
##                break
        time.sleep(2)
        keyboard.send('tab')
        keyboard.send(ask)
        time.sleep(1)
        keyboard.send('enter')
        time.sleep(6)
        pyautogui.click(1340, 16)
        time.sleep(3)
        send_email(BranchList[Branch][1])



##        keyboard.send('enter')
##        time.sleep(2)
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
        # keyboard.write('PRICE OVERRIDES ', 0.25)
        # keyboard.write((datetime.date.today() - timedelta(1)).strftime('%d-%m-%y'), 0.25)
        # keyboard.send('tab')
        # keyboard.write('Goeie Dag', 0.25)
        # keyboard.send('enter')
        # keyboard.send('enter')
        # keyboard.write('Sien asb aangeheg overrides wat opgevolg moet word.', 0.25)
        # keyboard.send('enter')
        # keyboard.send('enter')
        # keyboard.send('ctrl+shift+a')
        # time.sleep(5)
        # # pyautogui.click(76,217)
        # # time.sleep(2)
        # # keyboard.send('tab')
        # # keyboard.send('tab')
        # # keyboard.send('tab')
        # # time.sleep(1)
        # # keyboard.write('PRICE OVERRIDES FWR.pdf', 0.25)
        # keyboard.write(r"C:\Users\User\Desktop\PRICE OVERRIDES FWR.pdf", 0.25)
        # time.sleep(1)
        # keyboard.send('enter')
        # time.sleep(2)
        # keyboard.send('tab')
        # time.sleep(1)
        # keyboard.send('ctrl+enter')
        # time.sleep(2)
        # keyboard.send('enter')
        # time.sleep(3)
        # pyautogui.doubleClick(1250, 12)       #minimize email
        time.sleep(2)
##        print(Branch[2])
        continueScript = pyautogui.confirm(text='Do you want to continue to next branch or stop?', title='Continue or Stop', buttons=['Yes', 'No'])
        if continueScript == 'No':
            quit()
        else:
            continue
