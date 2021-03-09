##this scripts check two boards on trello for overdue cards
##and then emails the details to the relevant branch as reminder

import requests
import json
import datetime
import smtplib
from email.mime.text import MIMEText

key1 = # removed for security purposes
token1 = # removed for security purposes

##dictionary of user IDs
Members = {"550bf4bc2094ac8b5df163dc":"strand@pnastrandgroup.co.za",
           "550aa531b66b3a41f9cbeb50":"colours@pnastrandgroup.co.za",
           "550fd08951e882750fa5a557":"swes@pnastrandgroup.co.za",
           "5df8704ca18b9f4d9a31c958":"eikestad@pnastrandgroup.co.za",
           "60068017690baa8596c41612":"stelsquare@pnastrandgroup.co.za",
           "5c3711bcd127c4444bd74167":"sales2@pnastrandgroup.co.za",
           "57f36c796b25693e34822da8":"sanctuary@pnastrandgroup.co.za",
           "5a1d98407eb93a0ddf5a0867":"whalecoast@pnastrandgroup.co.za",
           "5a283ea3be6c95d3fb2348ae":"gordonsbay@pnastrandgroup.co.za",
           "58da5eac5891d87ac1c2fe58":"warehouse@pnastrandgroup.co.za",
           "561f84f36d45d07ffc3b4dec":"sales@pnastrandgroup.co.za",
           "5c98854c0ca5cd64c112fd47":"swellendam@pnastrandgroup.co.za"}

##dictionary of list IDs
myList = {"To Do":"55042f9dd8eec487931646ea",
          "Doing":"55042fa1c1beafd66acdf60f",
          "Done":"58163ee99eb2ae47a2be5e85"}

##format due date from Trello to datetime
def dueDate(a):
    test = a
    y=int(test[:4])
    m=int(test[5:7])
    d=int(test[8:10])
    return datetime.date(y, m, d)

##function for email to be sent
def send_email(name, date, url, email):
    from_email=# removed for security purposes
    from_password=# removed for security purposes
    to_email=email
    subject='Outstanding Trello '+name
    CC = 'liana@pnastrandgroup.co.za'
    BCC = 'it@pnastrandgroup.co.za'


    message = """
        <html>
            <head></head>
            <body>
                <p>Goeie Dag</p>
                <p></p>
                <p>Volg die kaart op of trek na done:</p>
                <p>Card Description: %s <br>
                Was due on: %s <br>
                Link: <a href="%s">%s</a></p>
                <p>Groete</p>
                <p>Rick Robins</p>
                <image src='https://i.imgur.com/iRJ0Pe8.png'></image>
                </body>
        </html>
    """ %(name, date, url, url)
    msg=MIMEText(message, 'html')
    msg["Subject"]=subject
    msg["To"]=to_email
    msg["From"]=from_email
    msg["Cc"]=CC
    msg["Bcc"]=BCC

    emacc=smtplib.SMTP('smtp-mail.outlook.com', 587)
    emacc.ehlo()
    emacc.starttls()
    emacc.login(from_email, from_password)
    emacc.send_message(msg)


##loop though cards on 'To Do' list and email
response4 = requests.get("https://api.trello.com/1/lists/"+myList["To Do"]+"/cards?key="+key1+"&token="+token1)
data4 = response4.json()
for x in range(len(data4)):    
    if dueDate(data4[x]["due"]) < datetime.date.today():
        tName = data4[x]["name"]
        tDate = dueDate(data4[x]["due"])
        tURL = data4[x]["shortUrl"]
        tEmail = Members[data4[x]["idMembers"][0]]
        send_email(tName, tDate, tURL, tEmail)
        
        
##loop through card on 'Doing' list and email
response5 = requests.get("https://api.trello.com/1/lists/"+myList["Doing"]+"/cards?key="+key1+"&token="+token1)
data5 = response5.json()
for x in range(len(data5)):
    if dueDate(data5[x]["due"]) < datetime.date.today():
        tName = data5[x]["name"]
        tDate = dueDate(data5[x]["due"])
        tURL = data5[x]["shortUrl"]
        tEmail = Members[data5[x]["idMembers"][0]]
        send_email(tName, tDate, tURL, tEmail)
    

print("Done")
