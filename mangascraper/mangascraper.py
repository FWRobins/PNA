##basic web scraper to download manga scans
##has since been unusable after website updates

import requests, pandas, re, ast, pyautogui, os
from bs4 import BeautifulSoup
import urllib.request

##get manga details from user
manganame = pyautogui.prompt(text="Nameof manga?", title="Manga Name", default="")
version = pyautogui.prompt(text="Manga scan version? (case sensative)", title="Scans", default="Fox")
baseurl = pyautogui.prompt(text="Paste base URL", title="Manga URL", default="")


if os.path.isdir("./"+manganame):
    pass
else:
    os.makedirs("./"+manganame)

image = 1

mainr = requests.get(baseurl)
c=mainr.content
soup=BeautifulSoup(c, "html.parser")

all=soup.find_all("div", {"class":"mt-3 stream collapsed"})
# print(len(all))
t = None
for x in range(len(all)):
    if all[x].find_all("span")[0].text == "Version "+version:
        t = all[x]

if t == None:
    all=soup.find_all("div", {"class":"mt-3 stream"})
    print(len(all))
    for x in range(len(all)):
        if all[x].find_all("span")[0].text == "Version "+version:
            t = all[x]

list=[]
y = t.find_all("a")
for x in range(len(y)):
    if y[x].text == "all":
        #print(re.sub(".*/", "", y[x].get("href")))
        list.append("https://mangapark.net"+y[x].get("href"))
# print(list)
# print(len(list))
for chapter in range(len(list)):
    if chapter > 25:
        print(chapter)
    # for chapter in range(1):
        subr = requests.get(str(list[chapter]))
        c=subr.content
        soup=BeautifulSoup(c, "html.parser")
        # print(soup)

        x = soup.find_all("script")
        # x = soup.find_all("a", {"class":"img-link"})
        # print(x[4])
        # print(x[25:-5])
        # print(len(x[25:-5]))
        # print(str(x[25])[-14:-8].replace("/c", "").replace("/", ""))
        # pages = (x[-6].text)
        # # print(x[5])
        l = ast.literal_eval(x[4].text[20:-3].replace("\\", ""))
        print(l)
        print(len(l))
        for page in range(len(l)):

        #     urllib.request.urlretrieve(x[page], "./"+manganame+"/" + re.sub(".*/", "", list[chapter]) + "-" + str(page+1)+ ".jpg")
        #     # print(re.sub(".*/", "", list[chapter])+"-"+str(page+1))
        #     # if len(list)-chapter == 37:

            if "https:" in l[page]['u']:
                urllib.request.urlretrieve(l[page]['u'], "./"+manganame+"/" + re.sub(".*/", "", list[chapter]) + "-" + str(page+1)+ ".jpg")
            else:
                urllib.request.urlretrieve("https:"+l[page]['u'], "./"+manganame+"/" + re.sub(".*/", "", list[chapter]) + "-" + str(page+1)+ ".jpg")
