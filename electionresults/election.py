##I made this to check the election results in our national election

import pandas as pd
import os
import openpyxl
from tkinter import *



window=Tk()

def progress():
    def sortSecond(val): 
        return val[1]
    
    url = 'https://www.elections.org.za/content/NPEPublicReports/699/Results%20Report/National.xls'
    df = pd.read_excel(url)
    url2 = 'https://www.elections.org.za/content/NPEPublicReports/827/Results%20Report/National.xls'
    df1 = pd.read_excel(url2)

##get data for provincial results
    wc = []
    for party, perc in zip(df1['Unnamed: 6'],df1['Unnamed: 29']):
        if str(party) != 'nan' and str(perc) != '0.00%':
            wc.append([party, int(perc*10000)/100])
    wc.sort(key=sortSecond)

##get data for national results
    nat = []
    for party, perc in zip(df['Unnamed: 6'],df['Unnamed: 34']):
        if str(party) != 'nan':
            nat.append([party, int(perc*10000)/100])
    nat.sort(key=sortSecond)

    t1.delete("1.0", END)
    t1.insert(END, str(wc[-1][0]))
    t2.delete("1.0", END)
    t2.insert(END, str(wc[-2][0]))
    t3.delete("1.0", END)
    t3.insert(END, str(wc[-3][0]))

    t5.delete("1.0", END)
    t5.insert(END, str(wc[-1][1]))
    t6.delete("1.0", END)
    t6.insert(END, str(wc[-2][1]))
    t7.delete("1.0", END)
    t7.insert(END, str(wc[-3][1]))

    t9.delete("1.0", END)
    t9.insert(END, str(nat[-1][0]))
    t10.delete("1.0", END)
    t10.insert(END, str(nat[-2][0]))
    t11.delete("1.0", END)
    t11.insert(END, str(nat[-3][0]))

    t13.delete("1.0", END)
    t13.insert(END, str(nat[-1][1]))
    t14.delete("1.0", END)
    t14.insert(END, str(nat[-2][1]))
    t15.delete("1.0", END)
    t15.insert(END, str(nat[-3][1]))
    

l1 = Label(window, text='Western Cape Provincial Vote:')
l1.grid(row=0, column=0, rowspan=2, columnspan=2)

l2 = Label(window, text='National Vote:')
l2.grid(row=0, column=2, rowspan=2)

l3 = Label(window, text='Party')
l3.grid(row=2, column=0)

l4 = Label(window, text='Vote')
l4.grid(row=2, column=1)

l5 = Label(window, text='Party')
l5.grid(row=2, column=2)

l6 = Label(window, text='Vote')
l6.grid(row=2, column=3)

t1=Text(window, height=1, width=20)
t1.grid(row=3, column=0)

t2=Text(window, height=1, width=20)
t2.grid(row=4, column=0)

t3=Text(window, height=1, width=20)
t3.grid(row=5, column=0)

t5=Text(window, height=1, width=20)
t5.grid(row=3, column=1)

t6=Text(window, height=1, width=20)
t6.grid(row=4, column=1)

t7=Text(window, height=1, width=20)
t7.grid(row=5, column=1)

t9=Text(window, height=1, width=20)
t9.grid(row=3, column=2)

t10=Text(window, height=1, width=20)
t10.grid(row=4, column=2)

t11=Text(window, height=1, width=20)
t11.grid(row=5, column=2)

t13=Text(window, height=1, width=20)
t13.grid(row=3, column=3)

t14=Text(window, height=1, width=20)
t14.grid(row=4, column=3)

t15=Text(window, height=1, width=20)
t15.grid(row=5, column=3)

b1=Button(window, text='Refresh', command=progress)
b1.grid(row=6, column=1, columnspan=2)

progress()

window.mainloop()

