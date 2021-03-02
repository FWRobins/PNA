##this application was to select 4 random staff members from a random branch who had to do/undergo a bag search.

import pandas as pd
from datetime import datetime
import random
import PySimpleGUI as sg


def bag_check():
    
##    past.txt is a control file to save previous results

    past_file = pd.read_csv('past.txt')

##excel sheet consists of staff names and branch locations
    base_file = pd.read_excel('bag search.xlsx', names=['Account', 'Name', 'Alternative_name', 'CRlimit'])
    branches = [1, 2, 3, 4, 6, 7, 9, 10]

##re-create list of branched that have been selected
    branches_used = [None]
    for row, index in past_file.iterrows():
        branches_used.append(past_file.iloc[row]['branch'])

##if all branched have been used, clear list and pastfile.txt
    if len(branches_used) == 9:
        branches_used = [None]
        clean_df = pd.DataFrame({'branch': [], 'day': []})
        clean_df.to_csv('past.txt', index=None)
        past_file = pd.read_csv('past.txt')

    def random_branch():
        random.seed(datetime.now())
        branch = random.choice(branches)
        return branch

    def random_staf():
        random.seed(datetime.now())
        staff_member = random.choice(staff)
        return staff_member

    branch = None
    while branch in branches_used:
        branch = random_branch()

    staff = ["All"]
    for row, index in base_file.iterrows():
        if base_file.iloc[row]['Alternative_name'] == branch:
            staff.append(base_file.iloc[row]['Name'])

    random_staff1 = random_staf()
    while random_staff1 == "All":
        random_staff1 = random_staf()

    random_staff2 = random_staf()
    while random_staff2 == random_staff1 or random_staff2 == "All":
        random_staff2 = random_staf()

    random_staff3 = random_staf()
    while random_staff3 == random_staff1 or random_staff3 == random_staff2:
        random_staff3 = random_staf()

    random_staff4 = random_staf()
    while random_staff4 == random_staff1 or random_staff4 == random_staff2 or random_staff4 == random_staff3:
        random_staff4 = random_staf()

    add_df = pd.DataFrame({'branch': [branch], 'day': [datetime.now().weekday()]})
    past_file_new = past_file.append(add_df, ignore_index=True)
    past_file_new.to_csv('past.txt', index=None)

    branch_dict = {'1': 'Strand',
                   '2': 'Colours',
                   '3': 'Swes',
                   '4': 'Eikestad',
                   '6': 'Sanctuary',
                   '7': 'WhaleCoast',
                   '9': 'GordonsBay',
                   '10': 'Swellendam'}
    checking_branch = branch_dict[str(branch)]

    return [checking_branch, random_staff1, random_staff2, random_staff3, random_staff4]

##get data and disply in window via GUI
names_list = bag_check()
print(names_list)
sg.theme('DarkAmber')  # Add a touch of color

layout = [[sg.Text(f'Branch selected is {names_list[0]}, ', key='-text1-', size=(400, 1))],
          [sg.Text(f'Staff member to check {names_list[1]}, else {names_list[2]},', key='-text2-', size=(400, 1))],
          [sg.Text(f'Staff member to be checked is {names_list[3]}, else {names_list[4]}', key='-text3-', size=(400, 1))],
          [sg.Button('ReRun', key='-rerun-'), sg.Button('Cancel')]]

window = sg.Window('Bag Check Randomizer', layout, size=(500, 150))

while True:
    event, values = window.read()
    if event in (None, 'Cancel'):  # if user closes window or clicks cancel
        break
    if event == '-rerun-':
        names_list = bag_check()
        print(f'Branch selected is {names_list[0]},')
        print(f'Staff member to check {names_list[1]}, else {names_list[2]},')
        print(f'Staff to be checked is {names_list[3]}, else {names_list[4]}')
        window['-text1-'].update(f'Branch selected is {names_list[0]},')
        window['-text2-'].update(f'Staff member to check {names_list[1]}, else {names_list[2]},')
        window['-text3-'].update(f'Staff member to be checked is {names_list[3]}, else {names_list[4]}')

window.close()
