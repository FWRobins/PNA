##this was to select a random branch and staff whos till would have to be spot checked

##openpyxl module for pandas is needed, install first
import pandas as pd
from datetime import datetime
import random
import PySimpleGUI as sg

##clear history in past.txt file on mondays
# print(datetime.now().weekday())
if datetime.now().weekday() == 1:
    clean_df = pd.DataFrame({'branch': [], 'day': []})
    clean_df.to_csv('past.txt', index=None)

##past.txt file is to control previouse results
past_file = pd.read_csv('past.txt')
# print(past_file)

##base file is list of staff members and branch locations to work from
base_file = pd.read_excel('Till Audit.xlsx', names=['Account', 'Name', 'Alternative_name', 'CRlimit'])
branches = [1, 2, 3, 4, 6, 7, 9, 10]

##what branches have been checked this week
branches_used = [None]
for row, index in past_file.iterrows():
    branches_used.append(past_file.iloc[row]['branch'])

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

staff = []
for row, index in base_file.iterrows():
    if base_file.iloc[row]['Alternative_name'] == branch:
        staff.append(base_file.iloc[row]['Name'])

random_staff1 = random_staf()
random_staff2 = random_staf()
while random_staff1 == random_staff2:
    random_staff2 = random_staf()

##add results to past.txt file
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

print(
    f'branch to be checked is {branch_dict[str(branch)]}, first staff member is {random_staff1}, else {random_staff2}')
##display to user
sg.Popup(
    f'Branch to be checked is {branch_dict[str(branch)]}, first staff member is {random_staff1}, else {random_staff2}')
