# -*- coding: utf-8 -*-
"""
Created on Sat Feb  9 12:13:07 2019

@author: Josh
"""
  
import datetime
import tkinter
from tkinter.filedialog import askopenfilename
import pickle
import xlrd
from utils import writeTables


def formatNameWithoutFamilyName(guest):
    if '-' in guest:
        guest = guest.split('-')[-1].strip()
    return guest    


## read guest data from pickle, but read seating chart from excel because it might be changed manually
def readSeatingChart():
    # Make a top-level instance and hide since it is ugly and big.
    root = tkinter.Tk()
    root.withdraw()

    # Make it almost invisible - no decorations, 0 size, top left corner.
    root.overrideredirect(True)
    root.geometry('0x0+0+0')

    # Show window again and lift it to top so it can get focus,
    # otherwise dialogs will end up behind the terminal.
    root.deiconify()
    root.lift()
    root.focus_force()

    file_path = askopenfilename(parent=root)  # Or some other dialog

    wb = xlrd.open_workbook(file_path)
    xl_sheet = wb.sheet_by_index(0)
    print('readSeatingChart: processing', xl_sheet.name)

    # Extract column headings from the first row
    column_headings = [xl_sheet.cell(0, col).value for col in range(xl_sheet.ncols)]

    # Find Choice column index (for table roster output)
    try:
        choice_col_index = column_headings.index("Choice")
    except ValueError:
        choice_col_index = None

    # Create tables List of sets of guests from excel file; also map guest -> choice
    tablesDict = {}
    guestChoice = {}
    for row_idx in range(1, xl_sheet.nrows):  # Iterate through rows, excluding the header
        tableNum = xl_sheet.cell(row_idx, 0).value
        guestName = xl_sheet.cell(row_idx, 1).value.strip()

        if tableNum not in tablesDict:
            tablesDict[tableNum] = set()

        tablesDict[tableNum].add(guestName)

        if choice_col_index is not None and choice_col_index < xl_sheet.ncols:
            val = xl_sheet.cell(row_idx, choice_col_index).value
            guestChoice[guestName] = val if val else ""
        else:
            guestChoice[guestName] = ""

    tablesList = []
    for tableNum in sorted(tablesDict):
        tablesList.append(tablesDict[tableNum])

    return tablesList, guestChoice


def ask_for_column(column_headings):
    # Ask the user to select a column heading
    print("Available Column Headings:")
    for index, heading in enumerate(column_headings, 1):
        print(f"{index}. {heading}")

    selected_index = input("Enter the number corresponding to the column heading you want to select: ")

    try:
        selected_index = int(selected_index)
        if 1 <= selected_index <= len(column_headings):
            selected_heading = column_headings[selected_index - 1]
            print(f"You selected: {selected_heading}")
        else:
            print("Invalid input. Please enter a valid number.")
    except ValueError:
        print("Invalid input. Please enter a number.")
    return selected_index


if __name__ == "__main__":    
    timestamp=datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
    with open('guest_data.p', 'rb') as f:    
        guests, antiRequests, emails, polls = pickle.load(f)    

    tables, guestChoice = readSeatingChart()

    writeTables(tables, emails, polls, timestamp, writeEmails=True, extraColumn=guestChoice)
    writeTables(tables, emails, polls, timestamp, writeEmails=False, extraColumn=guestChoice)
