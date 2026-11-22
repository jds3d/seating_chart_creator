# -*- coding: utf-8 -*-
"""
Created on Sat Feb  9 12:13:07 2019

@author: Josh
"""
  
import sys
import tkinter
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook


def formatNameWithoutFamilyName(guest):
    if '-' in guest:
        guest = guest.split('-')[-1].strip()
    return guest    


## read guest data from pickle, but read seating chart from excel because it might be changed manually
def readSeatingChart(file_path=None):
    if not file_path:
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

        file_path = askopenfilename(
            parent=root,
            title='Select a seating chart',
            filetypes=[('Excel workbook', '*.xlsx'), ('All files', '*.*')],
        )
        root.destroy()
        if not file_path:
            print('No seating chart selected.')
            sys.exit(1)

    wb = load_workbook(file_path, data_only=True, read_only=True)
    try:
        xl_sheet = wb.worksheets[0]
        print('readSeatingChart: processing', xl_sheet.title)
        rows = xl_sheet.iter_rows(values_only=True)
        header = next(rows)
        column_headings = ['' if cell is None else str(cell) for cell in header]

        try:
            choice_col_index = column_headings.index("Choice")
        except ValueError:
            choice_col_index = None

        tablesDict = {}
        guestChoice = {}
        for row in rows:
            if not row or row[0] is None or row[1] is None:
                continue
            tableNum = row[0]
            guestName = str(row[1]).strip()
            if not guestName:
                continue

            if tableNum not in tablesDict:
                tablesDict[tableNum] = set()
            tablesDict[tableNum].add(guestName)

            if choice_col_index is not None and choice_col_index < len(row):
                val = row[choice_col_index]
                guestChoice[guestName] = val if val else ""
            else:
                guestChoice[guestName] = ""
    finally:
        wb.close()

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
    from gui import main
    main()
