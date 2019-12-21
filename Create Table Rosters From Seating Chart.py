# -*- coding: utf-8 -*-
"""
Created on Sat Feb  9 12:13:07 2019

@author: Josh
"""
  
import datetime
import tkinter
from tkinter.filedialog import askopenfilename
try:
    import cPickle as pickle

except:
    import pickle
import xlrd, xlwt

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
    
    file_path = askopenfilename(parent=root) # Or some other dialog

#    file_path = filedialog.askopenfilename(parent=root, title = "Select Seating Chart_yyyy-mm-dd_tttttt.xls file.")
    
    
    wb = xlrd.open_workbook(file_path)
    xl_sheet = wb.sheet_by_index(0)    
    print('readSeatingChart: processing', xl_sheet.name)
    ## create tables List of sets of guests from excel file
    tablesDict = {}
    for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows, don't include header    
        tableNum = xl_sheet.cell(row_idx, 0).value   
        
        guestName = xl_sheet.cell(row_idx, 1).value.strip()
        if tableNum not in tablesDict:
            tablesDict[tableNum] = set()
            
        tablesDict[tableNum].add(guestName)
#        print('readSeatingChart: table number {}: guest {}'.format(tableNum, guestName))
        
    tablesList = []
    for tableNum in sorted(tablesDict):
        tablesList.append(tablesDict[tableNum])
    return tablesList

def writeTables(tables, emails, timestamp, writeEmails):
    print('writeTables...')
    wb = xlwt.Workbook()
    i = 1

    style = xlwt.XFStyle()
    
    
    if not writeEmails:

        # font
        font = xlwt.Font()
        font.height = 640 
        font.bold = True
        style.font = font
        
        # borders
        borders = xlwt.Borders()
        borders.bottom = xlwt.Borders.DASHED
        style.borders = borders
        

    for table in tables:
        row_idx = 0
        col_idx = 0   
        sheet = wb.add_sheet("Table {}".format(i))
        sheet.portrait = writeEmails
        sheet.write(row_idx, col_idx, 'Table {} -- {} people'.format(i, len(table)), style=style); col_idx += 1
        if writeEmails: sheet.write(row_idx, col_idx, 'Email'); col_idx += 1
        
        
        
        for guest in sorted(table):
            row_idx += 1
            col_idx = 0
            style = xlwt.XFStyle()
            if not writeEmails:
                # font
                font = xlwt.Font()
                font.height = 640 
                style.font = font

                
            sheet.write(row_idx, 0, guest, style=style); col_idx += 1
            try:
                if writeEmails: sheet.write(row_idx, 1, emails[guest])
            except:
                print('ERROR: writeTables: missing the email address for {}'.format(guest))
#            print('writeTables: ', row_idx, guest)
            
        i += 1
    outputFilename = 'output/Table Rosters' + '_' + ('emails' if writeEmails else 'no emails') + '_' + timestamp + '.xls'
    print('saving table rosters...', outputFilename)
    wb.save(outputFilename)    

if __name__ == "__main__":    

    timestamp=datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')  
    with open('guest_data.p', 'rb') as f:    
        guests, antiRequests, emails = pickle.load(f)    
        
    tables = readSeatingChart()
    
    writeTables(tables, emails, timestamp, True)
    writeTables(tables, emails, timestamp, False)