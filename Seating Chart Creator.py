# -*- coding: utf-8 -*-
"""
Created on Sat Nov 17 10:07:26 2018

@author: jds3d
"""
import xlrd
import xlwt
import datetime
import sys
import pprint
import random
import pickle
import os


def removeDoubleSpaces(name):
    while '  ' in name:
        name = name.replace('  ', ' ')
    return name


# read data from Punchbowl_Event_Guest_List.1
def readGuestList():
    antiRequests = {} ## dictionary, key is guestName, value is who they can't sit with
    guests = {} ## dictionary, key is guestName and value is set of baggages
    emails = {}
    wb = xlrd.open_workbook('Seating Chart Creator.xls')
    xl_sheet = wb.sheet_by_name('Punchbowl_Event_Guest_List')
    
    for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows, don't include header        
        if xl_sheet.cell(row_idx, 4).value == 'Yes':

            guestName = removeDoubleSpaces(xl_sheet.cell(row_idx, 0).value.strip())
            email = xl_sheet.cell(row_idx, 1).value.strip()
            if "@" not in email and "." not in email and email != '':
                print(f"{guestName}--Invalid email address: {email}")
#           if guestName.count('-') > 1:
#                familyName = '-'.join(guestName.split('-')[:-1]).strip()
#                individualName = guestName.split('-')[-1].strip()
#                print('INFO: guest "{}" has multiple hyphens Family Name="{}", Individual Name="{}".'.format(guestName, familyName, individualName))
            if guestName in guests:
                print('ERROR: guest {} is duplicated!'.format(guestName))

            guests[guestName] = set()
            emails[formatNameWithoutFamilyName(guestName)] = email
            print (1111, guestName, email)
    
    ## read data from Requests
    xl_sheet = wb.sheet_by_name('Requests')    


    for row_idx in range(1, xl_sheet.nrows):
        guestName = removeDoubleSpaces(xl_sheet.cell(row_idx, 0).value.strip())
        request = removeDoubleSpaces(xl_sheet.cell(row_idx, 1).value.strip())
        if request != '':
            if request in guests:
                guests[request].add(guestName)
                try:
                    guests[guestName].add(request)                            
                except:
                    print('REQUEST IGNORED: "{}" is not attending this event, so the request with "{}" will be ignored.'.format(guestName, request))
            else:
                print('REQUEST IGNORED: "{}" is not attending this event, so the request with "{}" will be ignored.'.format(request, guestName))
    
    ## make sure each couple is together 
    for guestName, baggageSet in guests.items():
        ##if there is a hyphen, check for family members

        if '-' in guestName:            
            lastNameMatches = [key.strip() for key in guests
                               if '-'.join(guestName.split('-')[0]).strip() 
                               == '-'.join(key.split('-')[0]).strip()]
            
            for name in lastNameMatches:
                baggageSet.add(name)     
#            print(lastNameMatches)
    for guestName in guests:
        guests[guestName].add(guestName)
        
    xl_sheet = wb.sheet_by_name('Antirequests')
    for row_idx in range(1, xl_sheet.nrows):
        guestName = removeDoubleSpaces(xl_sheet.cell(row_idx, 0).value.strip())
        antiRequest = removeDoubleSpaces(xl_sheet.cell(row_idx, 1).value.strip())
        
        if antiRequest != '':
            if antiRequest in guests:                
                if antiRequest not in antiRequests: antiRequests[antiRequest] = set()
                antiRequests[antiRequest].add(guestName)
                try:
                    if guestName not in antiRequests: antiRequests[guestName] = set()
                    antiRequests[guestName].add(antiRequest)                            
                
                except:
                    print('ANTI-REQUEST IGNORED: "{}" is not attending this event, so the anti-request with "{}" will be ignored.'.format(guestName, antiRequest))
            else:
                print('ANTI-REQUEST IGNORED: "{}" is not attending this event, so the anti-request with "{}" will be ignored.'.format(antiRequest, guestName))

    print("Printing EMAILS", len(emails))
  
#    for guestName, email in emails.items(): 
#        print(f"2222 -- {guestName}, {email}")


    with open('guest_data.p', 'wb') as f:
        pickle.dump((guests, antiRequests, emails), f)    
    return guests, antiRequests, emails
        
def addGuestsRequests(group):
    for guestName in group.copy():            
#        if guestName in guests:
        try:
            group |= guests[guestName]
        except:
            pass
#        else:
#            print('FATAL: guest "{}" is not attending the event but are listed on the Requests tab.'.format(guestName))
    return group


## generate seating chart
## first find largest groups and assign them to first table with enough space

def generateSeatingChart(guests, antiRequests):
    print('generating seating chart...')
    tables = [set()] ## initialize tables as a list of sets
    guestsSeated = set()
    
#    for key, value in guests.items():
        
    guestList = [(k,v) for k,v in guests.items()]
    random.shuffle(guestList)
    
#    guestList.sort(key=lambda x:len(x[1]))
    for guest in sorted(guestList, key=lambda k: len(guestList[1]), reverse=True):
        k = guest[0]
        if k in guestsSeated: continue ## don't seat guests that have already been seated via other guests' baggage
            
        ## count the number of guests in this group (family and baggage together)        
        
        group = set()        
        numInGroup = len(group)
        group = guests[k]
        while len(group) > numInGroup:
            numInGroup = len(group)
            group = addGuestsRequests(group)
        
        guests[k] = group

        singleGroupMaxSize = 9        
        if numInGroup > singleGroupMaxSize: 
            print('Group has {} people, larger than max of {}'.format(numInGroup, singleGroupMaxSize))
            ## ask user if it's ok.
            pprint.pprint(group)
            proceed = input("Do you want to allow it? Blank is no, anything else means yes: ")
            if not proceed:
                sys.exit(1)
            
        ## seat group
        groupAdded = False
        normalTableMaxSize = 8
        for table in tables:  
            doNotSitHere = False
            ## if there is space at the table:
            if normalTableMaxSize - len(table) - numInGroup >= 0:
                ## if the people all want to sit together:
                if len(table) > 0:
                    for p1 in group:
                        if p1 in antiRequests:
                            for antiRequest in antiRequests[p1]:
                                if antiRequest in table:
                                    print(p1)
                                    ## don't seat them here.
                                    doNotSitHere = True
                        if doNotSitHere: break
                  ## assign to table everyone in this group and the current person 
                if doNotSitHere: continue ## try all the rest of the tables too before creating a new one.
                table |= guests[k]
                groupAdded = True
                break
        if not(groupAdded):
            table = set()
            table |= guests[k]
            tables.append(table)
        guestsSeated |= table
    return tables
    

#for i in range(len(tables)):
#    print(i, tables[i])
    
## Task List:
## DONE: randomize seating
## DONE: notify about 2 hyphens
## DONE: notify about duplicate guests
        
def formatNameWithoutFamilyName(guest):
    if '-' in guest:
        guest = guest.split('-')[-1].strip()
    return guest    

def writeSeatingChart(tables, timestamp):
    ## write seating chart into output file
    wb = xlwt.Workbook()
    
    sheet = wb.add_sheet("Seating Chart")
    
    row_idx = 0    
    col_idx = 0
    sheet.write(row_idx, col_idx, 'Table Number'); col_idx += 1
    sheet.write(row_idx, col_idx, 'Name'); col_idx += 1
    sheet.write(row_idx, col_idx, 'Number of People'); col_idx += 1
    
    row_idx = 0
    tableNum = 1
    for table in tables:
        for guest in sorted(table):
            row_idx += 1
            sheet.write(row_idx, 0, tableNum)
            sheet.write(row_idx, 1, formatNameWithoutFamilyName(guest))
            
            sheet.write(row_idx, 2, len(table))
        tableNum += 1
    
    outputFilename = 'output/Seating Chart' + '_' + timestamp + '.xls'
    print('saving seating chart...', outputFilename)
    wb.save(outputFilename)
    
def writeTables(tables, emails, timestamp, writeEmails):
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

            if '-' in guest:
                sheet.write(row_idx, 0, guest.split('-')[-1].strip(), style=style); col_idx += 1                
            else:    
                sheet.write(row_idx, 0, guest, style=style); col_idx += 1



            try:
                if writeEmails: sheet.write(row_idx, 1,  emails[formatNameWithoutFamilyName(guest)])
            except:
                if writeEmails: sheet.write(row_idx, 1, '')
            
        i += 1
    outputFilename = 'output/Table Rosters' + '_' + ('emails' if writeEmails else 'no emails') + '_' + timestamp + '.xls'
    print('saving table rosters...', outputFilename)
    wb.save(outputFilename)


def editTableNumbers(tables):
    ## edit the table numbers?
    finish = ""
    while finish not in ('n', 'y'):
        finish = input("Would you like to edit the table numbers (y/n)? ")
        if finish == 'n': return tables
    tableNumbers = set(range(1, len(tables)+4)) ## +4 because we want up to 3 empty tables.

    ## create tables as a dict with key tablenum and value=set of guests, then convert that to a list of sets in that order
    newTablesDict = {}
    i = 1
    for table in tables:
        print ('Table {}'.format(i))
        for guest in table:
            
            
            print(guest)
        print('Table #s remaining: ', sorted(tableNumbers))
        tableNumber = 0
        
        
        
        
        while tableNumber not in tableNumbers:
            tableNumber = input('What # should this table have? If you hit enter, it will be ' + str(min(tableNumbers)) + ': ')
            if tableNumber == "":
                tableNumber = min(tableNumbers)
            else:
                tableNumber = int(tableNumber)
        newTablesDict[tableNumber] = table
        tableNumbers.remove(tableNumber)
        i += 1
    
    newTables = []
    
    for tableNumber in sorted(newTablesDict):
        newTables.append(newTablesDict[tableNumber])
    return newTables
             
             
    ## ask what number it should be.  show all numbers that are left.  enter will not change it.
    
if __name__ == "__main__":    
    print("reading guest list...")
    guests, antiRequests, emails = readGuestList()
    print("generating seating chart...")
    tables = generateSeatingChart(guests, antiRequests)
    print("editing table numbers...")
    timestamp=datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')  
    tables = editTableNumbers(tables)
    print("writing seating chart...")
    writeSeatingChart(tables, timestamp)    
    print("writing tables...")
    writeTables(tables, emails, timestamp, True)
    writeTables(tables, emails, timestamp, False)
    
    