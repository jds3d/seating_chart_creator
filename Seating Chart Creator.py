# -*- coding: utf-8 -*-
"""
Created on Sat Nov 17 10:07:26 2018

@author: jds3d
"""
import datetime
import pprint
import random
import sys

from utils import writeTables, writeSeatingChart, readGuestList

###### CONSTANTS ####################
###### You can edit these. ##########
normalTableMaxSize = 8  # no more than this many people will be sat at a table without being part of a larger group.
singleGroupMaxSize = 9  # above this number, approval will be needed.


def addGuestsRequests(group):
    for guestName in group.copy():
        try:
            group |= guests[guestName]
        except:
            pass
    return group


# generate seating chart
# first find largest groups and assign them to first table with enough space
def generateSeatingChart(guests, antiRequests):
    print('generating seating chart...')
    tables = [set()]  # initialize tables as a list of sets
    guestsSeated = set()

    # randomize the guestList so you don't have the same people sitting next to each other every time.       
    guestList = [(k, v) for k, v in guests.items()]
    random.shuffle(guestList)

    #    guestList.sort(key=lambda x:len(x[1]))
    for guest in sorted(guestList, key=lambda k: len(guestList[1]), reverse=True):
        k = guest[0]
        if k in guestsSeated: continue  # don't seat guests that have already been seated via other guests' baggage

        # count the number of guests in this group (family and baggage together)
        group = set()
        numInGroup = len(group)
        group = guests[k]
        while len(group) > numInGroup:
            numInGroup = len(group)
            group = addGuestsRequests(group)

        guests[k] = group

        if numInGroup > singleGroupMaxSize:
            print('Group has {} people, larger than max of {}'.format(numInGroup, singleGroupMaxSize))

            # ask user if it's ok.
            pprint.pprint(group)
            proceed = input("Do you want to allow it? Blank is no, anything else means yes: ")
            if not proceed:
                print('Exiting so you can fix the configuration.')
                sys.exit(1)

        # seat group
        groupAdded = False
        for table in tables:
            doNotSitHere = False
            # if there is space at the table:
            if normalTableMaxSize - len(table) - numInGroup >= 0:
                # if the people all want to sit together:
                if len(table) > 0:
                    for p1 in group:
                        if p1 in antiRequests:
                            for antiRequest in antiRequests[p1]:
                                if antiRequest in table:
                                    print(p1)
                                    # don't seat them here.
                                    doNotSitHere = True
                        if doNotSitHere: break
                # assign to table everyone in this group and the current person
                if doNotSitHere: continue  # try all the rest of the tables too before creating a new one.
                table |= guests[k]
                groupAdded = True
                break
        if not (groupAdded):
            table = set()
            table |= guests[k]
            tables.append(table)
        guestsSeated |= table

        # randomize the tables so that 
        random.shuffle(tables)
    return tables


def editTableNumbers(tables):
    # edit the table numbers?
    finish = ""
    while finish not in ('n', 'y'):
        finish = input("Would you like to edit the table numbers (y/n)? ")
        if finish == 'n': return tables
    tableNumbers = set(range(1, len(tables) + 4))  # +4 because we want up to 3 empty tables.

    # create tables as a dict with key tablenum and value=set of guests, then convert that to a list of sets in that order
    newTablesDict = {}
    i = 1
    for table in tables:
        print('Table {}'.format(i))
        for guest in table:
            print(guest)
        print('Table #s remaining: ', sorted(tableNumbers))
        tableNumber = 0

        while tableNumber not in tableNumbers:
            tableNumber = input(
                'What # should this table have? If you hit enter, it will be ' + str(min(tableNumbers)) + ': ')
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


# ask what number it should be.  show all numbers that are left.  enter will not change it.
if __name__ == "__main__":
    print("reading guest list...")
    guests, antiRequests, extraGuestData, emails, polls = readGuestList()
    print("generating seating chart...")
    tables = generateSeatingChart(guests, antiRequests)
    print("editing table numbers...")
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
    tables = editTableNumbers(tables)
    print("writing seating chart...")
    writeSeatingChart(tables, extraGuestData, timestamp)
    print("writing tables...")
    writeTables(tables, emails, polls, timestamp, writeEmails=True)
    writeTables(tables, emails, polls, timestamp, writeEmails=False, extraColumn=extraGuestData)
