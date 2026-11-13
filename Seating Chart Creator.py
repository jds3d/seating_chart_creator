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


def addGuestsRequests(group, guests):
    for guestName in group.copy():
        try:
            group |= guests[guestName]
        except:
            pass
    return group


def _expandAllGroups(guests):
    """Expand all guest groups (family + requests) until no more changes."""
    changed = True
    while changed:
        changed = False
        for k in guests:
            old_len = len(guests[k])
            guests[k] = addGuestsRequests(guests[k].copy(), guests)
            if len(guests[k]) > old_len:
                changed = True


def _tables_conflict(left, right, antiRequests):
    for person in left:
        blocked = antiRequests.get(person, ())
        if any(other in blocked for other in right):
            return True
    for person in right:
        blocked = antiRequests.get(person, ())
        if any(other in blocked for other in left):
            return True
    return False


def _merge_tables_that_fit(tables, antiRequests):
    """Join leftover tables when the combined group still fits and nobody asked to be separated."""
    tables = [table for table in tables if table]
    merged = True
    while merged:
        merged = False
        tables.sort(key=len)
        for i in range(len(tables)):
            for j in range(i + 1, len(tables)):
                if len(tables[i]) + len(tables[j]) > normalTableMaxSize:
                    continue
                if _tables_conflict(tables[i], tables[j], antiRequests):
                    continue
                print('Combining a table of {} with a table of {}.'.format(len(tables[i]), len(tables[j])))
                tables[i] |= tables[j]
                del tables[j]
                merged = True
                break
            if merged:
                break
    return tables


# generate seating chart
# first find largest groups and assign them to first table with enough space
# singles (people with no seating requests) are seated together in their own tables
# then small leftover tables are combined when they still fit
def generateSeatingChart(guests, antiRequests):
    print('generating seating chart...')
    tables = [set()]  # initialize tables as a list of sets
    guestsSeated = set()

    # expand all groups first so we can correctly identify singles vs groups
    _expandAllGroups(guests)

    # randomize the guestList so you don't have the same people sitting next to each other every time.
    guestList = [(k, v) for k, v in guests.items()]
    random.shuffle(guestList)

    # split into groups (size > 1) and singles (size == 1)
    groupsToSeat = [(k, v) for k, v in guestList if len(v) > 1]
    singlesToSeat = [(k, v) for k, v in guestList if len(v) == 1]

    # sort groups by size (largest first)
    groupsToSeat = sorted(groupsToSeat, key=lambda x: len(x[1]), reverse=True)

    # Phase 1: seat groups (families, couples, etc.)
    for guest in groupsToSeat:
        k = guest[0]
        if k in guestsSeated:
            continue

        group = guests[k]
        numInGroup = len(group)

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
                        if doNotSitHere:
                            break
                # assign to table everyone in this group and the current person
                if doNotSitHere:
                    continue  # try all the rest of the tables too before creating a new one.
                table |= guests[k]
                groupAdded = True
                break
        if not groupAdded:
            table = set()
            table |= guests[k]
            tables.append(table)
        guestsSeated |= group

        # randomize the tables so that
        random.shuffle(tables)

    # Phase 2: seat singles together on dedicated tables
    for guest in singlesToSeat:
        k = guest[0]
        if k in guestsSeated:
            continue

        group = guests[k]
        groupAdded = False

        # try to add to an existing singles-only table with space
        for table in tables:
            if len(table) >= normalTableMaxSize:
                continue
            # only add to tables that have only singles (or are empty)
            if any(len(guests[p]) > 1 for p in table):
                continue
            # check anti-requests
            doNotSitHere = False
            for p1 in group:
                if p1 in antiRequests:
                    for antiRequest in antiRequests[p1]:
                        if antiRequest in table:
                            doNotSitHere = True
                            break
                if doNotSitHere:
                    break
            if doNotSitHere:
                continue
            table |= group
            groupAdded = True
            break

        if not groupAdded:
            new_table = set()
            new_table |= group
            tables.append(new_table)
        guestsSeated |= group

        random.shuffle(tables)

    return _merge_tables_that_fit(tables, antiRequests)


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
    export_path = sys.argv[1] if len(sys.argv) > 1 else None
    guests, antiRequests, extraGuestData, emails, polls = readGuestList(export_path)
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
