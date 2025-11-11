import xlrd
import xlwt
import pickle
import os
import sys

def removeDoubleSpaces(name):
    while '  ' in name:
        name = name.replace('  ', ' ')
    return name


# read data from Punchbowl_Event_Guest_List
def readGuestList():
    antiRequests = {}  # Dictionary, key is guestName, value is who they can't sit with
    guests = {}  # Dictionary, key is guestName and value is a set of requests
    extraGuestData = {}  # Dictionary, key is guestName and value is from the user-selected column
    emails = {}
    polls = {}

    # Open the Excel workbook
    wb = xlrd.open_workbook('Seating Chart Creator.xls')

    # Read the Punchbowl_Event_Guest_List sheet
    xl_sheet = wb.sheet_by_name('Punchbowl_Event_Guest_List')

    # Read header row directly from the sheet
    header = [xl_sheet.cell(0, col_index).value for col_index in range(xl_sheet.ncols)]

    try:
        name_col_index = header.index("Name")
    except ValueError:
        print("❌ ERROR: No 'Name' column found in Punchbowl_Event_Guest_List sheet.")
        sys.exit(1)   # stop execution

    try:
        try:
            email_col_index = header.index("Email")
        except ValueError:
            email_col_index = header.index("Email/Phone")
    except ValueError:
        print("❌ ERROR: No 'Email' column found in Punchbowl_Event_Guest_List sheet.")
        sys.exit(1)   # stop execution

    poll_col_index = None
    try:
        print(header)
        # poll_col_index = header.index("Please choose one:")
        poll_col_index = 6
    except ValueError:
        print("❌ ERROR: No 'Please choose one:' column found in Punchbowl_Event_Guest_List sheet.")
            

    try:
        rsvp_col_index = header.index("RSVP")
    except ValueError:
        print("❌ ERROR: No 'RSVP' column found in Punchbowl_Event_Guest_List sheet.")
        sys.exit(1)   # stop execution



    # Ask the user to select the column
    print("Available Columns in 'Punchbowl_Event_Guest_List' sheet:")
    print('0. No additional column required')
    for col_index in range(xl_sheet.ncols):
        print(f"{col_index + 1}. {xl_sheet.cell(0, col_index).value}")

    selected_col_index = int(input("If you would like an additional column in the seating chart, select it here or (0)?: ")) - 1

    for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows, don't include header
        if xl_sheet.cell(row_idx, rsvp_col_index).value == 'Yes':

            guestName = removeDoubleSpaces(xl_sheet.cell(row_idx, name_col_index).value.strip())
            email = xl_sheet.cell(row_idx, email_col_index).value.strip()
            
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

            poll = None
            if poll_col_index:
                poll = xl_sheet.cell(row_idx, poll_col_index).value.strip()
                polls[formatNameWithoutFamilyName(guestName)] = poll

            # Extract the data from the user-selected column
            if selected_col_index != 0:
                extraGuestData[guestName] = xl_sheet.cell(row_idx, selected_col_index).value
            print(row_idx, guestName, email, poll)

    ## read data from Requests
    xl_sheet = wb.sheet_by_name('Requests')

    for row_idx in range(1, xl_sheet.nrows):
        guestName = removeDoubleSpaces(xl_sheet.cell(row_idx, 0).value.strip())
        request = removeDoubleSpaces(xl_sheet.cell(row_idx, 1).value.strip())
        if request != '':
            if request in guests and guestName in guests:
                guests[request].add(guestName)
                guests[guestName].add(request)
                print(f'Adding request: {request=}, {guestName=}')
            else:
                if guestName not in request:
                    print('REQUEST IGNORED: "{}" is not attending this event, so the request with "{}" will be ignored.'.format(guestName, request))
                else:
                    print('REQUEST IGNORED: "{}" is not attending this event, so the request with "{}" will be ignored.'.format(request, guestName))
                

    ## make sure each couple is together
    for guestName, baggageSet in guests.items():
        # if there is a hyphen, check for family members

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

    with open('guest_data.p', 'wb') as f:
        pickle.dump((guests, antiRequests, emails, polls), f)
    return guests, antiRequests, extraGuestData, emails, polls


def formatNameWithoutFamilyName(guest):
    if '-' in guest:
        guest = guest.split('-')[-1].strip()
    return guest


def writeSeatingChart(tables, extraGuestData, timestamp):
    ## write seating chart into output file
    wb = xlwt.Workbook()

    sheet = wb.add_sheet("Seating Chart")

    row_idx = 0
    col_idx = 0
    sheet.write(row_idx, col_idx, 'Table Number');
    col_idx += 1
    sheet.write(row_idx, col_idx, 'Name');
    col_idx += 1
    sheet.write(row_idx, col_idx, 'Number of People');
    col_idx += 1
    sheet.write(row_idx, col_idx, 'Choice');
    col_idx += 1

    row_idx = 0
    tableNum = 1
    for table in tables:
        for guest in sorted(table):
            row_idx += 1
            sheet.write(row_idx, 0, tableNum)
            sheet.write(row_idx, 1, formatNameWithoutFamilyName(guest))
            sheet.write(row_idx, 2, len(table))
            if guest in extraGuestData:
                sheet.write(row_idx, 3, extraGuestData[guest])
        tableNum += 1
        
    outputFilename = 'output/Seating Chart' + '_' + timestamp + '.xls'
    print('saving seating chart...', outputFilename)

    if not os.path.exists('output'):
        os.makedirs('output')

    wb.save(outputFilename)


def writeTables(tables, emails, polls, timestamp, writeEmails, extraColumn=None):
    wb = xlwt.Workbook()
    i = 1

    style = xlwt.XFStyle()

    if not writeEmails:
        # font
        font = xlwt.Font()
        font.height = 640
        font.bold = False
        style.font = font

        # borders
        borders = xlwt.Borders()
        borders.bottom = xlwt.Borders.DASHED
        style.borders = borders

    for table in tables:
        row_idx = 0
        col_idx = 0
        sheet = wb.add_sheet("Table {}".format(i))
        sheet.portrait = True
        sheet.write(row_idx, col_idx, 'Table #{} -- ({} people)'.format(i, len(table)), style=style)
        sheet.col(col_idx).width = 15000
        col_idx += 1

        if writeEmails:
            sheet.write(row_idx, col_idx, 'Email')
            col_idx += 1
        else:
            # font
            sheet.write(row_idx, col_idx, 'Choice', style=style)
            col_idx += 1

            font = xlwt.Font()
            font.height = 640
            font.bold = False
            style = xlwt.XFStyle()
            style.font = font

        for guest in sorted(table):
            row_idx += 1
            col_idx = 0

            if '-' in guest:
                sheet.write(row_idx, col_idx, guest.split('-')[-1].strip(), style=style)
                col_idx += 1
            else:
                sheet.write(row_idx, col_idx, guest, style=style)
                col_idx += 1
            try:
                if writeEmails:
                    sheet.write(row_idx, col_idx, emails[formatNameWithoutFamilyName(guest)])
                else:
                    # print(1, extraColumn)
                    sheet.write(row_idx, col_idx, extraColumn[guest], style=style)
                    col_idx += 1
                    sheet.write(row_idx, col_idx, polls[formatNameWithoutFamilyName(guest)])
            except:
                if writeEmails:
                    sheet.write(row_idx, col_idx, '')

        i += 1
    outputFilename = 'output/Table Rosters' + '_' + (
        'emails' if writeEmails else 'no emails') + '_' + timestamp + '.xls'
    print('saving table rosters...', outputFilename)
    wb.save(outputFilename)
