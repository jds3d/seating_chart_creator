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
        print(f'{header=}')
        print('If you want to have a poll it has be titled "Please choose one:"')
        poll_col_index = header.index("Please choose one:")

    except ValueError:
        print("❌ WARNING: No 'Please choose one:' column found in Punchbowl_Event_Guest_List sheet.")
            

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
                print(row_idx, poll_col_index)
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

    outputFilename = 'output/Table Rosters' + '_' + (
        'emails' if writeEmails else 'no emails') + '_' + timestamp + '.xls'
    # xlwt expects bytes for footer_str (Python 3 compatibility)
    footer_path = os.path.abspath(outputFilename).encode('utf-8')

    style = xlwt.XFStyle()

    if not writeEmails:
        # font - slightly smaller so it fits on one page
        font = xlwt.Font()
        font.height = 400  # 20pt - reduced from 640
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
        sheet.portrait = True
        sheet.footer_str = footer_path

        if writeEmails:
            sheet.write(row_idx, col_idx, 'Table #{} -- ({} people)'.format(i, len(table)), style=style)
            sheet.col(col_idx).width = 15000
            col_idx += 1
            sheet.write(row_idx, col_idx, 'Email')
            col_idx += 1
        else:
            # Row 0: "Table #1" centered and larger across the top (merged)
            title_style = xlwt.XFStyle()
            title_font = xlwt.Font()
            title_font.height = 560  # ~28pt - larger
            title_font.bold = True
            title_style.font = title_font
            title_style.alignment = xlwt.Alignment()
            title_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
            sheet.write_merge(0, 0, 0, 1, 'Table #{}'.format(i), style=title_style)
            sheet.col(0).width = 10000
            sheet.col(1).width = 8000

            # Row 1: Name and Choice headers (underlined with solid line, darker than dashed separators)
            header_style = xlwt.XFStyle()
            header_font = xlwt.Font()
            header_font.height = 400
            header_font.bold = True
            header_style.font = header_font
            header_borders = xlwt.Borders()
            header_borders.bottom = xlwt.Borders.THIN  # Solid line, darker than DASHED
            header_style.borders = header_borders
            sheet.write(1, 0, 'Name', style=header_style)
            sheet.write(1, 1, 'Choice', style=header_style)

            font = xlwt.Font()
            font.height = 400
            font.bold = True
            style = xlwt.XFStyle()
            style.font = font
            borders = xlwt.Borders()
            borders.bottom = xlwt.Borders.DASHED
            style.borders = borders

        for guest in sorted(table):
            row_idx += 1
            col_idx = 0

            if writeEmails:
                data_row = row_idx
            else:
                data_row = row_idx + 1  # Skip title row (0) and header row (1)

            if '-' in guest:
                sheet.write(data_row, col_idx, guest.split('-')[-1].strip(), style=style)
                col_idx += 1
            else:
                sheet.write(data_row, col_idx, guest, style=style)
                col_idx += 1
            try:
                if writeEmails:
                    sheet.write(data_row, col_idx, emails[formatNameWithoutFamilyName(guest)])
                else:
                    sheet.write(data_row, col_idx, extraColumn[guest], style=style)
            except:
                if writeEmails:
                    sheet.write(data_row, col_idx, '')

        if not writeEmails:
            # Footnote: "-- (8 people)" centered at bottom (row 0=title, 1=headers, 2+=guests)
            footnote_row = 2 + len(table)
            footnote_style = xlwt.XFStyle()
            footnote_font = xlwt.Font()
            footnote_font.height = 180  # ~9pt - small, for organizer only
            footnote_font.bold = False
            footnote_style.font = footnote_font
            footnote_style.alignment = xlwt.Alignment()
            footnote_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
            sheet.write_merge(footnote_row, footnote_row, 0, 1,
                             '({} people)'.format(len(table)), style=footnote_style)

        i += 1
    print('saving table rosters...', outputFilename)
    wb.save(outputFilename)
