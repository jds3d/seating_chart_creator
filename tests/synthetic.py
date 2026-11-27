# -*- coding: utf-8 -*-
"""Made-up Cheddar Up exports for tests. No real guests."""

import os
import sys

from openpyxl import Workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

ITEM_HEADERS = [
    'Payment Date', 'Payment Method', 'Last 4 of Acct', 'Payment Status',
    'Payer First Name', 'Payer Last Name', 'Payer Email', 'Qty Sold', 'Cost',
    'Subtotal', 'Discounts', 'Item Total', 'Refunded Amount', 'Quantity Refunded',
    'Net Amount', 'Net Quantity',
    'Attendee 1 Name', 'Choose your entree',
    'Attendee 2 Name', 'Choose your entree',
    'Requested Seatmate(s)',
]

# (first, last, email, status, [(attendee, entree)], seatmate request)
MEMBER_ROWS = [
    ('Alma', 'Archer', 'alma@example.com', 'available', [('Alma Archer', 'Chicken'), ('Boris Archer', 'Fish')], 'Cyrus Bell'),
    ('Cyrus', 'Bell', 'cyrus@example.com', 'available', [('Cyrus Bell', 'Beef'), ('Dahlia Bell', 'Chicken')], ''),
    ('Edgar', 'Cohen', 'edgar@example.com', 'available', [('Edgar Cohen', 'Fish')], ''),
    ('Fiona', 'Dunn', 'fiona@example.com', 'available', [('Fiona Dunn', 'Vegetarian'), ('Gus Dunn', 'Beef')], ''),
    ('Hazel', 'Eaton', 'hazel@example.com', 'available', [('Hazel Eaton', 'Chicken')], ''),
    ('Ivan', 'Foster', 'ivan@example.com', 'available', [('Ivan Foster', 'Fish'), ('Jade Foster', 'Fish')], ''),
    ('Kurt', 'Garcia', 'kurt@example.com', 'refunded', [('Kurt Garcia', 'Beef')], ''),
    ('Lena', 'Hughes', 'lena@example.com', 'available', [('Lena Hughes', 'Chicken')], ''),
    ('Milo', 'Ibarra', 'milo@example.com', 'available', [('Milo Ibarra', 'Beef'), ('Nora Ibarra', 'Vegetarian')], ''),
]
GUEST_ROWS = [
    ('Oscar', 'Jensen', 'oscar@example.com', 'available', [('Oscar Jensen', 'Chicken'), ('Petra Jensen', 'Fish')], ''),
    ('Quinn', 'Kowalski', 'quinn@example.com', 'available', [('Quinn Kowalski', 'Beef')], ''),
    ('Rosa', 'Lindqvist', 'rosa@example.com', 'available', [('Rosa Lindqvist', 'Vegetarian'), ('Silas Moreau', 'Chicken')], ''),
    ('Tobias', 'Nakamura', 'tobias@example.com', 'available', [('Tobias Nakamura', 'Fish')], ''),
]
REFUNDED = {'Kurt Garcia'}


def guest_names():
    names = set()
    for rows in (MEMBER_ROWS, GUEST_ROWS):
        for row in rows:
            if row[3] == 'refunded':
                continue
            names.update(name for name, _choice in row[4])
    return names


def _item_sheet(workbook, title, rows, cost):
    sheet = workbook.create_sheet(title)
    sheet.append(ITEM_HEADERS)
    for row in rows:
        first, last, email, status, attendees, seatmate = row[:6]
        # An optional 7th value is the ticket quantity, for purchases that name fewer people.
        qty = row[6] if len(row) > 6 else len(attendees)
        net_qty = 0 if status == 'refunded' else qty
        cells = [
            'Nov 20, 2026 10:00 AM EST', 'card', '4242', status, first, last, email,
            qty, cost, cost * qty, 0, cost * qty, cost * qty if status == 'refunded' else 0,
            qty if status == 'refunded' else None, cost * net_qty, net_qty,
        ]
        for index in range(2):
            name, choice = attendees[index] if index < len(attendees) else ('', '')
            cells.extend([name, choice])
        cells.append(seatmate)
        sheet.append(cells)
    return sheet


def write_export(path, extra_tabs=1, collection='TVNJC Test Dinner', member_rows=None, guest_rows=None):
    """A collection summary shaped like Cheddar Up's: guests on tabs 5 and 6."""
    workbook = Workbook()
    summary = workbook.active
    summary.title = 'Summary'
    summary.append(['Collection Name', collection])
    summary.append(['Date Created', 'Nov 01, 2026 09:00 AM EST'])
    workbook.create_sheet('Payments').append(['Payment ID', 'Payment Method', 'Status', 'First Name', 'Last Name'])
    workbook.create_sheet('All Purchases').append(['Payment ID', 'Item', 'Quantity'])
    workbook.create_sheet('Item Summary').append(['Item Name', 'Cost', 'Net Quantity'])
    _item_sheet(workbook, '1-Member-Dinner-40', MEMBER_ROWS if member_rows is None else member_rows, 40)
    _item_sheet(workbook, '2-Guest-Dinner-50', GUEST_ROWS if guest_rows is None else guest_rows, 50)
    for index in range(extra_tabs):
        workbook.create_sheet('{}-Raffle'.format(7 + index)).append(['Item Name', 'Qty'])
    workbook.save(path)
    return path


def write_requests(path):
    workbook = Workbook()
    requests = workbook.active
    requests.title = 'Requests'
    requests.append(['Name', 'Sit with'])
    requests.append(['Edgar Cohen', 'Hazel Eaton'])
    anti = workbook.create_sheet('Antirequests')
    anti.append(['Name', 'Do not sit with'])
    anti.append(['Fiona Dunn', 'Ivan Foster'])
    workbook.save(path)
    return path
