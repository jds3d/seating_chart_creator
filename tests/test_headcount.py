# -*- coding: utf-8 -*-
"""Headcount: every attendee named, counted the way the seating chart reads guests."""

import importlib.util
import os
import random
import shutil
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO

import synthetic

import cheddar_up
import corrections
import utils

# Shaped like a real export: a party repeated on a second ticket row, a purchase
# that names nobody, a blank attendee slot, refunded and zero-quantity rows, and
# someone on the guest tab who is also on the member tab.
MEMBERS = [
    ('Alma', 'Archer', 'alma@example.com', 'available', [('Alma Archer', 'Chicken'), ('Boris Archer', 'Fish')], '', 1),
    ('Alma', 'Archer', 'alma@example.com', 'available', [('Alma Archer', 'Chicken'), ('Boris Archer', 'Fish')], '', 1),
    ('Cyrus', 'Bell', 'cyrus@example.com', 'available', [], '', 2),
    ('Edgar', 'Cohen', 'edgar@example.com', 'available', [('', 'Fish'), ('Edgar Cohen', 'Beef')], ''),
    ('Fiona', 'Dunn', 'fiona@example.com', 'refunded', [('Fiona Dunn', 'Beef'), ('Gus Dunn', 'Beef')], ''),
    ('Hazel', 'Eaton', 'hazel@example.com', 'available', [('Hazel Eaton', 'Fish')], '', 0),
]
GUESTS = [
    ('Ivan', 'Foster', 'ivan@example.com', 'available', [('Ivan Foster', 'Fish'), ('Edgar Cohen', 'Beef')], ''),
]


def _load_creator():
    path = os.path.join(synthetic.ROOT, 'Seating Chart Creator.py')
    spec = importlib.util.spec_from_file_location('seating_creator_headcount', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class HeadcountTests(unittest.TestCase):
    def setUp(self):
        self.work = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.work, True)
        self.export = synthetic.write_export(
            os.path.join(self.work, 'TVNJC_Count-collection_summary.xlsx'), member_rows=MEMBERS, guest_rows=GUESTS)
        self.requests = synthetic.write_requests(os.path.join(self.work, 'Requests.xlsx'))

    def read(self):
        with redirect_stdout(StringIO()):
            return utils.readGuestList(self.export, self.requests)

    def test_counts(self):
        counts = utils.headcount(self.read())
        # Alma, Boris, Cyrus (payer only), Edgar on tab 5; Ivan on tab 6.
        self.assertEqual(counts['people'], 5)
        self.assertEqual([tab['new'] for tab in counts['tabs']], [4, 1])
        # Names listed: 2 + 2 + 1 (payer) + 1 on tab 5, 2 on tab 6.
        self.assertEqual([tab['heads'] for tab in counts['tabs']], [6, 2])
        self.assertEqual(counts['names'], 8)
        self.assertEqual(counts['repeats'], ['Alma Archer', 'Boris Archer', 'Edgar Cohen'])
        self.assertEqual(counts['payer_only'], 1)
        self.assertEqual(counts['skipped_rows'], 2)
        self.assertEqual(counts['tickets'], 1 + 1 + 2 + 2 + 2)

    def test_lines(self):
        lines = utils.headcount_lines(utils.headcount(self.read()))
        self.assertEqual(lines[0], 'Headcount: 5 guests (4 from tab 5 + 1 from tab 6).')
        self.assertEqual(lines[1], 'Names listed: 8 (6 on tab 5 + 2 on tab 6).')
        self.assertIn('3 of those repeat a name already listed', lines[2])
        self.assertIn('Alma Archer, Boris Archer, Edgar Cohen', lines[2])
        self.assertIn('1 purchase named nobody, so only the payer is counted.', lines)
        self.assertIn('Tickets paid (Net Quantity): 8, which is 3 more than the guests named.', lines)
        self.assertIn('2 refunded or zero-quantity rows not counted.', lines)

    def test_matches_the_guests_the_chart_seats(self):
        book = self.read()
        counts = utils.headcount(book)
        self.assertEqual(counts['people'], len(book.guests))
        state = corrections.build_state(book, {key: [] for key in corrections._SHEETS})
        corrections.apply_all(book, state)
        book.finish()
        random.seed(1)
        with redirect_stdout(StringIO()):
            tables = _load_creator().generateSeatingChart(book.guests, book.anti, confirm_large_group=lambda *a: True)
        self.assertEqual(sum(len(table) for table in tables), counts['people'])

    def test_download_summary_uses_the_same_count(self):
        summary = cheddar_up.describe_export(self.export)
        self.assertEqual(summary['problems'], [])
        counts = summary['headcount']
        self.assertEqual(counts['people'], utils.headcount(self.read())['people'])
        self.assertEqual([tab['heads'] for tab in summary['guest_tabs']], [6, 2])
        self.assertEqual([tab['guests'] for tab in summary['guest_tabs']], [4, 1])
        lines = cheddar_up.summary_lines(summary)
        self.assertIn("Tab 5 '1-Member-Dinner-40': 6 rows, 2 attendee name columns, 1 seatmate request column, "
                      "6 names listed, 4 different guests (2 refunded or zero-quantity rows skipped).", lines)
        self.assertIn("Tab 6 '2-Guest-Dinner-50': 1 row, 2 attendee name columns, 1 seatmate request column, "
                      "2 names listed, 1 guest not already on tab 5.", lines)
        self.assertIn('Headcount: 5 guests (4 from tab 5 + 1 from tab 6).', lines)

    def test_dry_run_prints_the_headcount(self):
        with redirect_stdout(StringIO()) as out:
            self.assertEqual(cheddar_up.main(['--folder', self.work]), 0)
        self.assertIn('Headcount: 5 guests (4 from tab 5 + 1 from tab 6).', out.getvalue())

    def test_reading_guests_logs_names_per_tab(self):
        with redirect_stdout(StringIO()) as out:
            utils.readGuestList(self.export, self.requests)
        self.assertIn("Tab 5 '1-Member-Dinner-40': 4 purchase rows, 6 names listed, 4 guests not on an earlier tab.",
                      out.getvalue())

    def test_tab_without_attendee_columns_has_no_headcount(self):
        path = os.path.join(self.work, 'Odd-collection_summary.xlsx')
        from openpyxl import Workbook
        workbook = Workbook()
        for title in ('Payments', 'All Purchases', 'Item Summary', 'Tab 5', 'Tab 6'):
            workbook.create_sheet(title).append(['First Name', 'Last Name'])
        workbook.save(path)
        summary = cheddar_up.describe_export(path)
        self.assertIsNone(summary['headcount'])
        self.assertEqual(len(summary['problems']), 2)
        self.assertFalse(any(line.startswith('Headcount') for line in cheddar_up.summary_lines(summary)))


if __name__ == '__main__':
    unittest.main()
