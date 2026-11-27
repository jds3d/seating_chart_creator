# -*- coding: utf-8 -*-
"""The existing seating steps, run on a made-up export picked up from a fake Downloads folder."""

import datetime
import glob
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


def _load_creator():
    path = os.path.join(synthetic.ROOT, 'Seating Chart Creator.py')
    spec = importlib.util.spec_from_file_location('seating_creator_test', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.work = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.work, True)
        self.previous = os.getcwd()
        os.chdir(self.work)
        self.addCleanup(os.chdir, self.previous)
        self.downloads = os.path.join(self.work, 'Downloads')
        os.makedirs(self.downloads)
        synthetic.write_export(os.path.join(
            self.downloads, 'TVNJC_Dec_5_Test_Dinner-1234567-collection_summary.xlsx'), extra_tabs=2)
        self.requests = synthetic.write_requests(os.path.join(self.work, utils.REQUESTS_WORKBOOK))
        self.creator = _load_creator()

    def test_download_then_seat(self):
        found = cheddar_up.find_latest_export(self.downloads)
        export = cheddar_up.import_export(found, os.path.join(self.work, cheddar_up.EXPORTS_FOLDER))
        self.assertEqual(cheddar_up.describe_export(export)['problems'], [])

        with redirect_stdout(StringIO()):
            book = utils.readGuestList(export, self.requests)
        self.assertEqual(set(book.guests), synthetic.guest_names())
        self.assertNotIn('Kurt Garcia', book.guests)

        state = corrections.build_state(book, {key: [] for key in corrections._SHEETS})
        self.assertEqual(corrections.validate(book, state), [])
        corrections.apply_all(book, state)
        book.finish()

        random.seed(5)
        with redirect_stdout(StringIO()):
            tables = self.creator.generateSeatingChart(book.guests, book.anti, confirm_large_group=lambda *a: True)

        seated = [guest for table in tables for guest in table]
        self.assertEqual(sorted(seated), sorted(synthetic.guest_names()))
        self.assertTrue(all(len(table) <= self.creator.normalTableMaxSize for table in tables))

        def table_of(name):
            return next(index for index, table in enumerate(tables) if name in table)

        # Same purchase, the seatmate column, and the Requests tab keep people together.
        self.assertEqual(table_of('Alma Archer'), table_of('Boris Archer'))
        self.assertEqual(table_of('Alma Archer'), table_of('Cyrus Bell'))
        self.assertEqual(table_of('Edgar Cohen'), table_of('Hazel Eaton'))
        # The Antirequests tab keeps people apart.
        self.assertNotEqual(table_of('Fiona Dunn'), table_of('Ivan Foster'))

        timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
        with redirect_stdout(StringIO()):
            utils.writeSeatingChart(tables, book.extra, timestamp)
            utils.writeTables(tables, book.emails, book.polls, timestamp, writeEmails=True)
            utils.writeTables(tables, book.emails, book.polls, timestamp, writeEmails=False, extraColumn=book.extra)
        written = sorted(os.path.basename(path) for path in glob.glob(os.path.join('output', '*.xlsx')))
        self.assertEqual(written, [
            'Seating Chart_{}.xlsx'.format(timestamp),
            'Table Rosters_emails_{}.xlsx'.format(timestamp),
            'Table Rosters_no emails_{}.xlsx'.format(timestamp),
        ])

    def test_short_export_still_stops_like_before(self):
        from openpyxl import Workbook
        path = os.path.join(self.downloads, 'TVNJC_Short-collection_summary.xlsx')
        workbook = Workbook()
        for title in ('Payments', 'All Purchases', 'Item Summary'):
            workbook.create_sheet(title)
        workbook.save(path)
        with redirect_stdout(StringIO()):
            with self.assertRaises(SystemExit):
                utils.readGuestList(path, self.requests)


if __name__ == '__main__':
    unittest.main()
