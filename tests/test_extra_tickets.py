# -*- coding: utf-8 -*-
"""Missing attendees for orders with more tickets than names, and the Names sheet for full names."""

import os
import shutil
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO

from openpyxl import Workbook

import synthetic

import corrections
import utils

# Three ticket rows, one order: 3 tickets, but only two people named (7th value is the quantity).
MEMBERS = [
    ('Ken', 'Abel', 'ken@example.com', 'available', [('Ken Abel', 'Turkey'), ('Nan Abel', 'Turkey')], '', 1),
    ('Gail', 'Gray', 'gail@example.com', 'available', [('Gail Gray', 'Ham'), ('Drbra Carosells', 'Ham')], 'RICKI RODKIN', 2),
    ('Eve', 'Mull', 'eve@example.com', 'available', [], '', 2),
    ('Hal', 'Hart', 'hal@example.com', 'available', [('Hal Hart', 'Ham'), ('Ida Hart', 'Ham')], '', 2),
    ('Hal', 'Hart', 'hal@example.com', 'refunded', [('Hal Hart', 'Ham'), ('Ida Hart', 'Ham')], '', 1),
]
GUESTS = [
    ('Ken', 'Abel', 'ken@example.com', 'available', [('Ken Abel', 'Turkey'), ('Nan Abel', 'Turkey')], '', 1),
    ('Ken', 'Abel', 'ken@example.com', 'available', [('Ken Abel', 'Turkey'), ('Nan Abel', 'Turkey')], '', 1),
    ('Deb', 'Kane', 'deb@example.com', 'available', [('Deb Kane', 'Ham')], 'Debra Carosella', 1),
]


class ExtraTicketTests(unittest.TestCase):
    def setUp(self):
        self.work = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.work, True)
        self.export = synthetic.write_export(
            os.path.join(self.work, 'TVNJC_Extra-collection_summary.xlsx'), member_rows=MEMBERS, guest_rows=GUESTS)
        self.requests = synthetic.write_requests(os.path.join(self.work, 'Requests.xlsx'))

    def read(self):
        with redirect_stdout(StringIO()):
            return utils.readGuestList(self.export, self.requests)

    def corrections_file(self, **sheets):
        path = os.path.join(self.work, 'Corrections.xlsx')
        workbook = Workbook()
        workbook.active.title = 'Read me'
        for key, rows in sheets.items():
            title, headers = corrections._SHEETS[key]
            sheet = workbook.create_sheet(title)
            sheet.append(headers)
            for row in rows:
                sheet.append(list(row))
        workbook.save(path)
        return path

    def run_corrections(self, **sheets):
        book = self.read()
        loaded = corrections.load_corrections(self.corrections_file(**sheets)) if sheets else {
            key: [] for key in corrections._SHEETS}
        state = corrections.build_state(book, loaded)
        corrections.refresh_open_seatmates(book, state)
        corrections.apply_all(book, state)
        book.finish()
        return book, state

    def test_orders_are_summed_across_tabs(self):
        orders = self.read().short_orders()
        self.assertEqual([(order['payer'], order['tickets'], order['extra']) for order in orders],
                         [('Ken Abel', 3, 1)])
        self.assertEqual(orders[0]['names'], ['Ken Abel', 'Nan Abel'])

    def test_blank_and_refunded_orders_are_not_short_orders(self):
        book = self.read()
        # Eve names nobody. Hal's refunded row is ignored, so 2 tickets and 2 names.
        self.assertEqual([order['payer'] for order in book.short_orders()], ['Ken Abel'])
        self.assertEqual([item['payer'] for item in book.missing_attendees], ['Eve Mull'])

    def test_short_order_is_asked_about(self):
        book = self.read()
        state = corrections.build_state(book, {key: [] for key in corrections._SHEETS})
        item = state['missing']['Ken Abel']
        self.assertEqual(item['kind'], 'short')
        self.assertFalse(item['decided'])
        self.assertEqual(item['people'], [{'name': '', 'choice': ''}])
        self.assertEqual(item['named'], ['Ken Abel', 'Nan Abel'])
        self.assertEqual(state['missing']['Eve Mull']['kind'], 'blank')
        self.assertIn('Missing attendees', [title for _key, title, _pending in corrections.section_defs(state)])
        self.assertTrue(any('Ken Abel paid for 3 tickets but names only 2' in warning for warning in book.warnings))

    def test_sheet_row_adds_the_extra_person_to_the_party(self):
        book, state = self.run_corrections(missing=[('Ken Abel', 'Rita Rowe', 'Tuna')])
        self.assertIn('Rita Rowe', book.guests)
        self.assertEqual(book.extra['Rita Rowe'], 'Tuna')
        self.assertEqual(book.guests['Rita Rowe'], {'Ken Abel', 'Nan Abel', 'Rita Rowe'})
        self.assertIn('Rita Rowe', book.guests['Ken Abel'])
        self.assertIn(['Ken Abel', 'Nan Abel', 'Rita Rowe'], book.parties)
        # The payer stays seated.
        self.assertIn('Ken Abel', book.guests)
        self.assertEqual(book.emails['Rita Rowe'], 'ken@example.com')

    def test_seatmate_request_names_the_new_person(self):
        book, state = self.run_corrections(
            missing=[('Ken Abel', 'Ricki Rodkin', 'Tuna')], names=[('Drbra Carosells', 'Debra Carosella')])
        self.assertIn('Ricki Rodkin', book.guests)
        self.assertIn('Ricki Rodkin', book.guests['Gail Gray'])
        self.assertNotIn('Finish the seatmate request "RICKI RODKIN".', corrections.validate(book, state))
        self.assertIn(('RICKI RODKIN', 'Ricki Rodkin', False), book.applied_seatmates)

    def test_blank_order_still_adds_the_people_as_a_party(self):
        book, state = self.run_corrections(missing=[('Eve Mull', 'Eileen Mullen', 'Turkey'), ('Eve Mull', 'Joe Boyer', 'Beef')])
        self.assertEqual(book.guests['Eileen Mullen'], {'Eileen Mullen', 'Joe Boyer'})
        self.assertEqual(book.extra['Joe Boyer'], 'Beef')
        self.assertEqual(state['missing']['Eve Mull']['kind'], 'blank')

    def test_unclaimed_meal_on_a_short_order_goes_to_the_new_person(self):
        member_rows = [('Ken', 'Abel', 'ken@example.com', 'available', [('Ken Abel', 'Turkey'), ('', 'Tuna')], '', 2)]
        export = synthetic.write_export(
            os.path.join(self.work, 'Unclaimed-collection_summary.xlsx'), member_rows=member_rows, guest_rows=[])
        with redirect_stdout(StringIO()):
            book = utils.readGuestList(export, self.requests)
        state = corrections.build_state(book, {key: [] for key in corrections._SHEETS})
        self.assertEqual(state['unclaimed'], {})
        self.assertEqual(state['missing']['Ken Abel']['people'], [{'name': '', 'choice': 'Tuna'}])

    def test_names_sheet_renames_a_full_name(self):
        book, state = self.run_corrections(names=[('drbra  carosells.', 'Debra Carosella')])
        self.assertIn('Debra Carosella', book.guests)
        self.assertNotIn('Drbra Carosells', book.guests)
        # Gail's order is still a party with her, and the request names her directly.
        self.assertIn('Gail Gray', book.guests['Debra Carosella'])
        self.assertIn('Debra Carosella', book.guests['Deb Kane'])
        self.assertNotIn('Finish the seatmate request "Debra Carosella".', corrections.validate(book, state))
        self.assertIn(('Debra Carosella', 'Debra Carosella', False), book.applied_seatmates)

    def test_names_sheet_row_for_an_unknown_name_is_ignored(self):
        book, state = self.run_corrections(names=[('Nobody Here', 'Someone Else')])
        self.assertNotIn('Someone Else', book.guests)
        self.assertNotIn('Nobody Here', state['names'])

    def test_saving_keeps_both_kinds_of_rows(self):
        book = self.read()
        loaded = corrections.load_corrections(self.corrections_file(
            missing=[('Ken Abel', 'Rita Rowe', 'Tuna')], names=[('Drbra Carosells', 'Debra Carosella')]))
        state = corrections.build_state(book, loaded)
        path = os.path.join(self.work, 'Saved.xlsx')
        corrections.save_corrections(path, state, loaded)
        saved = corrections.load_corrections(path)
        self.assertEqual([(row['Payer'], row['Person'], row['Choice']) for row in saved['missing']],
                         [('Ken Abel', 'Rita Rowe', 'Tuna')])
        self.assertEqual([(row['As written'], row['Correct name']) for row in saved['names']],
                         [('Drbra Carosells', 'Debra Carosella')])


if __name__ == '__main__':
    unittest.main()
