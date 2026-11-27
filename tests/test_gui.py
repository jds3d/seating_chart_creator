# -*- coding: utf-8 -*-
"""The window's Cheddar Up buttons, end to end. Needs a display (on Linux: xvfb-run)."""

import glob
import json
import os
import shutil
import tempfile
import time
import unittest
from unittest import mock

import synthetic

try:
    import tkinter as tk
    _root = tk.Tk()
    _root.destroy()
    HAVE_DISPLAY = True
except Exception:
    HAVE_DISPLAY = False

EXPORT_NAME = 'TVNJC_Dec_5_Test_Dinner-1234567-collection_summary.xlsx'


@unittest.skipUnless(HAVE_DISPLAY, 'Tkinter needs a display')
class WindowTests(unittest.TestCase):
    def setUp(self):
        import gui
        import review
        from test_pipeline import _load_creator

        self.gui = gui
        self.review = review
        self.work = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.work, True)
        self.previous = os.getcwd()
        os.chdir(self.work)
        self.addCleanup(os.chdir, self.previous)
        self.downloads = os.path.join(self.work, 'Downloads')
        os.makedirs(self.downloads)
        synthetic.write_requests(os.path.join(self.work, 'Requests.xlsx'))
        with open(os.path.join(self.work, gui.CONFIG_FILE), 'w', encoding='utf-8') as handle:
            json.dump({'downloads_folder': self.downloads}, handle)

        patches = [
            mock.patch.object(gui, 'APP_DIR', self.work),
            mock.patch.object(gui.cheddar_up, 'open_cheddar_up'),
            mock.patch.object(gui.messagebox, 'showinfo'),
            mock.patch.object(gui.messagebox, 'showerror'),
            mock.patch.object(gui.messagebox, 'showwarning'),
            mock.patch.object(gui.review, 'ask_review', side_effect=self._finish_review),
            mock.patch.object(gui.TableNumberDialog, 'ask', autospec=True, side_effect=self._save_numbers),
        ]
        self.mocks = [patch.start() for patch in patches]
        for patch in patches:
            self.addCleanup(patch.stop)
        self.opened = self.mocks[1]
        self.errors = self.mocks[3]

        self.root = tk.Tk()
        self.addCleanup(self.root.destroy)
        self.app = gui.App(self.root, _load_creator(), gui._load_module('table_rosters', 'Create Table Rosters From Seating Chart.py'))

    def _finish_review(self, parent, book, path):
        # The real review window, answered the way a person clicking Continue would.
        dialog = self.review.ReviewDialog(parent, book, path)
        dialog.after(200, dialog._continue)
        return dialog.ask()

    def _save_numbers(self, dialog):
        dialog.after(200, dialog.save)
        dialog.grab_set()
        dialog.wait_window()
        return dialog.result

    def pump(self, until, seconds=15):
        deadline = time.time() + seconds
        while time.time() < deadline:
            self.root.update()
            if until():
                return True
            time.sleep(0.05)
        return False

    def config(self):
        with open(os.path.join(self.work, self.gui.CONFIG_FILE), encoding='utf-8') as handle:
            return json.load(handle)

    def outputs(self):
        return sorted(os.path.basename(path) for path in glob.glob(os.path.join(self.work, 'output', '*.xlsx')))

    def test_waits_for_download_then_creates_the_chart(self):
        seen = []

        def yes(*_args, **_kwargs):
            seen.append(self.log())  # the export summary, before the chart run clears the log
            return True

        with mock.patch.object(self.gui.messagebox, 'askyesno', side_effect=yes) as asked:
            self.app.download_latest_export()
            self.opened.assert_called_once_with('https://my.cheddarup.com/')
            self.assertEqual(self.app.download_button.cget('text'), 'Stop waiting')
            self.pump(lambda: False, seconds=1.5)
            self.assertEqual(self.app.export_var.get(), '')

            synthetic.write_export(os.path.join(self.downloads, EXPORT_NAME))
            self.assertTrue(self.pump(lambda: len(self.outputs()) == 3, seconds=20), self.log())
        asked.assert_called_once()
        saved = os.path.join(self.work, 'Cheddar Up exports', EXPORT_NAME)
        self.assertEqual(self.app.export_var.get(), saved)
        self.assertEqual(self.config()['cheddar_up_export'], saved)
        self.assertEqual(self.config()['downloads_folder'], self.downloads)
        self.assertTrue(os.path.isfile(saved))
        self.assertTrue(os.path.isfile(os.path.join(self.work, 'Corrections.xlsx')))
        self.assertTrue(os.path.isfile(os.path.join(self.work, 'guest_data.p')))
        names = self.outputs()
        self.assertTrue(names[0].startswith('Seating Chart_'))
        self.assertTrue(names[1].startswith('Table Rosters_emails_'))
        self.assertTrue(names[2].startswith('Table Rosters_no emails_'))
        self.assertEqual(self.app.download_button.cget('text'), 'Download latest from Cheddar Up')
        self.errors.assert_not_called()
        self.assertIn('Collection: TVNJC Test Dinner', seen[0])
        self.assertIn("Tab 5 '1-Member-Dinner-40': 9 rows, 2 attendee name columns, 1 seatmate request column.", seen[0])

    def test_uses_an_export_already_in_downloads(self):
        synthetic.write_export(os.path.join(self.downloads, EXPORT_NAME))
        with mock.patch.object(self.gui.messagebox, 'askyesnocancel', return_value=True), \
                mock.patch.object(self.gui.messagebox, 'askyesno', return_value=False):
            self.app.download_latest_export()
        self.opened.assert_not_called()
        saved = os.path.join(self.work, 'Cheddar Up exports', EXPORT_NAME)
        self.assertEqual(self.app.export_var.get(), saved)
        self.assertEqual(self.config()['cheddar_up_export'], saved)
        self.assertEqual(self.outputs(), [])

    def test_cancel_and_stop_waiting(self):
        synthetic.write_export(os.path.join(self.downloads, EXPORT_NAME))
        with mock.patch.object(self.gui.messagebox, 'askyesnocancel', return_value=None):
            self.app.download_latest_export()
        self.assertEqual(self.app.export_var.get(), '')
        self.opened.assert_not_called()
        with mock.patch.object(self.gui.messagebox, 'askyesnocancel', return_value=False):
            self.app.download_latest_export()
        self.opened.assert_called_once()
        self.app.download_latest_export()
        self.assertIsNone(self.app._watcher)
        self.assertEqual(self.app.download_button.cget('text'), 'Download latest from Cheddar Up')

    def test_browse_still_works(self):
        path = synthetic.write_export(os.path.join(self.work, 'picked-collection_summary.xlsx'))
        with mock.patch.object(self.gui.filedialog, 'askopenfilename', return_value=path):
            self.app.browse_export()
        self.assertEqual(self.app.export_var.get(), path)
        self.assertEqual(self.config()['cheddar_up_export'], path)

    def test_create_from_a_browsed_export(self):
        path = synthetic.write_export(os.path.join(self.work, 'picked-collection_summary.xlsx'))
        self.app.export_var.set(path)
        self.app.create_seating_chart()
        self.assertEqual(len(self.outputs()), 3, self.log())
        self.errors.assert_not_called()

    def log(self):
        return self.app.log_text.get('1.0', tk.END)


if __name__ == '__main__':
    unittest.main()
