# -*- coding: utf-8 -*-
"""Download step, with fake files in a temporary Downloads folder. No Cheddar Up account needed."""

import os
import shutil
import tempfile
import time
import unittest
from unittest import mock

import synthetic  # noqa: F401  (puts the project folder on sys.path)

import cheddar_up

EXPORT_NAME = 'TVNJC_Dec_5_Test_Dinner-1234567-collection_summary.xlsx'


def _touch(path, when, data=b''):
    with open(path, 'wb') as handle:
        handle.write(data)
    os.utime(path, (when, when))
    return path


class DownloadsFolderTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.folder, True)

    def export(self, name, when=None, **kwargs):
        path = os.path.join(self.folder, name)
        synthetic.write_export(path, **kwargs)
        if when is not None:
            os.utime(path, (when, when))
        return path

    def test_export_names(self):
        self.assertTrue(cheddar_up.is_export_name(EXPORT_NAME))
        self.assertTrue(cheddar_up.is_export_name('TVNJC-collection_summary (1).xlsx'))
        self.assertTrue(cheddar_up.is_export_name('Gala-Collection-Summary.XLSX'))
        self.assertFalse(cheddar_up.is_export_name('~$TVNJC-collection_summary.xlsx'))
        self.assertFalse(cheddar_up.is_export_name('TVNJC-collection_summary.xlsx.crdownload'))
        self.assertFalse(cheddar_up.is_export_name('TVNJC-item_summary.xlsx'))
        self.assertFalse(cheddar_up.is_export_name('Requests.xlsx'))

    def test_latest_prefers_the_event(self):
        # Files are compared by arrival time, so set it in the past and leave newer ones newer.
        now = time.time()
        older = self.export('TVNJC_Old-collection_summary.xlsx', now - 3600)
        other = self.export('Soccer_Club-collection_summary.xlsx')
        newest = self.export('TVNJC_New-collection_summary.xlsx', now - 60)
        with mock.patch.object(cheddar_up, '_stamp', side_effect=lambda path: os.stat(path).st_mtime):
            self.assertEqual(cheddar_up.list_exports(self.folder), [other, newest, older])
            self.assertEqual(cheddar_up.find_latest_export(self.folder, 'TVNJC'), newest)
            self.assertEqual(cheddar_up.find_latest_export(self.folder, ''), other)
            self.assertEqual(cheddar_up.list_exports(self.folder, since=now - 120), [other, newest])

    def test_latest_falls_back_to_any_collection(self):
        only = self.export('Soccer_Club-collection_summary.xlsx')
        self.assertEqual(cheddar_up.find_latest_export(self.folder, 'TVNJC'), only)

    def test_unfinished_downloads_are_ignored(self):
        now = time.time()
        _touch(os.path.join(self.folder, 'TVNJC_A-collection_summary.xlsx'), now)  # Firefox: empty until done
        partial = self.export('TVNJC_B-collection_summary.xlsx')
        _touch(partial + '.part', now)
        _touch(os.path.join(self.folder, 'TVNJC_C-collection_summary.xlsx'), now, b'not a zip')
        self.assertIsNone(cheddar_up.find_latest_export(self.folder))
        os.remove(partial + '.part')
        self.assertEqual(cheddar_up.find_latest_export(self.folder), partial)

    def test_missing_folder(self):
        self.assertEqual(cheddar_up.list_exports(os.path.join(self.folder, 'nope')), [])
        self.assertIsNone(cheddar_up.find_latest_export(os.path.join(self.folder, 'nope')))

    def test_saved_downloads_folder_wins(self):
        self.assertEqual(cheddar_up.downloads_folder(self.folder), self.folder)
        fallback = cheddar_up.downloads_folder(os.path.join(self.folder, 'missing'))
        self.assertTrue(fallback.endswith('Downloads'))

    def test_is_newer(self):
        now = time.time()
        old = self.export('TVNJC_Old-collection_summary.xlsx', now - 600)
        new = self.export('TVNJC_New-collection_summary.xlsx')
        self.assertTrue(cheddar_up.is_newer(new, ''))
        self.assertTrue(cheddar_up.is_newer(new, os.path.join(self.folder, 'gone.xlsx')))
        self.assertFalse(cheddar_up.is_newer(new, new))
        with mock.patch.object(cheddar_up, '_stamp', side_effect=lambda path: os.stat(path).st_mtime):
            self.assertTrue(cheddar_up.is_newer(new, old))
            self.assertFalse(cheddar_up.is_newer(old, new))

    def test_import_copies_once(self):
        source = self.export(EXPORT_NAME)
        exports = os.path.join(self.folder, 'Cheddar Up exports')
        first = cheddar_up.import_export(source, exports)
        self.assertEqual(os.path.basename(first), EXPORT_NAME)
        self.assertTrue(os.path.isfile(first))
        self.assertEqual(cheddar_up.import_export(source, exports), first)
        self.assertEqual(cheddar_up.import_export(first, exports), os.path.abspath(first))

    def test_import_keeps_a_different_file_with_the_same_name(self):
        exports = os.path.join(self.folder, 'Cheddar Up exports')
        first = cheddar_up.import_export(self.export(EXPORT_NAME), exports)
        changed = self.export(EXPORT_NAME, collection='TVNJC Changed')
        second = cheddar_up.import_export(changed, exports)
        self.assertNotEqual(first, second)
        self.assertTrue(os.path.basename(second).startswith(EXPORT_NAME[:-5] + '_'))
        self.assertTrue(os.path.isfile(first))

    def test_describe_export(self):
        summary = cheddar_up.describe_export(self.export(EXPORT_NAME))
        self.assertEqual(summary['problems'], [])
        self.assertEqual(summary['collection'], 'TVNJC Test Dinner')
        self.assertEqual(len(summary['tabs']), 7)
        tabs = summary['guest_tabs']
        self.assertEqual([tab['position'] for tab in tabs], [5, 6])
        self.assertEqual(tabs[0]['rows'], len(synthetic.MEMBER_ROWS))
        self.assertEqual(tabs[0]['attendee_columns'], 2)
        self.assertEqual(tabs[0]['seatmate_columns'], 1)
        lines = cheddar_up.summary_lines(summary)
        self.assertIn('Collection: TVNJC Test Dinner', lines)

    def test_describe_short_export(self):
        path = os.path.join(self.folder, 'Short-collection_summary.xlsx')
        from openpyxl import Workbook
        workbook = Workbook()
        workbook.create_sheet('Payments')
        workbook.save(path)
        summary = cheddar_up.describe_export(path)
        self.assertEqual(len(summary['problems']), 1)
        self.assertIn('tabs 5 and 6', summary['problems'][0])

    def test_describe_not_excel(self):
        path = _touch(os.path.join(self.folder, 'broken-collection_summary.xlsx'), time.time(), b'xx')
        self.assertTrue(cheddar_up.describe_export(path)['problems'])

    def test_open_cheddar_up_uses_the_browser(self):
        opener = mock.Mock(return_value=True)
        cheddar_up.open_cheddar_up('', opener=opener)
        opener.assert_called_once_with(cheddar_up.CHEDDAR_UP_URL)
        with mock.patch('webbrowser.open') as browser:
            cheddar_up.open_cheddar_up('https://my.cheddarup.com/collection/1/manage')
        browser.assert_called_once_with('https://my.cheddarup.com/collection/1/manage')


class WatcherTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.folder, True)

    def test_old_files_are_not_picked_up(self):
        old = os.path.join(self.folder, EXPORT_NAME)
        synthetic.write_export(old)
        past = time.time() - 3600
        os.utime(old, (past, past))
        watcher = cheddar_up.DownloadWatcher(self.folder, timeout=60)
        with mock.patch.object(cheddar_up, '_stamp', side_effect=lambda path: os.stat(path).st_mtime):
            self.assertIsNone(watcher.check())
            self.assertIsNone(watcher.check())

    def test_new_download_is_returned_once_it_stops_growing(self):
        watcher = cheddar_up.DownloadWatcher(self.folder, timeout=60)
        self.assertIsNone(watcher.check())
        partial = os.path.join(self.folder, EXPORT_NAME + '.crdownload')
        _touch(partial, time.time(), b'abc')
        self.assertIsNone(watcher.check())
        path = os.path.join(self.folder, EXPORT_NAME)
        synthetic.write_export(path)
        os.remove(partial)
        self.assertIsNone(watcher.check())  # first sight: size noted
        self.assertEqual(watcher.check(), path)  # same size again: finished

    def test_wait_times_out(self):
        clock = [1000.0]
        watcher = cheddar_up.DownloadWatcher(self.folder, timeout=5, clock=lambda: clock[0])

        def sleep(seconds):
            clock[0] += seconds

        self.assertIsNone(watcher.wait(poll=1, sleep=sleep))
        self.assertTrue(watcher.expired())

    def test_wait_returns_a_download_made_by_another_thread(self):
        import threading
        path = os.path.join(self.folder, EXPORT_NAME)
        timer = threading.Timer(0.3, synthetic.write_export, args=(path,))
        timer.start()
        self.addCleanup(timer.cancel)
        self.assertEqual(cheddar_up.DownloadWatcher(self.folder, timeout=10).wait(poll=0.2), path)

    def test_command_line_dry_run(self):
        synthetic.write_export(os.path.join(self.folder, EXPORT_NAME))
        with mock.patch('builtins.print') as printed:
            self.assertEqual(cheddar_up.main(['--folder', self.folder]), 0)
        text = '\n'.join(str(call.args[0]) for call in printed.call_args_list)
        self.assertIn(EXPORT_NAME, text)
        self.assertIn("Tab 5 '1-Member-Dinner-40'", text)
        empty = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, empty, True)
        with mock.patch('builtins.print'):
            self.assertEqual(cheddar_up.main(['--folder', empty]), 1)


if __name__ == '__main__':
    unittest.main()
