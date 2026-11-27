# -*- coding: utf-8 -*-
"""Get the newest Cheddar Up collection summary into the program.

Cheddar Up has no documented way for a program to download the Collection
Summary spreadsheet. Its API key is only on the Team plan, its developer
docs are not public, and the web app's report download sits behind
Cloudflare. So the program does not log in to Cheddar Up. It opens
Cheddar Up in your normal browser, where you are already signed in, and
watches the Downloads folder. As soon as a new *collection_summary*.xlsx
finishes downloading, the file is copied into the "Cheddar Up exports"
folder beside the program and selected.

No password, token, or cookie is read or saved.

Check from a command prompt without opening the window:

    python cheddar_up.py              newest export in Downloads
    python cheddar_up.py --wait 600   wait up to 10 minutes for a new one
"""

import datetime
import fnmatch
import os
import shutil
import sys
import time
import webbrowser
import zipfile

from openpyxl import load_workbook

import utils

CHEDDAR_UP_URL = 'https://my.cheddarup.com/'
EXPORTS_FOLDER = 'Cheddar Up exports'
DEFAULT_COLLECTION_FILTER = 'TVNJC'
EXPORT_PATTERNS = ('*collection_summary*.xlsx', '*collection-summary*.xlsx', '*collection summary*.xlsx')
# Browsers write these while a download is still running.
PARTIAL_SUFFIXES = ('.crdownload', '.part', '.partial', '.download', '.tmp')
WAIT_SECONDS = 15 * 60

# Windows "Downloads" known folder. People sometimes move it to another drive.
_DOWNLOADS_GUID = '{374DE290-123F-4565-9164-39C4925E467B}'


def _windows_downloads():
    import ctypes
    from ctypes import wintypes

    class GUID(ctypes.Structure):
        _fields_ = [
            ('Data1', wintypes.DWORD),
            ('Data2', wintypes.WORD),
            ('Data3', wintypes.WORD),
            ('Data4', wintypes.BYTE * 8),
        ]

    guid = GUID()
    ctypes.oledll.ole32.CLSIDFromString(_DOWNLOADS_GUID, ctypes.byref(guid))
    path = ctypes.c_wchar_p()
    ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(path))
    try:
        return path.value
    finally:
        ctypes.windll.ole32.CoTaskMemFree(path)


def downloads_folder(saved=''):
    """The folder the browser saves to. A saved folder in configuration.json wins."""
    if isinstance(saved, str) and saved.strip() and os.path.isdir(saved.strip()):
        return saved.strip()
    if sys.platform == 'win32':
        try:
            folder = _windows_downloads()
            if folder and os.path.isdir(folder):
                return folder
        except Exception:
            pass
    return os.path.join(os.path.expanduser('~'), 'Downloads')


def is_export_name(name):
    lowered = name.casefold()
    if lowered.startswith('~$'):
        return False
    return any(fnmatch.fnmatch(lowered, pattern) for pattern in EXPORT_PATTERNS)


def _stamp(path):
    """When the file arrived.

    On Windows st_ctime is the creation time, so a copy made today counts as
    new even though it keeps the download's modified time.
    """
    info = os.stat(path)
    if sys.platform == 'win32':
        return max(info.st_mtime, info.st_ctime)
    return info.st_mtime


def _still_downloading(path):
    for suffix in PARTIAL_SUFFIXES:
        if os.path.exists(path + suffix):
            return True
    return False


def is_complete(path):
    """A finished xlsx is a zip file. Firefox leaves an empty file until the download ends."""
    try:
        if os.path.getsize(path) == 0:
            return False
    except OSError:
        return False
    if _still_downloading(path):
        return False
    try:
        return zipfile.is_zipfile(path)
    except OSError:
        return False


def list_exports(folder, since=None):
    """Finished collection summaries in folder, newest first."""
    found = []
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    for name in names:
        if not is_export_name(name):
            continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        try:
            stamp = _stamp(path)
        except OSError:
            continue
        if since is not None and stamp < since:
            continue
        if not is_complete(path):
            continue
        found.append((stamp, path))
    found.sort(key=lambda item: item[0], reverse=True)
    return [path for _stamp_value, path in found]


def matches_collection(path, collection_filter=DEFAULT_COLLECTION_FILTER):
    if not collection_filter:
        return True
    return collection_filter.casefold() in os.path.basename(path).casefold()


def find_latest_export(folder, collection_filter=DEFAULT_COLLECTION_FILTER, since=None):
    """Newest export for the event. Falls back to the newest export of any collection."""
    exports = list_exports(folder, since)
    for path in exports:
        if matches_collection(path, collection_filter):
            return path
    return exports[0] if exports else None


def is_newer(path, than):
    """True when path arrived after the file than, or than is missing."""
    if not than or not os.path.isfile(than):
        return True
    try:
        return _stamp(path) > _stamp(than) and os.path.getsize(path) > 0 and not _same_file(path, than)
    except OSError:
        return False


def _same_file(left, right):
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


def _same_contents(left, right):
    if os.path.getsize(left) != os.path.getsize(right):
        return False
    with open(left, 'rb') as a, open(right, 'rb') as b:
        while True:
            chunk = a.read(65536)
            if chunk != b.read(65536):
                return False
            if not chunk:
                return True


def import_export(path, exports_dir):
    """Copy the download into exports_dir and return the copy's path.

    The Cheddar Up file name is kept. If a different file already has that
    name, the copy gets the time it was saved added to its name.
    """
    os.makedirs(exports_dir, exist_ok=True)
    if _same_file(os.path.dirname(os.path.abspath(path)), exports_dir):
        return os.path.abspath(path)
    target = os.path.join(exports_dir, os.path.basename(path))
    if os.path.exists(target):
        if _same_contents(path, target):
            return target
        stem, ext = os.path.splitext(os.path.basename(path))
        stamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
        target = os.path.join(exports_dir, '{}_{}{}'.format(stem, stamp, ext))
    shutil.copy2(path, target)
    return target


def describe_export(path):
    """Summarize an export and list anything that would stop a seating chart.

    Returns a dict with collection, tabs, guest_tabs and problems.
    """
    summary = {'path': path, 'collection': '', 'tabs': [], 'guest_tabs': [], 'problems': []}
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
    except Exception as error:
        summary['problems'].append('Excel could not open it: {}'.format(error))
        return summary
    try:
        sheets = workbook.worksheets
        summary['tabs'] = [sheet.title for sheet in sheets]
        if sheets:
            for row in sheets[0].iter_rows(min_row=1, max_row=20, max_col=2, values_only=True):
                if row and utils._as_text(row[0]).casefold() == 'collection name':
                    summary['collection'] = utils._as_text(row[1] if len(row) > 1 else '')
                    break
        if len(sheets) < max(utils.GUEST_TAB_POSITIONS):
            summary['problems'].append(
                'It has {} tabs. Guests are read from tabs 5 and 6.'.format(len(sheets)))
            return summary
        for position in utils.GUEST_TAB_POSITIONS:
            sheet = sheets[position - 1]
            try:
                headers = utils._headers_of(sheet)
            except StopIteration:
                headers = []
            attendees = utils._attendee_columns(headers)
            rows = 0
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if any(utils._as_text(cell) for cell in row):
                    rows += 1
            summary['guest_tabs'].append({
                'position': position,
                'title': sheet.title,
                'rows': rows,
                'attendee_columns': len(attendees),
                'seatmate_columns': sum(1 for header in headers if utils._is_seatmate_header(header)),
            })
            if not attendees:
                summary['problems'].append('Tab {} ({}) has no attendee name columns.'.format(position, sheet.title))
    finally:
        workbook.close()
    return summary


def _count(number, noun):
    return '{} {}{}'.format(number, noun, '' if number == 1 else 's')


def summary_lines(summary):
    lines = ['Cheddar Up export: {}'.format(os.path.basename(summary['path']))]
    if summary['collection']:
        lines.append('Collection: {}'.format(summary['collection']))
    if summary['tabs']:
        lines.append('{} tabs.'.format(len(summary['tabs'])))
    for tab in summary['guest_tabs']:
        lines.append('Tab {} {!r}: {}, {}, {}.'.format(
            tab['position'], tab['title'],
            _count(tab['rows'], 'row'),
            _count(tab['attendee_columns'], 'attendee name column'),
            _count(tab['seatmate_columns'], 'seatmate request column')))
    for problem in summary['problems']:
        lines.append('WARNING: ' + problem)
    return lines


def open_cheddar_up(url=CHEDDAR_UP_URL, opener=None):
    """Open Cheddar Up in the normal browser, where the user is already signed in."""
    if opener is None:
        opener = webbrowser.open
    return opener(url or CHEDDAR_UP_URL)


class DownloadWatcher:
    """Wait for a new collection summary to finish downloading.

    Call check() every second or so. It returns the file's path once a new
    export has arrived and stopped growing, otherwise None.
    """

    def __init__(self, folder, since=None, timeout=WAIT_SECONDS, clock=time.time):
        self.folder = folder
        self.clock = clock
        self.started = clock() if since is None else since
        # A little slack for file systems that round timestamps.
        self.since = self.started - 2
        self.deadline = clock() + timeout
        self._sizes = {}

    def expired(self):
        return self.clock() > self.deadline

    def check(self):
        for path in list_exports(self.folder, self.since):
            try:
                size = os.path.getsize(path)
            except OSError:
                continue
            if self._sizes.get(path) == size:
                return path
            self._sizes[path] = size
        return None

    def wait(self, poll=1.0, sleep=time.sleep):
        while not self.expired():
            path = self.check()
            if path:
                return path
            sleep(poll)
        return None


def main(argv=None):
    import argparse

    parser = argparse.ArgumentParser(description='Find the newest Cheddar Up collection summary in Downloads.')
    parser.add_argument('--folder', default='', help='folder to look in (default: Downloads)')
    parser.add_argument('--filter', default=DEFAULT_COLLECTION_FILTER, help='text the file name should contain')
    parser.add_argument('--wait', type=int, default=0, help='seconds to wait for a new download')
    parser.add_argument('--open', action='store_true', help='open Cheddar Up in the browser first')
    args = parser.parse_args(argv)

    folder = downloads_folder(args.folder)
    print('Looking in {}'.format(folder))
    if args.open:
        open_cheddar_up()
    if args.wait:
        path = DownloadWatcher(folder, timeout=args.wait).wait()
    else:
        path = find_latest_export(folder, args.filter)
    if not path:
        print('No collection summary found.')
        return 1
    for line in summary_lines(describe_export(path)):
        print(line)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
