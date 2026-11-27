# -*- coding: utf-8 -*-
"""Window for creating a seating chart and table rosters."""

import contextlib
import datetime
import importlib.util
import json
import os
import pickle
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cheddar_up
import corrections
import review
import utils


def _frozen():
    return getattr(sys, 'frozen', False)


def code_dir():
    """Folder that holds the program files. Inside an exe, that is the unpack directory."""
    if _frozen():
        return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(sys.executable)))
    return os.path.dirname(os.path.abspath(__file__))


def work_dir():
    """Folder mom works in: Requests.xlsx, output, and error.log sit next to the exe."""
    if _frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = work_dir()
CONFIG_FILE = 'configuration.json'


def _config_path():
    return os.path.join(APP_DIR, CONFIG_FILE)


def _read_config():
    try:
        with open(_config_path(), encoding='utf-8') as handle:
            data = json.load(handle)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def _load_cheddar_up_export():
    saved = _read_config().get('cheddar_up_export', '')
    return saved.strip() if isinstance(saved, str) else ''


def _load_setting(key, default=''):
    saved = _read_config().get(key, default)
    return saved.strip() if isinstance(saved, str) else default


def _save_cheddar_up_export(path):
    data = _read_config()
    data['cheddar_up_export'] = path
    with open(_config_path(), 'w', encoding='utf-8') as handle:
        json.dump(data, handle, indent=2)
        handle.write('\n')


def _ensure_stdio():
    """A windowed exe has no console, so stdout and stderr start out as None."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, 'w', encoding='utf-8')
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w', encoding='utf-8')


def _install_excepthook():
    def handle(exc_type, exc, tb):
        log_path = os.path.join(work_dir(), 'error.log')
        try:
            with open(log_path, 'a', encoding='utf-8') as log_file:
                log_file.write('\n')
                traceback.print_exception(exc_type, exc, tb, file=log_file)
        except Exception:
            log_path = 'error.log'
        try:
            if tk._default_root is not None:
                messagebox.showerror(
                    'Seating Chart Creator',
                    'Something went wrong. Details were saved to:\n{}'.format(log_path),
                    parent=tk._default_root,
                )
        except Exception:
            pass

    sys.excepthook = handle


_ensure_stdio()


def _load_module(module_name, filename):
    path = os.path.join(code_dir(), filename)
    if not os.path.isfile(path):
        raise FileNotFoundError('Missing program file: {}'.format(path))
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _enable_dpi():
    if sys.platform != 'win32':
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class _LogStream:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        if not text:
            return
        self.widget.configure(state=tk.NORMAL)
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state=tk.DISABLED)
        self.widget.update_idletasks()

    def flush(self):
        pass


@contextlib.contextmanager
def _capture_stdout(widget):
    previous = sys.stdout
    sys.stdout = _LogStream(widget)
    try:
        yield
    finally:
        sys.stdout = previous


class TableNumberDialog(tk.Toplevel):
    """Ask for each table's number. Tables are kept from lowest number to highest."""

    def __init__(self, master, tables):
        super().__init__(master)
        self.title('Table numbers')
        self.result = None
        self.tables = list(tables)
        self.vars = []
        self.minsize(560, 420)
        self.geometry('680x560')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        highest = len(self.tables) + 3
        intro = ttk.Label(
            self,
            wraplength=620,
            text=(
                'Give each table a different number from 1 to {}. '
                'Tables are saved from the lowest number to the highest.'
            ).format(highest),
        )
        intro.grid(row=0, column=0, sticky='ew', padx=16, pady=(16, 8))

        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=canvas.yview)
        body = ttk.Frame(canvas, padding=(0, 0, 12, 0))
        window_id = canvas.create_window((0, 0), window=body, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=1, column=0, sticky='nsew', padx=(16, 0), pady=4)
        scrollbar.grid(row=1, column=1, sticky='ns', padx=(0, 8), pady=4)

        def _fit_width(event):
            canvas.itemconfigure(window_id, width=event.width)

        def _fit_scroll(_event):
            canvas.configure(scrollregion=canvas.bbox('all'))

        canvas.bind('<Configure>', _fit_width)
        body.bind('<Configure>', _fit_scroll)

        for index, table in enumerate(self.tables, 1):
            names = ', '.join(sorted(utils.formatNameWithoutFamilyName(guest) for guest in table))
            block = ttk.Frame(body, padding=(0, 6))
            block.pack(fill=tk.X)
            ttk.Label(block, text='{} people'.format(len(table))).grid(row=0, column=0, sticky='w')
            number = tk.StringVar(value=str(index))
            self.vars.append(number)
            ttk.Entry(block, textvariable=number, width=6).grid(row=0, column=1, padx=(12, 0))
            ttk.Label(block, text=names, wraplength=520).grid(row=1, column=0, columnspan=2, sticky='w', pady=(2, 0))

        buttons = ttk.Frame(self, padding=(16, 8, 16, 16))
        buttons.grid(row=2, column=0, columnspan=2, sticky='e')
        ttk.Button(buttons, text='Cancel', command=self.cancel).pack(side=tk.RIGHT)
        ttk.Button(buttons, text='Save with these numbers', command=self.save).pack(side=tk.RIGHT, padx=(0, 8))

        self.protocol('WM_DELETE_WINDOW', self.cancel)
        self.bind('<MouseWheel>', lambda event: canvas.yview_scroll(int(-event.delta / 120), 'units'))
        self.bind('<Escape>', lambda _event: self.cancel())
        self.after(50, lambda: self.vars and self.focus_set())
        utils.keep_minmax(self, master)

    def ask(self):
        self.grab_set()
        self.wait_window()
        return self.result

    def cancel(self):
        self.result = None
        self.destroy()

    def save(self):
        highest = len(self.tables) + 3
        numbers = []
        for var in self.vars:
            text = var.get().strip()
            if not text.isdigit() or not 1 <= int(text) <= highest:
                messagebox.showerror(
                    'Table numbers',
                    'Use a whole number from 1 to {} for every table.'.format(highest),
                    parent=self,
                )
                return
            numbers.append(int(text))
        if len(set(numbers)) != len(numbers):
            messagebox.showerror('Table numbers', 'Each table needs its own number.', parent=self)
            return
        self.result = [table for _number, table in sorted(zip(numbers, self.tables), key=lambda item: item[0])]
        self.destroy()


class App(ttk.Frame):
    def __init__(self, master, creator, rosters):
        super().__init__(master, padding=16)
        self.creator = creator
        self.rosters = rosters
        self.export_var = tk.StringVar(value=_load_cheddar_up_export())
        self.table_size_var = tk.StringVar(value=str(creator.normalTableMaxSize))
        self.group_limit_var = tk.StringVar(value=str(creator.singleGroupMaxSize))
        saved_export = self.export_var.get().strip()
        self.status_var = tk.StringVar(
            value=os.path.basename(saved_export) if saved_export else 'Choose a Cheddar Up export and a requests workbook.'
        )
        self._busy = False
        self._watcher = None
        self._watch_job = None

        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.grid(sticky='nsew')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky='ew')
        ttk.Label(header, text='Seating Chart Creator', style='Title.TLabel').grid(row=0, column=0, sticky='w')
        ttk.Label(
            header,
            style='Hint.TLabel',
            text='Reads tabs 5 and 6 of a Cheddar Up export, plus a requests workbook.',
        ).grid(row=1, column=0, sticky='w', pady=(2, 0))

        default_requests = os.path.join(APP_DIR, utils.REQUESTS_WORKBOOK)
        self.requests_var = tk.StringVar(value=default_requests if os.path.isfile(default_requests) else '')

        config = ttk.LabelFrame(self, text='Configuration', padding=12)
        config.grid(row=1, column=0, sticky='ew', pady=(16, 0))
        config.columnconfigure(1, weight=1)
        ttk.Label(config, text='Cheddar Up export').grid(row=0, column=0, sticky='w')
        self.export_entry = ttk.Entry(config, textvariable=self.export_var)
        self.export_entry.grid(row=0, column=1, sticky='ew', padx=8)
        ttk.Button(config, text='Browse…', command=self.browse_export).grid(row=0, column=2)
        cheddar = ttk.Frame(config)
        cheddar.grid(row=1, column=1, columnspan=2, sticky='w', padx=8, pady=(6, 0))
        self.download_button = ttk.Button(
            cheddar, text='Download latest from Cheddar Up', command=self.download_latest_export)
        self.download_button.pack(side=tk.LEFT)
        ttk.Button(cheddar, text='Open Cheddar Up', command=self.open_cheddar_up).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(config, text='Requests').grid(row=2, column=0, sticky='w', pady=(8, 0))
        self.requests_entry = ttk.Entry(config, textvariable=self.requests_var)
        self.requests_entry.grid(row=2, column=1, sticky='ew', padx=8, pady=(8, 0))
        ttk.Button(config, text='Browse…', command=self.browse_requests).grid(row=2, column=2, pady=(8, 0))
        self.export_entry.xview_moveto(1)
        self.requests_entry.xview_moveto(1)

        settings = ttk.LabelFrame(self, text='Table size', padding=12)
        settings.grid(row=2, column=0, sticky='ew', pady=(12, 0))
        ttk.Label(settings, text='People per table').grid(row=0, column=0, sticky='w')
        ttk.Spinbox(settings, from_=1, to=20, width=5, textvariable=self.table_size_var).grid(row=0, column=1, padx=(8, 24))
        ttk.Label(settings, text='Ask when a group is larger than').grid(row=0, column=2, sticky='w')
        ttk.Spinbox(settings, from_=1, to=30, width=5, textvariable=self.group_limit_var).grid(row=0, column=3, padx=(8, 0))
        ttk.Label(
            settings,
            style='Hint.TLabel',
            text='A group over the table size still gets its own table when you allow it.',
        ).grid(row=1, column=0, columnspan=4, sticky='w', pady=(8, 0))

        actions = ttk.Frame(self)
        actions.grid(row=3, column=0, sticky='nsew', pady=(12, 0))
        actions.columnconfigure(0, weight=1)
        actions.rowconfigure(1, weight=1)

        buttons = ttk.Frame(actions)
        buttons.grid(row=0, column=0, sticky='ew')
        self.create_button = ttk.Button(buttons, text='Create seating chart', command=self.create_seating_chart)
        self.create_button.pack(side=tk.LEFT)
        self.roster_button = ttk.Button(buttons, text='Rosters from a seating chart', command=self.create_rosters)
        self.roster_button.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text='Open output folder', command=self.open_output_folder).pack(side=tk.RIGHT)

        log_frame = ttk.LabelFrame(actions, text='Activity', padding=8)
        log_frame.grid(row=1, column=0, sticky='nsew', pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=16, wrap=tk.WORD, font=('Consolas', 10), relief=tk.FLAT, padx=8, pady=8)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky='nsew')
        log_scroll.grid(row=0, column=1, sticky='ns')

        status = ttk.Label(self, textvariable=self.status_var, style='Hint.TLabel')
        status.grid(row=4, column=0, sticky='w', pady=(8, 0))
        if saved_export:
            self._log('Using the saved Cheddar Up export:\n{}\n'.format(saved_export))
            if os.path.isfile(saved_export):
                self._log('Browse to pick a different file.\n')
            else:
                self._log('That file was not found. Browse to pick the new one.\n')
        else:
            self._log('Choose a Cheddar Up collection summary and a requests workbook, then create a seating chart.\n')
        self._note_newer_download(saved_export)

    def browse_export(self):
        current = self.export_var.get().strip()
        initial = os.path.dirname(current) if current and os.path.isdir(os.path.dirname(current)) else APP_DIR
        path = filedialog.askopenfilename(
            parent=self,
            title='Select the Cheddar Up collection summary',
            initialdir=initial,
            filetypes=[('Excel workbook', '*.xlsx'), ('All files', '*.*')],
        )
        if not path:
            return
        self.export_var.set(path)
        self.export_entry.xview_moveto(1)
        self.status_var.set(os.path.basename(path))
        self._remember_export(path)

    def browse_requests(self):
        current = self.requests_var.get().strip()
        initial = os.path.dirname(current) if current and os.path.isdir(os.path.dirname(current)) else APP_DIR
        path = filedialog.askopenfilename(
            parent=self,
            title='Select the requests workbook',
            initialdir=initial,
            filetypes=[('Excel workbook', '*.xlsx'), ('All files', '*.*')],
        )
        if not path:
            return
        self.requests_var.set(path)
        self.requests_entry.xview_moveto(1)
        self.status_var.set(os.path.basename(path))

    def _downloads_folder(self):
        return cheddar_up.downloads_folder(_load_setting('downloads_folder'))

    def _collection_filter(self):
        return _load_setting('cheddar_up_collection_filter', cheddar_up.DEFAULT_COLLECTION_FILTER)

    def _cheddar_up_url(self):
        return _load_setting('cheddar_up_url', cheddar_up.CHEDDAR_UP_URL) or cheddar_up.CHEDDAR_UP_URL

    def _note_newer_download(self, current):
        try:
            newest = cheddar_up.find_latest_export(self._downloads_folder(), self._collection_filter())
        except Exception:
            return
        if newest and cheddar_up.is_newer(newest, current):
            self._log('A newer Cheddar Up export is in Downloads:\n{}\n'
                      'Click "Download latest from Cheddar Up" to use it.\n'.format(os.path.basename(newest)))

    def open_cheddar_up(self):
        url = self._cheddar_up_url()
        self._log('Opening {}\n'.format(url))
        cheddar_up.open_cheddar_up(url)

    def download_latest_export(self):
        if self._watcher is not None:
            self._stop_watching('Stopped waiting for the Cheddar Up download.')
            return
        folder = self._downloads_folder()
        current = self.export_var.get().strip()
        newest = cheddar_up.find_latest_export(folder, self._collection_filter())
        if newest and cheddar_up.is_newer(newest, current):
            when = datetime.datetime.fromtimestamp(os.path.getmtime(newest)).strftime('%b %d %I:%M %p')
            answer = messagebox.askyesnocancel(
                'Cheddar Up export',
                'This export is already in Downloads:\n\n{}\n(saved {})\n\n'
                'Yes: use it.\nNo: open Cheddar Up and download a new one.'.format(os.path.basename(newest), when),
                parent=self,
            )
            if answer is None:
                return
            if answer:
                self._use_download(newest)
                return
        self._start_watching(folder)

    def _start_watching(self, folder):
        self._watcher = cheddar_up.DownloadWatcher(folder)
        self.download_button.configure(text='Stop waiting')
        self.status_var.set('Waiting for the Cheddar Up download…')
        self._log(
            '\nIn Cheddar Up, open the {} collection, then Manage > Reports > Collection Summary.\n'
            'Leave the file in {}. It is picked up as soon as it finishes downloading.\n'.format(
                self._collection_filter() or 'event', folder))
        self.open_cheddar_up()
        self._watch_job = self.after(1000, self._poll_download)

    def _stop_watching(self, message=None):
        if self._watch_job is not None:
            try:
                self.after_cancel(self._watch_job)
            except Exception:
                pass
        self._watch_job = None
        self._watcher = None
        self.download_button.configure(text='Download latest from Cheddar Up')
        if message:
            self._log(message + '\n')
            self.status_var.set(message)

    def _poll_download(self):
        self._watch_job = None
        watcher = self._watcher
        if watcher is None:
            return
        if self._busy:
            self._watch_job = self.after(1000, self._poll_download)
            return
        path = watcher.check()
        if path:
            self._stop_watching()
            self._use_download(path, offer_chart=True)
            return
        if watcher.expired():
            self._stop_watching('No Cheddar Up download arrived. Click the button to try again, or Browse to pick the file.')
            return
        self._watch_job = self.after(1000, self._poll_download)

    def _use_download(self, path, offer_chart=True):
        exports_dir = os.path.join(APP_DIR, cheddar_up.EXPORTS_FOLDER)
        try:
            saved = cheddar_up.import_export(path, exports_dir)
        except OSError as error:
            self._log('Could not copy {} into {}: {}\nUsing it from Downloads.\n'.format(path, exports_dir, error))
            saved = path
        self._log('\nSaved the Cheddar Up export as:\n{}\n'.format(saved))
        summary = cheddar_up.describe_export(saved)
        for line in cheddar_up.summary_lines(summary):
            self._log(line + '\n')
        if not cheddar_up.matches_collection(saved, self._collection_filter()):
            self._log('WARNING: That file name does not mention {}. Check it is the right collection.\n'.format(
                self._collection_filter()))
        self.export_var.set(saved)
        self.export_entry.xview_moveto(1)
        self.status_var.set(os.path.basename(saved))
        self._remember_export(saved)
        if offer_chart and not summary['problems'] and messagebox.askyesno(
                'Cheddar Up export',
                'Got {}.\n\nCreate the seating chart from it now?'.format(os.path.basename(saved)),
                parent=self):
            self.create_seating_chart()

    def _remember_export(self, path):
        try:
            _save_cheddar_up_export(path)
        except OSError as error:
            self._log('Could not save the Cheddar Up export path: {}\n'.format(error))

    def open_output_folder(self):
        folder = os.path.join(APP_DIR, 'output')
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def create_seating_chart(self):
        path = self.export_var.get().strip()
        requests_path = self.requests_var.get().strip()
        if not path:
            messagebox.showerror('Cheddar Up export', 'Choose a Cheddar Up collection summary.', parent=self)
            return
        if not os.path.isfile(path):
            messagebox.showerror('Cheddar Up export', 'That file was not found:\n{}'.format(path), parent=self)
            return
        if not requests_path:
            messagebox.showerror('Requests', 'Choose the requests workbook.', parent=self)
            return
        if not os.path.isfile(requests_path):
            messagebox.showerror('Requests', 'That file was not found:\n{}'.format(requests_path), parent=self)
            return
        if os.path.normcase(os.path.abspath(path)) == os.path.normcase(os.path.abspath(requests_path)):
            messagebox.showerror('Configuration', 'The Cheddar Up export and the requests workbook have to be different files.', parent=self)
            return
        self._remember_export(path)
        try:
            table_size = _whole_number(self.table_size_var.get(), 'People per table')
            group_limit = _whole_number(self.group_limit_var.get(), 'Group size')
        except ValueError as error:
            messagebox.showerror('Table size', str(error), parent=self)
            return

        self._set_busy(True)
        self._clear_log()
        try:
            self._log('Reading {}\n'.format(path))
            try:
                with _capture_stdout(self.log_text):
                    book = utils.readGuestList(path, requests_path)
                if not book.guests:
                    self._log('No guests were read from tabs 5 and 6.\n')
                    tables = None
                else:
                    self._set_cursor(False)
                    self._log('Check each warning. A guess is not used until you confirm it.\n')
                    corrections_file = os.path.join(APP_DIR, corrections.CORRECTIONS_WORKBOOK)
                    if not review.ask_review(self.winfo_toplevel(), book, corrections_file):
                        self._log('Cancelled. No seating chart was written.\n')
                        self.status_var.set('Cancelled. No files were written.')
                        return
                    book.finish()
                    utils.save_guest_pickle(book)
                    with _capture_stdout(self.log_text):
                        book.report()
                        print('Corrections saved in {}'.format(corrections_file))
                        tables = self.creator.generateSeatingChart(
                            book.guests,
                            book.anti,
                            confirm_large_group=self._confirm_large_group,
                            table_size=table_size,
                            group_limit=group_limit,
                        )
                    extra = book.extra
                    emails = book.emails
                    polls = book.polls
            except self.creator.SeatingCancelled:
                self.status_var.set('Stopped before writing files.')
                messagebox.showwarning(
                    'Seating stopped',
                    'That group was not seated, and no files were written.',
                    parent=self,
                )
                return
            except SystemExit:
                self.status_var.set('Could not create a seating chart.')
                messagebox.showerror(
                    'Seating chart',
                    'The seating chart was not created. The activity log has the reason.',
                    parent=self,
                )
                return

            if not tables:
                self.status_var.set('No guests to seat.')
                messagebox.showerror('Seating chart', 'No guests were read from tabs 5 and 6.', parent=self)
                return

            self._set_cursor(False)
            self._log('\nSet the table numbers in the dialog. Files are written after you save.\n')
            numbered = TableNumberDialog(self.winfo_toplevel(), tables).ask()
            if not numbered:
                self._log('Cancelled. No seating chart was written.\n')
                self.status_var.set('Cancelled. No files were written.')
                return

            timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
            chart = os.path.join('output', 'Seating Chart_{}.xlsx'.format(timestamp))
            with_emails = os.path.join('output', 'Table Rosters_emails_{}.xlsx'.format(timestamp))
            without_emails = os.path.join('output', 'Table Rosters_no emails_{}.xlsx'.format(timestamp))
            self._set_cursor(True)
            with _capture_stdout(self.log_text):
                utils.writeSeatingChart(numbered, extra, timestamp)
                utils.writeTables(numbered, emails, polls, timestamp, writeEmails=True)
                utils.writeTables(numbered, emails, polls, timestamp, writeEmails=False, extraColumn=extra)
            self.status_var.set('Saved a seating chart and both roster workbooks.')
            messagebox.showinfo(
                'Files saved',
                'Saved in the output folder:\n\n{}\n{}\n{}'.format(
                    os.path.basename(chart),
                    os.path.basename(with_emails),
                    os.path.basename(without_emails),
                ),
                parent=self,
            )
        except Exception:
            self._log('\n' + traceback.format_exc())
            self.status_var.set('Something went wrong.')
            messagebox.showerror('Seating chart', 'Something went wrong. The activity log has the details.', parent=self)
        finally:
            self._set_busy(False)

    def create_rosters(self):
        data_path = os.path.join(APP_DIR, 'guest_data.p')
        if not os.path.isfile(data_path):
            messagebox.showerror(
                'Guest details',
                'Create a seating chart once before building rosters from an edited chart. That saves the email addresses this step needs.',
                parent=self,
            )
            return
        path = filedialog.askopenfilename(
            parent=self,
            title='Select a seating chart',
            initialdir=os.path.join(APP_DIR, 'output'),
            filetypes=[('Excel workbook', '*.xlsx'), ('All files', '*.*')],
        )
        if not path:
            return

        self._set_busy(True)
        self._clear_log()
        try:
            self._log('Reading seating chart {}\n'.format(path))
            with _capture_stdout(self.log_text):
                with open(data_path, 'rb') as handle:
                    _guests, _anti_requests, emails, polls = pickle.load(handle)
                tables, choices = self.rosters.readSeatingChart(path)
            if not tables:
                self.status_var.set('That seating chart has no guests.')
                messagebox.showerror('Seating chart', 'That workbook has no guest rows on the first sheet.', parent=self)
                return
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
            with_emails = os.path.join('output', 'Table Rosters_emails_{}.xlsx'.format(timestamp))
            without_emails = os.path.join('output', 'Table Rosters_no emails_{}.xlsx'.format(timestamp))
            with _capture_stdout(self.log_text):
                utils.writeTables(tables, emails, polls, timestamp, writeEmails=True, extraColumn=choices)
                utils.writeTables(tables, emails, polls, timestamp, writeEmails=False, extraColumn=choices)
            self.status_var.set('Saved both roster workbooks.')
            messagebox.showinfo(
                'Files saved',
                'Saved in the output folder:\n\n{}\n{}'.format(
                    os.path.basename(with_emails),
                    os.path.basename(without_emails),
                ),
                parent=self,
            )
        except SystemExit:
            self.status_var.set('Could not read that seating chart.')
            messagebox.showerror('Seating chart', 'The activity log has the reason.', parent=self)
        except Exception:
            self._log('\n' + traceback.format_exc())
            self.status_var.set('Something went wrong.')
            messagebox.showerror('Rosters', 'Something went wrong. The activity log has the details.', parent=self)
        finally:
            self._set_busy(False)

    def _confirm_large_group(self, group, size, limit):
        names = '\n'.join(sorted(utils.formatNameWithoutFamilyName(name) for name in group))
        return messagebox.askyesno(
            'Large group',
            'This group has {} people. The usual limit is {}.\n\n{}\n\nSeat them at one table anyway?'.format(
                size, limit, names),
            parent=self.winfo_toplevel(),
        )

    def _log(self, text):
        _LogStream(self.log_text).write(text)

    def _clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _set_cursor(self, busy):
        self.winfo_toplevel().configure(cursor='watch' if busy else '')
        self.update_idletasks()

    def _set_busy(self, busy):
        self._busy = busy
        state = ['disabled'] if busy else ['!disabled']
        self.create_button.state(state)
        self.roster_button.state(state)
        self.download_button.state(state)
        self._set_cursor(busy)


def _whole_number(value, label):
    text = str(value).strip()
    if text.endswith('.0'):
        text = text[:-2]
    if not text.isdigit() or int(text) < 1:
        raise ValueError('{} needs a whole number of at least 1.'.format(label))
    return int(text)


def _apply_styles():
    style = ttk.Style()
    if 'vista' in style.theme_names():
        style.theme_use('vista')
    style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'))
    style.configure('Hint.TLabel', foreground='#555555')
    style.configure('TButton', padding=(10, 6))


def main(creator=None):
    global APP_DIR
    APP_DIR = work_dir()
    os.chdir(APP_DIR)
    _install_excepthook()
    if creator is None:
        creator = _load_module('seating_creator', 'Seating Chart Creator.py')
    rosters = _load_module('table_rosters', 'Create Table Rosters From Seating Chart.py')
    _enable_dpi()
    root = tk.Tk()
    root.title('Seating Chart Creator')
    root.minsize(780, 620)
    root.geometry('980x740')
    _apply_styles()
    utils.keep_minmax(root)
    App(root, creator, rosters)
    root.mainloop()


if __name__ == '__main__':
    main()
