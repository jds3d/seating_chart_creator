# -*- coding: utf-8 -*-
"""Window for checking warnings and saving corrections."""

import tkinter as tk
from tkinter import messagebox, ttk

import corrections
import utils


_INTROS = {
    'missing': (
        'These payers did not type attendee names. Enter every person who should be seated. '
        'If a seatmate request mentions someone who is missing, add that person here first.'
    ),
    'names': 'These names have no last name. Enter the full name, then use it.',
    'emails': 'Choose which email address to keep. Nothing is selected until you choose.',
    'choices': 'Choose which meal to keep.',
    'unclaimed': 'These meal choices were not tied to a person. Assign each one, or confirm that it stays unassigned.',
    'duplicates': 'Say whether each pair is two people or the same person.',
    'seatmates': (
        'Each box is one thing somebody wrote. A guess is filled in when a guest was found. '
        'A line that says "not on the guest list" was missed, the way "Joe Boyer and Eileen Mullin" can keep Eileen and drop Joe. '
        'Choose the right guest for every person, or mark that person not attending, then click This is correct. '
        'Add a missing person under No attendee names first if they are not in the list. '
        'A match that already looks exact can be changed with Change.'
    ),
    'pairs': 'These rows are from the requests workbook. Choose the guests on each side, or skip the row.',
}


class ReviewDialog(tk.Toplevel):
    def __init__(self, master, book, path):
        super().__init__(master)
        self.title('Fix warnings')
        self.book = book
        self.path = path
        self.accepted = False
        self._pull = lambda: None
        self._section = None
        try:
            self.loaded = corrections.load_corrections(path)
        except PermissionError:
            self.loaded = None
        self.state = corrections.build_state(book, self.loaded or {key: [] for key in corrections._SHEETS})
        self.minsize(860, 560)
        self.geometry('1040x700')
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)
        self.transient(master)
        self.protocol('WM_DELETE_WINDOW', self._cancel)

        ttk.Label(
            self,
            wraplength=980,
            text='Nothing is seated until these are settled. Confirmed decisions are saved in Corrections.xlsx, which you can also edit in Excel.',
        ).grid(row=0, column=0, columnspan=2, sticky='ew', padx=16, pady=(16, 8))

        self.section_list = tk.Listbox(self, width=32, exportselection=False, activestyle='none', font=('Segoe UI', 10))
        self.section_list.grid(row=1, column=0, sticky='ns', padx=(16, 8), pady=8)
        self.section_list.bind('<<ListboxSelect>>', self._on_select)

        holder = ttk.Frame(self)
        holder.grid(row=1, column=1, sticky='nsew', padx=(0, 16), pady=8)
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(holder, highlightthickness=0)
        scroll = ttk.Scrollbar(holder, orient=tk.VERTICAL, command=self.canvas.yview)
        self.body = ttk.Frame(self.canvas)
        self._window = self.canvas.create_window((0, 0), window=self.body, anchor='nw')
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        scroll.grid(row=0, column=1, sticky='ns')
        self.canvas.bind('<Configure>', lambda event: self.canvas.itemconfigure(self._window, width=event.width))
        self.body.bind('<Configure>', lambda _event: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<MouseWheel>', lambda event: self.canvas.yview_scroll(int(-event.delta / 120), 'units'))

        buttons = ttk.Frame(self)
        buttons.grid(row=2, column=0, columnspan=2, sticky='e', padx=16, pady=(0, 16))
        ttk.Button(buttons, text='Cancel', command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(buttons, text='Save corrections', command=self._save_only).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(buttons, text='Continue', command=self._continue).pack(side=tk.RIGHT, padx=(0, 8))

        self._refresh_list()
        self._select_first()

    def ask(self):
        if self.loaded is None:
            messagebox.showerror(
                'Corrections',
                'Corrections.xlsx is open in Excel. Close it, then create the seating chart again.',
                parent=self.master,
            )
            self.destroy()
            return False
        if not corrections.section_defs(self.state):
            corrections.apply_all(self.book, self.state)
            self.destroy()
            return True
        self.grab_set()
        self.wait_window()
        return self.accepted

    def _refresh_list(self):
        current = self._section
        self.section_list.delete(0, tk.END)
        self._keys = []
        for key, title, pending in corrections.section_defs(self.state):
            label = title if pending == 0 else '{} ({})'.format(title, pending)
            self.section_list.insert(tk.END, label)
            if pending:
                self.section_list.itemconfig(tk.END, foreground='#8a4b08')
            self._keys.append(key)
        if current in self._keys:
            self.section_list.selection_set(self._keys.index(current))

    def _select_first(self):
        pending = [index for index, key in enumerate(self._keys) if corrections.section_defs(self.state)[index][2]]
        index = pending[0] if pending else 0
        if not self._keys:
            return
        self.section_list.selection_clear(0, tk.END)
        self.section_list.selection_set(index)
        self.section_list.activate(index)
        self._show(self._keys[index], pull=False)

    def _on_select(self, _event):
        selection = self.section_list.curselection()
        if not selection:
            return
        key = self._keys[selection[0]]
        if key != self._section:
            self._show(key, pull=True)

    def _show(self, key, pull):
        if pull:
            self._pull()
        if key == 'seatmates':
            corrections.refresh_open_seatmates(self.book, self.state)
        self._section = key
        for child in self.body.winfo_children():
            child.destroy()
        intro = ttk.Label(self.body, text=_INTROS[key], wraplength=680)
        intro.pack(fill=tk.X, padx=8, pady=(4, 12))
        builders = {
            'missing': self._build_missing,
            'names': self._build_names,
            'emails': self._build_emails,
            'choices': self._build_choices,
            'unclaimed': self._build_unclaimed,
            'duplicates': self._build_duplicates,
            'seatmates': self._build_seatmates,
            'pairs': self._build_pairs,
        }
        self._pull = builders[key]()
        self.canvas.yview_moveto(0)
        self._refresh_list()

    def _card(self, needs_answer):
        card = ttk.Frame(self.body, padding=10)
        card.pack(fill=tk.X, padx=8, pady=6)
        banner = ttk.Label(card, text='Needs an answer' if needs_answer else 'Confirmed')
        banner.pack(anchor='w')
        if needs_answer:
            banner.configure(foreground='#8a4b08')
        else:
            banner.configure(foreground='#1b7f3a')
        return card

    def _build_missing(self):
        pulls = []
        for item in self.state['missing'].values():
            card = self._card(not item['decided'])
            ttk.Label(card, text=item['payer'], font=('Segoe UI', 11, 'bold')).pack(anchor='w')
            if item['email']:
                ttk.Label(card, text='Email on the payment: {}'.format(item['email'])).pack(anchor='w')
            rows = ttk.Frame(card)
            rows.pack(fill=tk.X, pady=(8, 0))
            editors = []

            def add_row(person, parent=rows, store=editors, choices=item['choices']):
                line = ttk.Frame(parent)
                line.pack(fill=tk.X, pady=2)
                name = tk.StringVar(value=person.get('name', ''))
                choice = tk.StringVar(value=person.get('choice', ''))
                ttk.Entry(line, textvariable=name, width=32).pack(side=tk.LEFT)
                box = ttk.Combobox(line, textvariable=choice, values=choices, width=18)
                box.pack(side=tk.LEFT, padx=8)
                ttk.Button(line, text='Remove', command=lambda target=line, saved=store: remove_row(target, saved)).pack(side=tk.LEFT)
                store.append((line, name, choice))

            def remove_row(line, store):
                if len(store) == 1:
                    return
                store[:] = [row for row in store if row[0] is not line]
                line.destroy()

            for person in item['people']:
                add_row(person)
            ttk.Button(card, text='Add a person', command=lambda add=add_row: add({'name': '', 'choice': ''})).pack(anchor='w', pady=(6, 0))

            def confirm(target=item, store=editors):
                people = [{'name': name.get().strip(), 'choice': choice.get().strip()} for _line, name, choice in store]
                if not people or any(not person['name'] for person in people):
                    messagebox.showerror('Attendees', 'Enter a name on every row.', parent=self)
                    return
                target['people'] = people
                target['decided'] = True
                self._show('missing', pull=False)

            ttk.Button(card, text='These people are correct', command=confirm).pack(anchor='w', pady=(8, 0))

            def pull(target=item, store=editors):
                target['people'] = [{'name': name.get().strip(), 'choice': choice.get().strip()} for _line, name, choice in store]

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _build_names(self):
        pulls = []
        for item in self.state['names'].values():
            card = self._card(not item['decided'])
            ttk.Label(
                card,
                text='"{}" was entered on {}\'s purchase without a last name.'.format(item['written'], item['payer']),
                wraplength=640,
            ).pack(anchor='w')
            name = tk.StringVar(value=item.get('name', ''))
            ttk.Entry(card, textvariable=name, width=36).pack(anchor='w', pady=(8, 0))

            def confirm(target=item, var=name):
                full = var.get().strip()
                if ' ' not in full:
                    messagebox.showerror('Name', 'Enter a first and last name.', parent=self)
                    return
                target['name'] = full
                target['decided'] = True
                self._show('names', pull=False)

            ttk.Button(card, text='Use this name', command=confirm).pack(anchor='w', pady=(8, 0))

            def pull(target=item, var=name):
                target['name'] = var.get().strip()

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _build_emails(self):
        return self._build_either(self.state['emails'], 'email', 'Choose this email')

    def _build_choices(self):
        return self._build_either(self.state['choices'], 'choice', 'Choose this meal')

    def _build_either(self, items, field, button_text):
        pulls = []
        for item in items.values():
            card = self._card(not item['decided'])
            ttk.Label(card, text=item['guest'], font=('Segoe UI', 11, 'bold')).pack(anchor='w')
            chosen = tk.StringVar(value=item.get(field, '') if item['decided'] else '')
            ttk.Radiobutton(card, text=item['first'], value=item['first'], variable=chosen).pack(anchor='w', pady=(6, 0))
            ttk.Radiobutton(card, text=item['second'], value=item['second'], variable=chosen).pack(anchor='w')

            def confirm(target=item, var=chosen):
                if not var.get():
                    messagebox.showerror('Choose one', 'Select one of the two values.', parent=self)
                    return
                target[field] = var.get()
                target['decided'] = True
                self._show(self._section, pull=False)

            ttk.Button(card, text=button_text, command=confirm).pack(anchor='w', pady=(8, 0))

            def pull(target=item, var=chosen):
                if var.get():
                    target[field] = var.get()

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _build_unclaimed(self):
        pulls = []
        names = corrections.preview_names(self.book, self.state)
        for item in self.state['unclaimed'].values():
            pending = any(not row['decided'] for row in item['rows'])
            card = self._card(pending)
            ttk.Label(card, text=item['payer'], font=('Segoe UI', 11, 'bold')).pack(anchor='w')
            editors = []
            for row in item['rows']:
                line = ttk.Frame(card)
                line.pack(fill=tk.X, pady=3)
                ttk.Label(line, text=row['choice'], width=18).pack(side=tk.LEFT)
                guest = tk.StringVar(value=row.get('guest', ''))
                ttk.Combobox(line, textvariable=guest, values=[''] + names, width=28).pack(side=tk.LEFT, padx=8)
                editors.append((row, guest))

            def confirm(store=editors):
                for row, guest in store:
                    row['guest'] = self._known_name(guest.get())
                    row['decided'] = True
                self._show('unclaimed', pull=False)

            ttk.Button(card, text='Use these assignments', command=confirm).pack(anchor='w', pady=(8, 0))
            ttk.Label(card, text='Leave a choice blank to keep it unassigned.').pack(anchor='w')

            def pull(store=editors):
                for row, guest in store:
                    row['guest'] = guest.get().strip()

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _build_duplicates(self):
        pulls = []
        for item in self.state['duplicates'].values():
            left, right = item['pair']
            card = self._card(not item['decided'])
            ttk.Label(card, text='{} and {}'.format(left, right), font=('Segoe UI', 11, 'bold')).pack(anchor='w')
            choice = tk.StringVar(value='same' if item.get('decided') and item.get('same_as') else ('separate' if item.get('decided') else ''))
            keep = tk.StringVar(value=item.get('same_as') or left)
            ttk.Radiobutton(card, text='Keep them as two people', value='separate', variable=choice).pack(anchor='w', pady=(6, 0))
            ttk.Radiobutton(card, text='They are the same person', value='same', variable=choice).pack(anchor='w')
            ttk.Combobox(card, textvariable=keep, values=[left, right], width=32, state='readonly').pack(anchor='w', padx=(18, 0))

            def confirm(target=item, which=choice, kept=keep):
                if which.get() not in ('separate', 'same'):
                    messagebox.showerror('Duplicates', 'Choose whether they are the same person.', parent=self)
                    return
                target['decided'] = True
                target['same_as'] = kept.get() if which.get() == 'same' else ''
                self._show('duplicates', pull=False)

            ttk.Button(card, text='Use this answer', command=confirm).pack(anchor='w', pady=(8, 0))

            def pull(target=item, which=choice, kept=keep):
                if which.get() == 'same':
                    target['same_as'] = kept.get()
                elif which.get() == 'separate':
                    target['same_as'] = ''

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _build_seatmates(self):
        pulls = []
        names = corrections.preview_names(self.book, self.state)
        ordered = sorted(self.state['seatmates'].values(), key=lambda item: (item['decided'], item['raw'].casefold()))
        for item in ordered:
            card = self._card(not item['decided'])
            ttk.Label(card, text='"{}"'.format(item['raw']), font=('Segoe UI', 11, 'bold'), wraplength=640).pack(anchor='w')
            if item['asked_by']:
                ttk.Label(card, text='Asked by {}'.format(', '.join(item['asked_by'])), wraplength=640).pack(anchor='w')
            if item['decided']:
                if item['skip']:
                    ttk.Label(card, text='Not a seating request.').pack(anchor='w', pady=(6, 0))
                else:
                    ttk.Label(card, text='Sits with {}'.format(', '.join(item['sit_with'])), wraplength=640).pack(anchor='w', pady=(6, 0))

                def change(target=item):
                    self._open_seatmate(target)
                    self._show('seatmates', pull=False)

                ttk.Button(card, text='Change', command=change).pack(anchor='w', pady=(8, 0))
                continue

            skip = tk.BooleanVar(value=item.get('skip', False))
            ttk.Checkbutton(card, text='This is not a seating request', variable=skip).pack(anchor='w', pady=(8, 0))
            editors = []
            for part in item['parts']:
                line = ttk.Frame(card)
                line.pack(fill=tk.X, pady=3)
                label = part['part'] or 'Also sit with'
                note = ''
                if part.get('approximate'):
                    note = ' (spelling differs)'
                elif not part.get('guest') and not part.get('added'):
                    note = ' (not on the guest list)'
                ttk.Label(line, text=label + note, wraplength=220, width=28).pack(side=tk.LEFT)
                guest = tk.StringVar(value='' if part.get('absent') else part.get('guest', ''))
                absent = tk.BooleanVar(value=part.get('absent', False))
                box = ttk.Combobox(line, textvariable=guest, values=names, width=28)
                box.pack(side=tk.LEFT, padx=8)
                if absent.get():
                    box.configure(state='disabled')

                def mark_absent(var=absent, combo=box, value=guest, target=item):
                    var.set(True)
                    value.set('')
                    combo.configure(state='disabled')
                    target['touched'] = True

                ttk.Button(line, text='Not attending', command=mark_absent).pack(side=tk.LEFT)
                editors.append((part, guest, absent))

            def add_person(target=item, store=editors, skip_var=skip):
                self._write_seatmate(target, skip_var, store, decide=False)
                target['parts'].append({'part': '', 'guest': '', 'absent': False, 'approximate': False, 'added': True})
                target['touched'] = True
                self._show('seatmates', pull=False)

            def confirm(target=item, store=editors, skip_var=skip):
                self._write_seatmate(target, skip_var, store, decide=True)
                if not target['decided']:
                    return
                self._show('seatmates', pull=False)

            ttk.Button(card, text='Add a person', command=add_person).pack(anchor='w', pady=(6, 0))
            ttk.Button(card, text='This is correct', command=confirm).pack(anchor='w', pady=(4, 0))

            def pull(target=item, store=editors, skip_var=skip):
                self._write_seatmate(target, skip_var, store, decide=False)

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _open_seatmate(self, item):
        item['decided'] = False
        item['from_file'] = False
        item['touched'] = True
        if item.get('skip'):
            item['parts'] = [{'part': item['raw'], 'guest': '', 'absent': True, 'approximate': False}]
            return
        if item['sit_with']:
            item['parts'] = [{
                'part': name,
                'guest': name,
                'absent': False,
                'approximate': False,
            } for name in item['sit_with']]
        else:
            item['parts'] = [{'part': item['raw'], 'guest': '', 'absent': False, 'approximate': False}]

    def _write_seatmate(self, item, skip, editors, decide):
        item['touched'] = True
        item['skip'] = bool(skip.get())
        parts = []
        for part, guest, absent in editors:
            parts.append({
                'part': part.get('part', ''),
                'guest': '' if absent.get() else guest.get().strip(),
                'absent': bool(absent.get()),
                'approximate': part.get('approximate', False),
                'added': part.get('added', False),
            })
        item['parts'] = parts
        if item['skip']:
            item['sit_with'] = []
            if decide:
                item['decided'] = True
            return
        names = corrections.preview_names(self.book, self.state)
        chosen = []
        for part in parts:
            if part['absent']:
                continue
            if not part['guest']:
                if decide:
                    messagebox.showerror(
                        'Seatmate request',
                        'Choose a guest for "{}", or mark that part not attending.'.format(part['part'] or item['raw']),
                        parent=self,
                    )
                item['decided'] = False
                return
            match = self._known_name(part['guest'], names)
            if not match:
                if decide:
                    messagebox.showerror(
                        'Seatmate request',
                        '"{}" is not on the guest list. Add that person under No attendee names, or mark them not attending.'.format(part['guest']),
                        parent=self,
                    )
                item['decided'] = False
                return
            part['guest'] = match
            if match not in chosen:
                chosen.append(match)
        item['sit_with'] = chosen
        if decide:
            if not chosen:
                if parts and all(part['absent'] for part in parts):
                    item['skip'] = True
                    item['decided'] = True
                    return
                messagebox.showerror(
                    'Seatmate request',
                    'Choose who should sit together, or mark this as not a seating request.',
                    parent=self,
                )
                item['decided'] = False
                return
            item['decided'] = True

    def _build_pairs(self):
        pulls = []
        names = corrections.preview_names(self.book, self.state)
        for item in self.state['pairs'].values():
            card = self._card(not item['decided'])
            kind = 'Sit together' if item['kind'] == 'request' else 'Do not sit together'
            ttk.Label(card, text=kind, font=('Segoe UI', 11, 'bold')).pack(anchor='w')
            ttk.Label(card, text='"{}"  /  "{}"'.format(item['left'], item['right']), wraplength=640).pack(anchor='w')
            if item['decided']:
                if item['skip']:
                    ttk.Label(card, text='Skipped.').pack(anchor='w', pady=(6, 0))
                else:
                    ttk.Label(
                        card,
                        text='{}  with  {}'.format(', '.join(item['left_list']), ', '.join(item['right_list'])),
                        wraplength=640,
                    ).pack(anchor='w', pady=(6, 0))

                def change(target=item):
                    target['decided'] = False
                    self._show('pairs', pull=False)

                ttk.Button(card, text='Change', command=change).pack(anchor='w', pady=(8, 0))
                continue

            skip = tk.BooleanVar(value=item.get('skip', False))
            left = tk.StringVar(value='; '.join(item.get('left_list', [])))
            right = tk.StringVar(value='; '.join(item.get('right_list', [])))
            ttk.Checkbutton(card, text='Skip this row', variable=skip).pack(anchor='w', pady=(8, 0))
            ttk.Label(card, text='Left guests').pack(anchor='w')
            ttk.Combobox(card, textvariable=left, values=names, width=36).pack(anchor='w')
            ttk.Label(card, text='Right guests').pack(anchor='w', pady=(6, 0))
            ttk.Combobox(card, textvariable=right, values=names, width=36).pack(anchor='w')
            ttk.Label(card, text='Separate more than one guest with a semicolon.').pack(anchor='w')

            def confirm(target=item, left_var=left, right_var=right, skip_var=skip):
                self._write_pair(target, left_var, right_var, skip_var, decide=True)
                if target['decided']:
                    self._show('pairs', pull=False)

            ttk.Button(card, text='This is correct', command=confirm).pack(anchor='w', pady=(8, 0))

            def pull(target=item, left_var=left, right_var=right, skip_var=skip):
                self._write_pair(target, left_var, right_var, skip_var, decide=False)

            pulls.append(pull)
        return lambda: [pull() for pull in pulls]

    def _write_pair(self, item, left_var, right_var, skip_var, decide):
        item['skip'] = bool(skip_var.get())
        if item['skip']:
            item['left_list'] = []
            item['right_list'] = []
            if decide:
                item['decided'] = True
            return
        names = corrections.preview_names(self.book, self.state)
        left = self._name_list(left_var.get(), names)
        right = self._name_list(right_var.get(), names)
        if decide and (left is None or right is None or not left or not right):
            messagebox.showerror('Requests workbook', 'Choose a guest on both sides, or skip this row.', parent=self)
            item['decided'] = False
            return
        item['left_list'] = left or []
        item['right_list'] = right or []
        if decide:
            item['decided'] = True

    def _name_list(self, text, names):
        parts = [utils._norm(part) for part in text.split(';') if utils._norm(part)]
        chosen = []
        for part in parts:
            match = self._known_name(part, names)
            if not match:
                return None
            if match not in chosen:
                chosen.append(match)
        return chosen

    def _known_name(self, text, names=None):
        if names is None:
            names = corrections.preview_names(self.book, self.state)
        for name in names:
            if utils._key(name) == utils._key(text):
                return name
        return ''

    def _save_only(self):
        self._pull()
        if self._write_file(notify=True):
            return

    def _write_file(self, notify):
        try:
            corrections.save_corrections(self.path, self.state, self.loaded)
        except PermissionError:
            messagebox.showerror(
                'Corrections',
                'Corrections.xlsx is open in Excel. Close it, then save again.',
                parent=self,
            )
            return False
        if notify:
            messagebox.showinfo('Corrections', 'Saved {}'.format(self.path), parent=self)
        return True

    def _continue(self):
        self._pull()
        errors = corrections.validate(self.book, self.state)
        if errors:
            messagebox.showerror('Still needs an answer', '\n'.join(errors[:12]), parent=self)
            self._refresh_list()
            for index, key in enumerate(self._keys):
                if corrections.section_defs(self.state)[index][2]:
                    self.section_list.selection_clear(0, tk.END)
                    self.section_list.selection_set(index)
                    self._show(key, pull=False)
                    break
            return
        corrections.apply_all(self.book, self.state)
        if not self._write_file(notify=False):
            return
        self.accepted = True
        self.destroy()

    def _cancel(self):
        self.accepted = False
        self.destroy()


def ask_review(parent, book, path):
    dialog = ReviewDialog(parent, book, path)
    return dialog.ask()
