# -*- coding: utf-8 -*-
"""Window for checking warnings and saving corrections."""

import tkinter as tk
from tkinter import messagebox, ttk

import corrections
import utils


class ReviewDialog(tk.Toplevel):
    def __init__(self, master, book, path):
        super().__init__(master)
        self.title('Fix warnings')
        self.book = book
        self.path = path
        self.accepted = False
        self._pull = lambda: None
        self._current = None
        self._nav_lock = False
        self._iid_info = {}
        self._wrapping = []
        try:
            self.loaded = corrections.load_corrections(path)
        except PermissionError:
            self.loaded = None
        self.state = corrections.build_state(book, self.loaded or {key: [] for key in corrections._SHEETS})
        self.minsize(860, 560)
        self.geometry('1040x700')
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)
        self.protocol('WM_DELETE_WINDOW', self._cancel)

        ttk.Label(
            self,
            wraplength=980,
            text=(
                'One warning is shown at a time. '
                'These people are correct, or Skip for now, moves to the next one. '
                'Confirmed answers are saved in Corrections.xlsx, which you can also edit in Excel.'
            ),
        ).grid(row=0, column=0, columnspan=2, sticky='ew', padx=16, pady=(16, 8))

        side = ttk.Frame(self)
        side.grid(row=1, column=0, sticky='ns', padx=(16, 8), pady=8)
        side.rowconfigure(0, weight=1)
        style = ttk.Style(self)
        style.configure('Review.Treeview', font=('Segoe UI', 10), rowheight=26)
        self.tree = ttk.Treeview(side, show='tree', selectmode='browse', style='Review.Treeview')
        self.tree.column('#0', width=340, minwidth=240, stretch=False)
        self.tree.tag_configure('pending', foreground='#8a4b08')
        self.tree.tag_configure('done', foreground='#1b7f3a')
        self.tree.tag_configure('group', font=('Segoe UI', 10, 'bold'))
        tree_scroll = ttk.Scrollbar(side, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.grid(row=0, column=0, sticky='ns')
        tree_scroll.grid(row=0, column=1, sticky='ns')
        self.tree.bind('<<TreeviewSelect>>', self._on_select)

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
        self.canvas.bind('<Configure>', self._fit_canvas)
        self.body.bind('<Configure>', lambda _event: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<MouseWheel>', lambda event: self.canvas.yview_scroll(int(-event.delta / 120), 'units'))

        buttons = ttk.Frame(self)
        buttons.grid(row=2, column=0, columnspan=2, sticky='e', padx=16, pady=(0, 16))
        ttk.Button(buttons, text='Cancel', command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(buttons, text='Save All', command=self._save_only).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(buttons, text='Continue', command=self._continue).pack(side=tk.RIGHT, padx=(0, 8))

        first = self._first_pending() or self._first_item()
        if first:
            self._show_node(first, pull=False)
        utils.keep_minmax(self, master)

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

    def _all_items(self):
        items = []
        for section, _title, _pending in corrections.section_defs(self.state):
            items.extend(self._section_items(section))
        return items

    def _section_items(self, section):
        rows = self.state[section]
        items = []
        if section == 'missing':
            for key, item in rows.items():
                items.append(self._node(section, key, item['payer'], not item['decided']))
        elif section == 'names':
            for key, item in rows.items():
                items.append(self._node(section, key, item['written'], not item['decided']))
        elif section in ('emails', 'choices'):
            for key, item in rows.items():
                items.append(self._node(section, key, item['guest'], not item['decided']))
        elif section == 'unclaimed':
            for key, item in rows.items():
                for index, row in enumerate(item['rows']):
                    label = '{} — {}'.format(item['payer'], row['choice'] or 'meal')
                    items.append(self._node(section, key, label, not row['decided'], row=index))
        elif section == 'duplicates':
            for key, item in rows.items():
                left, right = item['pair']
                items.append(self._node(section, key, '{} and {}'.format(left, right), not item['decided']))
        elif section == 'seatmates':
            for key, item in rows.items():
                items.append(self._node(section, key, _short(utils._display_phrase(item['raw'])), not item['decided']))
        elif section == 'pairs':
            for key, item in rows.items():
                kind = 'Sit together' if item['kind'] == 'request' else 'Keep apart'
                label = '{}: {} / {}'.format(
                    kind,
                    _short(utils._display_phrase(item['left']), 24),
                    _short(utils._display_phrase(item['right']), 24),
                )
                items.append(self._node(section, key, label, not item['decided']))
        return items

    def _node(self, section, key, label, pending, row=None):
        return {
            'kind': 'item',
            'section': section,
            'key': key,
            'row': row,
            'label': label,
            'pending': pending,
            'id': (section, key, row),
        }

    def _first_item(self):
        items = self._all_items()
        return items[0] if items else None

    def _first_pending(self):
        for item in self._all_items():
            if item['pending']:
                return item
        return None

    def _current_node(self):
        for item in self._all_items():
            if item['id'] == self._current:
                return item
        return None

    def _first_matching(self, section=None, pending=None):
        for item in self._all_items():
            if section is not None and item['section'] != section:
                continue
            if pending is not None and item['pending'] != pending:
                continue
            return item
        return None

    def _refresh_tree(self):
        self._nav_lock = True
        try:
            self.tree.delete(*self.tree.get_children())
            self._iid_info = {}
            selected = None
            groups = (
                ('Needs an answer', True, 'pending'),
                ('Confirmed', False, 'done'),
            )
            for group_title, want_pending, tag in groups:
                blocks = []
                for section, title, _pending in corrections.section_defs(self.state):
                    children = [child for child in self._section_items(section) if child['pending'] is want_pending]
                    if children:
                        blocks.append((section, title, children))
                if not blocks:
                    continue
                total = sum(len(children) for _section, _title, children in blocks)
                group_iid = self.tree.insert(
                    '', 'end', text='{} ({})'.format(group_title, total), open=True, tags=('group', tag),
                )
                self._iid_info[group_iid] = {'kind': 'group', 'pending': want_pending}
                for section, title, children in blocks:
                    parent = self.tree.insert(
                        group_iid, 'end', text='{} ({})'.format(title, len(children)), open=True, tags=(tag,),
                    )
                    self._iid_info[parent] = {'kind': 'section', 'section': section, 'pending': want_pending}
                    for child in children:
                        iid = self.tree.insert(parent, 'end', text=child['label'], tags=(tag,))
                        self._iid_info[iid] = child
                        if child['id'] == self._current:
                            selected = iid
            if selected:
                self.tree.selection_set(selected)
                self.tree.focus(selected)
                self.tree.see(selected)
        finally:
            self._nav_lock = False

    def _on_select(self, _event):
        if self._nav_lock:
            return
        iid = self.tree.focus()
        info = self._iid_info.get(iid)
        if not info:
            return
        if info['kind'] == 'group':
            target = self._first_matching(pending=info['pending'])
        elif info['kind'] == 'section':
            target = self._first_matching(section=info['section'], pending=info['pending'])
        elif info['id'] == self._current:
            return
        else:
            self._show_node(info, pull=True)
            return
        if target and target['id'] != self._current:
            self._show_node(target, pull=True)

    def _show_node(self, node, pull):
        if not node:
            return
        if pull:
            self._pull()
        if node['section'] == 'seatmates':
            corrections.refresh_open_seatmates(self.book, self.state)
        self._current = node['id']
        self._wrapping = []
        for child in self.body.winfo_children():
            child.destroy()
        item = self.state[node['section']][node['key']]
        builders = {
            'missing': self._build_missing,
            'names': self._build_names,
            'emails': self._build_emails,
            'choices': self._build_choices,
            'duplicates': self._build_duplicates,
            'seatmates': self._build_seatmates,
            'pairs': self._build_pairs,
        }
        if node['section'] == 'unclaimed':
            self._pull = self._build_unclaimed(item, item['rows'][node['row']])
        else:
            self._pull = builders[node['section']](item)
        self.canvas.yview_moveto(0)
        self._refresh_tree()
        self.after_idle(self._fit_wrapping)

    def _go_next(self):
        items = self._all_items()
        if not items:
            return
        index = 0
        for position, item in enumerate(items):
            if item['id'] == self._current:
                index = position
                break
        for item in items[index + 1:] + items[:index]:
            if item['pending']:
                self._show_node(item, pull=False)
                return
        current = self._current_node()
        if current:
            self._show_node(current, pull=False)

    def _skip(self):
        self._pull()
        self._go_next()

    def _actions(self, parent, confirm_text, confirm_command):
        row = ttk.Frame(parent)
        row.pack(anchor='w', pady=(12, 0))
        ttk.Button(row, text=confirm_text, command=confirm_command).pack(side=tk.LEFT)
        ttk.Button(row, text='Skip for now', command=self._skip).pack(side=tk.LEFT, padx=(8, 0))

    def _fit_canvas(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)
        self._fit_wrapping(event.width)

    def _fit_wrapping(self, width=None):
        if width is None:
            width = self.canvas.winfo_width()
        if width <= 1:
            return
        for label, pad in self._wrapping:
            try:
                label.configure(wraplength=max(width - pad, 180))
            except tk.TclError:
                pass

    def _wrap_label(self, parent, pad=56, **kwargs):
        label = ttk.Label(parent, justify='left', **kwargs)
        self._wrapping.append((label, pad))
        width = self.canvas.winfo_width()
        if width > 1:
            label.configure(wraplength=max(width - pad, 180))
        return label

    def _explain(self, parent, text):
        self._wrap_label(parent, text=text).pack(anchor='w', fill=tk.X, pady=(6, 8))

    def _card(self, needs_answer):
        card = ttk.Frame(self.body, padding=12)
        card.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        banner = ttk.Label(card, text='Needs an answer' if needs_answer else 'Confirmed')
        banner.pack(anchor='w')
        if needs_answer:
            banner.configure(foreground='#8a4b08')
        else:
            banner.configure(foreground='#1b7f3a')
        return card

    def _build_missing(self, item):
        card = self._card(not item['decided'])
        ttk.Label(card, text='Who else is attending?' if item.get('kind') == 'short' else 'Who is attending?',
                  font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        choices = [choice for choice in item['choices'] if choice]
        if len(choices) == 1:
            meal_sentence = 'A meal was ordered without a name: {}.'.format(choices[0])
        elif choices:
            meal_sentence = 'Meals were ordered without a name: {}.'.format(', '.join(choices))
        else:
            meal_sentence = 'No meal was listed either.'
        if item.get('kind') == 'short':
            self._explain(
                card,
                'This order paid for {} tickets but names only {}: {}. {} '
                'Enter the {} other {} below, and choose each meal. They are seated with the people already named.'.format(
                    item['tickets'], len(item['named']), ', '.join(item['named']), meal_sentence,
                    item['extra'], 'person' if item['extra'] == 1 else 'people'),
            )
        else:
            self._explain(
                card,
                'Cheddar Up has no attendee names on this payment. {} '
                'The name already in the box is only a guess: it is the person who paid, not a name from the form. '
                'Change it if someone else is coming. Add a row for each other person, and choose that person\'s meal.'.format(meal_sentence),
            )
        ttk.Label(card, text='Payer: {}'.format(item['payer'])).pack(anchor='w')
        if item['email']:
            ttk.Label(card, text='Email on the payment: {}'.format(item['email'])).pack(anchor='w')

        rows = ttk.Frame(card)
        rows.pack(fill=tk.X, pady=(10, 0))
        headers = ttk.Frame(rows)
        headers.pack(fill=tk.X)
        ttk.Label(headers, text='Who is coming' if item.get('kind') != 'short' else 'Who else is coming', width=32).pack(side=tk.LEFT)
        ttk.Label(headers, text='Meal').pack(side=tk.LEFT, padx=8)
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
            people = [{'name': utils._person_name(name.get()), 'choice': choice.get().strip()} for _line, name, choice in store]
            if not people or any(not person['name'] for person in people):
                messagebox.showerror('Attendees', 'Enter a name on every row.', parent=self)
                return
            target['people'] = people
            target['decided'] = True
            self._go_next()

        self._actions(card, 'These people are correct', confirm)

        def pull(target=item, store=editors):
            target['people'] = [{'name': utils._person_name(name.get()), 'choice': choice.get().strip()} for _line, name, choice in store]

        return pull

    def _build_names(self, item):
        card = self._card(not item['decided'])
        ttk.Label(card, text='What is the full name?', font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        if item['payer']:
            note = '"{}" was entered on {}\'s purchase without a last name. Enter the first and last name to use.'.format(
                item['written'], item['payer'])
        else:
            note = '"{}" is written this way in Cheddar Up. Your corrections file renames it. Enter the first and last name to use.'.format(
                item['written'])
        self._explain(card, note)
        name = tk.StringVar(value=item.get('name', ''))
        ttk.Entry(card, textvariable=name, width=36).pack(anchor='w')

        def confirm(target=item, var=name):
            full = utils._person_name(var.get())
            if ' ' not in full:
                messagebox.showerror('Name', 'Enter a first and last name.', parent=self)
                return
            target['name'] = full
            target['decided'] = True
            self._go_next()

        self._actions(card, 'Use this name', confirm)

        def pull(target=item, var=name):
            target['name'] = utils._person_name(var.get())

        return pull

    def _build_emails(self, item):
        return self._build_either(
            item,
            'email',
            'Choose this email',
            'Which email should we keep?',
            '{} has two email addresses. Choose the one to keep.'.format(item['guest']),
        )

    def _build_choices(self, item):
        return self._build_either(
            item,
            'choice',
            'Choose this meal',
            'Which meal should we keep?',
            '{} has two meal choices. Choose the one to keep.'.format(item['guest']),
        )

    def _build_either(self, item, field, button_text, title, explanation):
        card = self._card(not item['decided'])
        ttk.Label(card, text=title, font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        self._explain(card, explanation)
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
            self._go_next()

        self._actions(card, button_text, confirm)

        def pull(target=item, var=chosen):
            if var.get():
                target[field] = var.get()

        return pull

    def _build_unclaimed(self, item, row):
        card = self._card(not row['decided'])
        ttk.Label(card, text='Who is this meal for?', font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        self._explain(
            card,
            '{} ordered {}, and Cheddar Up did not put a name on that meal. '
            'Choose the guest it belongs to. Leave the box blank if it should stay unassigned.'.format(
                item['payer'], row['choice'] or 'a meal'),
        )
        names = corrections.preview_names(self.book, self.state)
        guest = tk.StringVar(value=row.get('guest', ''))
        ttk.Combobox(card, textvariable=guest, values=[''] + names, width=36).pack(anchor='w')

        def confirm(target=row, var=guest):
            text = var.get().strip()
            if text and not self._known_name(text):
                messagebox.showerror(
                    'Meal',
                    '"{}" is not on the guest list. Add that person under Missing attendees, or leave the meal unassigned.'.format(text),
                    parent=self,
                )
                return
            target['guest'] = self._known_name(text)
            target['decided'] = True
            self._go_next()

        self._actions(card, 'Use this assignment', confirm)

        def pull(target=row, var=guest):
            target['guest'] = var.get().strip()

        return pull

    def _build_duplicates(self, item):
        left, right = item['pair']
        card = self._card(not item['decided'])
        ttk.Label(card, text='Are these the same person?', font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        self._explain(card, '{} and {} look alike. Say whether to keep both, or which name to keep.'.format(left, right))
        choice = tk.StringVar(value='same' if item.get('decided') and item.get('same_as') else ('separate' if item.get('decided') else ''))
        keep = tk.StringVar(value=item.get('same_as') or left)
        ttk.Radiobutton(card, text='Keep them as two people', value='separate', variable=choice).pack(anchor='w', pady=(6, 0))
        ttk.Radiobutton(card, text='They are the same person', value='same', variable=choice).pack(anchor='w')
        ttk.Label(card, text='Name to keep').pack(anchor='w', pady=(6, 0))
        ttk.Combobox(card, textvariable=keep, values=[left, right], width=32, state='readonly').pack(anchor='w')

        def confirm(target=item, which=choice, kept=keep):
            if which.get() not in ('separate', 'same'):
                messagebox.showerror('Duplicates', 'Choose whether they are the same person.', parent=self)
                return
            target['decided'] = True
            target['same_as'] = kept.get() if which.get() == 'same' else ''
            self._go_next()

        self._actions(card, 'Use this answer', confirm)

        def pull(target=item, which=choice, kept=keep):
            if which.get() == 'same':
                target['same_as'] = kept.get()
            elif which.get() == 'separate':
                target['same_as'] = ''

        return pull

    def _labeled_value(self, parent, label, value):
        ttk.Label(parent, text=label, foreground='#555555').pack(anchor='w', pady=(10, 0))
        self._wrap_label(parent, text=value, font=('Segoe UI', 11, 'bold')).pack(anchor='w', fill=tk.X)

    def _build_seatmates(self, item):
        card = self._card(not item['decided'])
        ttk.Label(card, text='Who should sit together?', font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        self._labeled_value(card, 'Written as', '"{}"'.format(utils._display_phrase(item['raw'])))
        if item['asked_by']:
            self._wrap_label(
                card,
                text='On the purchase for {}'.format(', '.join(item['asked_by'])),
            ).pack(anchor='w', fill=tk.X, pady=(2, 0))
        if item['decided']:
            if item['skip']:
                self._labeled_value(card, 'Matched guest', 'Not a seating request')
            else:
                matched = ', '.join(item['sit_with']) or 'Nobody'
                label = 'Matched guest' if len(item['sit_with']) == 1 else 'Matched guests'
                self._labeled_value(card, label, matched)

            def change(target=item):
                self._open_seatmate(target)
                self._show_node(self._current_node(), pull=False)

            ttk.Button(card, text='Change', command=change).pack(anchor='w', pady=(8, 0))
            ttk.Button(card, text='Skip for now', command=self._skip).pack(anchor='w', pady=(4, 0))
            return lambda: None

        self._explain(
            card,
            'Choose a guest for each person named here, or mark that person not attending. '
            'If the name is not in the list, add that person under Missing attendees first.',
        )
        skip = tk.BooleanVar(value=item.get('skip', False))
        ttk.Checkbutton(card, text='This is not a seating request', variable=skip).pack(anchor='w', pady=(4, 0))
        names = corrections.preview_names(self.book, self.state)
        editors = []
        headers = ttk.Frame(card)
        headers.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(headers, text='Written as', width=28).pack(side=tk.LEFT)
        ttk.Label(headers, text='Matched guest').pack(side=tk.LEFT, padx=8)
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
            self._show_node(self._current_node(), pull=False)

        def confirm(target=item, store=editors, skip_var=skip):
            self._write_seatmate(target, skip_var, store, decide=True)
            if target['decided']:
                self._go_next()

        ttk.Button(card, text='Add a person', command=add_person).pack(anchor='w', pady=(6, 0))
        self._actions(card, 'This is now correct (Save)', confirm)

        def pull(target=item, store=editors, skip_var=skip):
            self._write_seatmate(target, skip_var, store, decide=False)

        return pull

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
                        '"{}" is not on the guest list. Add that person under Missing attendees, or mark them not attending.'.format(part['guest']),
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

    def _build_pairs(self, item):
        card = self._card(not item['decided'])
        kind = 'Who should sit together?' if item['kind'] == 'request' else 'Who should not sit together?'
        ttk.Label(card, text=kind, font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(6, 0))
        self._labeled_value(
            card,
            'Written as',
            '"{}"  /  "{}"'.format(utils._display_phrase(item['left']), utils._display_phrase(item['right'])),
        )
        if item['decided']:
            if item['skip']:
                self._labeled_value(card, 'Matched guests', 'This row will not be used')
            else:
                self._labeled_value(
                    card,
                    'Matched guests',
                    '{}  with  {}'.format(', '.join(item['left_list']), ', '.join(item['right_list'])),
                )

            def change(target=item):
                target['decided'] = False
                self._show_node(self._current_node(), pull=False)

            ttk.Button(card, text='Change', command=change).pack(anchor='w', pady=(8, 0))
            ttk.Button(card, text='Skip for now', command=self._skip).pack(anchor='w', pady=(4, 0))
            return lambda: None

        self._explain(card, 'This row is from the requests workbook. Choose the guests on each side, or skip the row.')
        skip = tk.BooleanVar(value=item.get('skip', False))
        left = tk.StringVar(value='; '.join(item.get('left_list', [])))
        right = tk.StringVar(value='; '.join(item.get('right_list', [])))
        ttk.Checkbutton(card, text='Skip this row', variable=skip).pack(anchor='w')
        ttk.Label(card, text='Left guests').pack(anchor='w', pady=(6, 0))
        names = corrections.preview_names(self.book, self.state)
        ttk.Combobox(card, textvariable=left, values=names, width=36).pack(anchor='w')
        ttk.Label(card, text='Right guests').pack(anchor='w', pady=(6, 0))
        ttk.Combobox(card, textvariable=right, values=names, width=36).pack(anchor='w')
        ttk.Label(card, text='Separate more than one guest with a semicolon.').pack(anchor='w')

        def confirm(target=item, left_var=left, right_var=right, skip_var=skip):
            self._write_pair(target, left_var, right_var, skip_var, decide=True)
            if target['decided']:
                self._go_next()

        self._actions(card, 'This is now correct (Save)', confirm)

        def pull(target=item, left_var=left, right_var=right, skip_var=skip):
            self._write_pair(target, left_var, right_var, skip_var, decide=False)

        return pull

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
        self._write_file(notify=True)

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
            pending = self._first_pending()
            if pending:
                self._show_node(pending, pull=False)
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


def _short(text, limit=48):
    text = ' '.join(str(text).split())
    if len(text) <= limit:
        return text
    return text[:limit - 1] + '…'
