import os
import pickle
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

REQUESTS_WORKBOOK = 'Requests.xlsx'
# Cheddar Up item sheets. Positions, not names: the export renames these every event.
GUEST_TAB_POSITIONS = (5, 6)
SKIP_PAYMENT_STATUSES = {'refunded', 'failed', 'canceled', 'cancelled', 'void', 'voided'}


def _as_text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace('\u00a0', ' ').strip()


def _norm(value):
    """Display form: strip notes in parentheses and collapse whitespace."""
    text = _as_text(value)
    text = re.sub(r'\([^)]*\)', ' ', text)
    text = text.replace('.', ' ').replace(',', ' ')
    text = re.sub(r'\s+', ' ', text).strip(' -')
    return text


def _key(value):
    return _norm(value).casefold()


def _edit_distance(a, b):
    if abs(len(a) - len(b)) > 1:
        return 2
    previous = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        current = [i]
        for j, cb in enumerate(b, 1):
            current.append(min(current[-1] + 1, previous[j] + 1, previous[j - 1] + (ca != cb)))
        previous = current
    return previous[-1]


def _first_names_match(a, b):
    if a == b:
        return True
    short, long = (a, b) if len(a) <= len(b) else (b, a)
    return len(short) >= 3 and long.startswith(short)


def _last_names_match(a, b):
    if a == b:
        return True
    return min(len(a), len(b)) >= 5 and _edit_distance(a, b) <= 1


def _find_header(headers, label):
    want = label.casefold()
    for index, header in enumerate(headers):
        if header.casefold() == want:
            return index
    return None


def _is_attendee_header(header):
    lowered = header.casefold()
    return lowered.startswith('attendee') and 'name' in lowered


def _is_choice_header(header):
    return header.casefold().startswith('choose')


def _is_seatmate_header(header):
    return header.casefold().startswith('requested seatmate')


def _headers_of(worksheet):
    row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
    return [_as_text(cell) for cell in row]


def _attendee_columns(headers):
    """Pair each attendee-name column with the choice column beside it."""
    pairs = []
    index = 0
    while index < len(headers):
        if _is_attendee_header(headers[index]):
            choice_index = None
            if index + 1 < len(headers) and _is_choice_header(headers[index + 1]):
                choice_index = index + 1
            pairs.append((index, choice_index))
            index += 2 if choice_index is not None else 1
        else:
            index += 1
    return pairs


def _at(row, index):
    if index is None or index >= len(row):
        return None
    return row[index]


def _is_zero(value):
    if value is None or value == '':
        return False
    try:
        return float(value) == 0
    except (TypeError, ValueError):
        return False


def chooseExportWorkbook():
    import tkinter
    from tkinter.filedialog import askopenfilename

    print('Select the Cheddar Up collection summary export.')
    print('Guest names are read from the 5th and 6th tabs.')
    root = tkinter.Tk()
    root.withdraw()
    root.overrideredirect(True)
    root.geometry('0x0+0+0')
    root.deiconify()
    root.lift()
    root.focus_force()
    root.attributes('-topmost', True)
    try:
        path = askopenfilename(
            parent=root,
            title='Select the Cheddar Up collection summary',
            initialdir=os.getcwd(),
            filetypes=[('Excel workbook', '*.xlsx'), ('All files', '*.*')],
        )
    finally:
        root.destroy()
    if not path:
        print('No export file selected.')
        sys.exit(1)
    return path


class GuestBook:
    def __init__(self):
        self.guests = {}
        self.by_key = {}
        self.extra = {}
        self.emails = {}
        self.polls = {}
        self.anti = {}
        self.parties = []
        self._party_keys = set()
        self.applied_seatmates = []
        self._applied_keys = set()
        self.unmatched_seatmates = []
        self._seatmate_notes = set()
        self.raw_seatmates = []
        self.raw_pair_rows = []
        self.missing_attendees = []
        self.incomplete_names = []
        self.email_conflicts = []
        self.choice_conflicts = []
        self.unclaimed_choices = []
        self.possible_duplicates = []
        self.placeholder_keys = set()
        self.real_name_keys = set()
        self.linked_requests = []
        self.ignored_requests = []
        self.linked_antis = []
        self.ignored_antis = []
        self.warnings = []
        self._similar_warned = set()

    def add(self, display_name, email, choice):
        display_name = _norm(display_name)
        key = _key(display_name)
        if not key:
            return None
        if key in self.by_key:
            canonical = self.by_key[key]
        else:
            canonical = display_name
            self.by_key[key] = canonical
            self.guests[canonical] = set()
            self._warn_similar(canonical)
        self._fill_email(canonical, email)
        self._fill_choice(canonical, choice)
        return canonical

    def _fill_email(self, canonical, email):
        email = _as_text(email)
        if not email:
            return
        stored_as = formatNameWithoutFamilyName(canonical)
        current = self.emails.get(stored_as, '')
        if current and current.casefold() != email.casefold():
            self.email_conflicts.append({'guest': canonical, 'first': current, 'second': email})
            self.warnings.append(
                '{} has two emails: {!r} and {!r}. Keeping {!r}.'.format(canonical, current, email, current))
            return
        if not current:
            self.emails[stored_as] = email
            if '@' not in email or '.' not in email:
                self.warnings.append('{} has an invalid email address: {}'.format(canonical, email))

    def _fill_choice(self, canonical, choice):
        choice = _norm(choice)
        if not choice:
            return
        current = self.extra.get(canonical, '')
        if current and current.casefold() != choice.casefold():
            self.choice_conflicts.append({'guest': canonical, 'first': current, 'second': choice})
            self.warnings.append(
                '{} has two choices: {!r} and {!r}. Keeping {!r}.'.format(canonical, current, choice, current))
            return
        if not current:
            self.extra[canonical] = choice
            self.polls[formatNameWithoutFamilyName(canonical)] = choice

    def _warn_similar(self, display_name):
        tokens = _key(display_name).split()
        for other in self.guests:
            if other == display_name:
                continue
            other_tokens = _key(other).split()
            short, long = (tokens, other_tokens) if len(tokens) <= len(other_tokens) else (other_tokens, tokens)
            if short and short != long and all(token in long for token in short):
                pair = tuple(sorted((display_name.casefold(), other.casefold())))
                if pair not in self._similar_warned:
                    self._similar_warned.add(pair)
                    ordered = tuple(sorted((display_name, other), key=str.casefold))
                    self.possible_duplicates.append(ordered)
                    self.warnings.append(
                        'Possible duplicate: "{}" and "{}". They will be seated as two people unless you add a request.'.format(
                            display_name, other))

    def link_party(self, party):
        unique = []
        for person in party:
            if person and person not in unique:
                unique.append(person)
        if len(unique) > 1:
            key = frozenset(unique)
            if key not in self._party_keys:
                self._party_keys.add(key)
                self.parties.append(list(unique))
        for person in unique:
            for other in unique:
                if person != other:
                    self.guests[person].add(other)
        return unique

    def match_one(self, text):
        key = _key(text)
        if not key:
            return None
        if key in self.by_key:
            return self.by_key[key]
        tokens = key.split()
        if len(tokens) < 2:
            return None
        first, last = tokens[0], tokens[-1]
        matches = []
        for guest_key, canonical in self.by_key.items():
            guest_tokens = guest_key.split()
            if len(guest_tokens) < 2:
                continue
            if not _last_names_match(last, guest_tokens[-1]):
                continue
            if not _first_names_match(first, guest_tokens[0]):
                continue
            if canonical not in matches:
                matches.append(canonical)
        if len(matches) == 1:
            return matches[0]
        return None

    def explain_people(self, raw):
        """Split a written name into people. 'A and B' can be two people.

        Each item is part, guest, and approximate. guest is None when nobody matched.
        """
        original = _as_text(raw)
        cleaned = _norm(original)
        if not cleaned:
            return []
        whole = self.match_one(cleaned)
        if whole:
            return [{
                'part': original,
                'guest': whole,
                'approximate': _key(whole) != _key(cleaned),
            }]
        pieces = [part.strip() for part in re.split(r'\s+and\s+', cleaned, flags=re.IGNORECASE) if part.strip()]
        if len(pieces) <= 1:
            return [{'part': original, 'guest': None, 'approximate': False}]
        explained = []
        for index, part in enumerate(pieces):
            found = self.match_one(part)
            compared = part
            if found is None and len(part.split()) == 1:
                for other_index, other in enumerate(pieces):
                    if other_index == index:
                        continue
                    other_tokens = other.split()
                    if len(other_tokens) >= 2:
                        candidate = part + ' ' + other_tokens[-1]
                        found = self.match_one(candidate)
                        if found:
                            compared = candidate
                            break
            explained.append({
                'part': part,
                'guest': found,
                'approximate': bool(found) and _key(found) != _key(compared),
            })
        return explained

    def resolve_people(self, raw):
        """Match a seatmate cell to guests. 'A and B' can match two people."""
        explained = self.explain_people(raw)
        matched = []
        unmatched = []
        for item in explained:
            if item['guest']:
                if item['guest'] not in matched:
                    matched.append(item['guest'])
            else:
                unmatched.append(item['part'])
        if matched or unmatched:
            return matched, unmatched
        return [], []

    def last_name_suggestions(self, text):
        tokens = _key(text).split()
        if len(tokens) < 2:
            return []
        last = tokens[-1]
        found = []
        for guest_key, canonical in self.by_key.items():
            guest_tokens = guest_key.split()
            if len(guest_tokens) >= 2 and guest_tokens[-1] == last and canonical not in found:
                found.append(canonical)
        return found

    def record_seatmate_cell(self, party, raw_values):
        for raw in raw_values:
            raw_text = _as_text(raw)
            if raw_text:
                self.raw_seatmates.append((raw_text, list(party)))

    def note_missing_attendee(self, payer, email, choices):
        key = _key(payer)
        for item in self.missing_attendees:
            if _key(item['payer']) == key:
                for choice in choices:
                    if choice and choice not in item['choices']:
                        item['choices'].append(choice)
                return
        self.missing_attendees.append({
            'payer': payer,
            'email': _as_text(email),
            'choices': [choice for choice in choices if choice],
        })

    def link_seatmates(self, party, raw_values):
        for raw in raw_values:
            raw_text = _as_text(raw)
            if not raw_text:
                continue
            matched, unmatched = self.resolve_people(raw)
            for name in matched:
                note_key = (raw_text, name)
                if note_key not in self._applied_keys:
                    self._applied_keys.add(note_key)
                    approximate = _key(name) != _key(raw_text)
                    self.applied_seatmates.append((raw_text, name, approximate))
                for person in party:
                    if person != name:
                        self.guests[person].add(name)
                        self.guests[name].add(person)
            for part in unmatched:
                note = (raw_text, part, tuple(party))
                if note not in self._seatmate_notes:
                    self._seatmate_notes.add(note)
                    self.unmatched_seatmates.append((part, list(party)))

    def link_request_row(self, left_raw, right_raw, kind):
        left_matched, left_unmatched = self.resolve_people(left_raw)
        right_matched, right_unmatched = self.resolve_people(right_raw)
        label = 'REQUEST' if kind == 'request' else 'ANTI-REQUEST'
        if not left_matched or not right_matched or left_unmatched or right_unmatched:
            message = '{} IGNORED: "{}" / "{}" does not match the guest list.'.format(
                label, _as_text(left_raw), _as_text(right_raw))
            if kind == 'request':
                self.ignored_requests.append(message)
            else:
                self.ignored_antis.append(message)
            return
        self._link_people(left_matched, right_matched, kind)

    def _link_people(self, left_matched, right_matched, kind):
        for left in left_matched:
            for right in right_matched:
                if left == right:
                    continue
                if kind == 'request':
                    self.guests[left].add(right)
                    self.guests[right].add(left)
                    self.linked_requests.append((left, right))
                else:
                    self.anti.setdefault(left, set()).add(right)
                    self.anti.setdefault(right, set()).add(left)
                    self.linked_antis.append((left, right))

    def record_pair_row(self, left, right, kind):
        left_text = _as_text(left)
        right_text = _as_text(right)
        if not left_text and not right_text:
            return
        self.raw_pair_rows.append((left_text, right_text, kind))

    def set_email(self, name, email):
        canonical = self.by_key.get(_key(name), _norm(name))
        self.emails[formatNameWithoutFamilyName(canonical)] = _as_text(email)

    def set_choice(self, name, choice):
        choice = _norm(choice)
        if not choice:
            return
        canonical = self.by_key.get(_key(name))
        if not canonical:
            return
        self.extra[canonical] = choice
        self.polls[formatNameWithoutFamilyName(canonical)] = choice

    def remove_guest(self, name):
        canonical = self.by_key.get(_key(name))
        if not canonical:
            return
        self.guests.pop(canonical, None)
        self.by_key.pop(_key(canonical), None)
        self.extra.pop(canonical, None)
        self.anti.pop(canonical, None)
        for members in self.guests.values():
            members.discard(canonical)
        for blocked in self.anti.values():
            blocked.discard(canonical)
        short = formatNameWithoutFamilyName(canonical)
        self.emails.pop(short, None)
        self.polls.pop(short, None)
        parties = []
        for party in self.parties:
            party = [person for person in party if person != canonical]
            if len(party) > 1:
                parties.append(party)
        self.parties = parties
        self._party_keys = {frozenset(party) for party in parties}
        self.raw_seatmates = [
            (raw, [person for person in party if person != canonical]) for raw, party in self.raw_seatmates]
        self.placeholder_keys.discard(_key(canonical))
        self.real_name_keys.discard(_key(canonical))

    def rename_guest(self, old, new):
        new = _norm(new)
        source = self.by_key.get(_key(old))
        if not source or not new:
            return
        target = self.by_key.get(_key(new))
        if target and target != source:
            self.merge_guest(source, target)
            return
        if source == new:
            return
        self._retarget_guest(source, new)

    def merge_guest(self, source, target):
        source = self.by_key.get(_key(source))
        target = self.by_key.get(_key(target))
        if not source or not target or source == target:
            return
        for other in list(self.guests.get(source, ())):
            if other in (source, target):
                continue
            self.guests[target].add(other)
            if other in self.guests:
                self.guests[other].discard(source)
                self.guests[other].add(target)
        if source in self.extra and target not in self.extra:
            self.extra[target] = self.extra[source]
        source_email = self.emails.get(formatNameWithoutFamilyName(source), '')
        if source_email and not self.emails.get(formatNameWithoutFamilyName(target), ''):
            self.emails[formatNameWithoutFamilyName(target)] = source_email
        self.remove_guest(source)

    def _retarget_guest(self, old, new):
        group = self.guests.pop(old)
        self.by_key.pop(_key(old), None)
        self.by_key[_key(new)] = new
        self.guests[new] = group
        for members in self.guests.values():
            if old in members:
                members.discard(old)
                members.add(new)
        if old in self.extra:
            self.extra[new] = self.extra.pop(old)
        if old in self.anti:
            self.anti[new] = self.anti.pop(old)
        for blocked in self.anti.values():
            if old in blocked:
                blocked.discard(old)
                blocked.add(new)
        old_short = formatNameWithoutFamilyName(old)
        new_short = formatNameWithoutFamilyName(new)
        if old_short != new_short:
            if old_short in self.emails:
                self.emails[new_short] = self.emails.pop(old_short)
            if old_short in self.polls:
                self.polls[new_short] = self.polls.pop(old_short)
        self.parties = [[new if person == old else person for person in party] for party in self.parties]
        self._party_keys = {frozenset(party) for party in self.parties if len(party) > 1}
        self.raw_seatmates = [
            (raw, [new if person == old else person for person in party]) for raw, party in self.raw_seatmates]
        if _key(old) in self.placeholder_keys:
            self.placeholder_keys.discard(_key(old))
            self.placeholder_keys.add(_key(new))
        if _key(old) in self.real_name_keys:
            self.real_name_keys.discard(_key(old))
            self.real_name_keys.add(_key(new))

    def finish(self):
        for guest_name in self.guests:
            self.guests[guest_name].add(guest_name)

    def report(self):
        print('\nGuests ({}):'.format(len(self.guests)))
        for name in sorted(self.guests, key=str.casefold):
            choice = self.extra.get(name, '')
            email = self.emails.get(formatNameWithoutFamilyName(name), '')
            print('  {} | {} | {}'.format(name, choice or '-', email or '-'))
        if self.parties:
            print('\nParties from the same purchase:')
            for party in self.parties:
                print('  ' + ', '.join(party))
        if self.applied_seatmates:
            print('\nSeatmate requests applied:')
            for raw, canonical, approximate in self.applied_seatmates:
                suffix = ' (approximate)' if approximate else ''
                print('  {!r} -> {}{}'.format(raw, canonical, suffix))
        if getattr(self, 'skipped_seatmates', None):
            print('\nSeatmate text ignored:')
            for raw in self.skipped_seatmates:
                print('  {!r}'.format(raw))
        if self.unmatched_seatmates:
            print('\nSeatmate names not on the guest list:')
            for part, party in self.unmatched_seatmates:
                with_whom = ', '.join(party) if party else 'unknown party'
                line = '  {!r}  (requested with {})'.format(part, with_whom)
                suggestions = self.last_name_suggestions(part)
                if suggestions:
                    line += '. Guests with that last name: {}'.format(', '.join(suggestions))
                print(line)
        if self.linked_requests:
            print('\nRequests from {}:'.format(getattr(self, 'requests_name', REQUESTS_WORKBOOK)))
            for left, right in self.linked_requests:
                print('  {} with {}'.format(left, right))
        for message in self.ignored_requests:
            print(message)
        if self.linked_antis:
            print('\nAntirequests from {}:'.format(getattr(self, 'requests_name', REQUESTS_WORKBOOK)))
            for left, right in self.linked_antis:
                print('  {} not with {}'.format(left, right))
        for message in self.ignored_antis:
            print(message)
        seen_warnings = set()
        for message in self.warnings:
            if message in seen_warnings:
                continue
            seen_warnings.add(message)
            print('WARNING: ' + message)
        print('\n{} guests, {} emails.'.format(len(self.guests), len(self.emails)))


def _load_workbook(path, label):
    if not os.path.exists(path):
        print('ERROR: {} was not found: {}'.format(label, os.path.abspath(path)))
        sys.exit(1)
    try:
        return load_workbook(path, data_only=True, read_only=True)
    except PermissionError:
        print('ERROR: {} is open in Excel. Close it and run this again.'.format(label))
        sys.exit(1)


def _read_pair_sheet(workbook, sheet_name, book, kind):
    names = workbook.sheetnames
    label = getattr(book, 'requests_name', REQUESTS_WORKBOOK)
    if sheet_name not in names:
        print('ERROR: {} has no "{}" tab. Tabs found: {}'.format(
            label, sheet_name, ', '.join(names)))
        sys.exit(1)
    worksheet = workbook[sheet_name]
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        left = _at(row, 0)
        right = _at(row, 1)
        if not _as_text(left) and not _as_text(right):
            continue
        book.record_pair_row(left, right, kind)


def _read_item_sheet(worksheet, position, book):
    headers = _headers_of(worksheet)
    attendee_pairs = _attendee_columns(headers)
    if not attendee_pairs:
        print('ERROR: Tab {} ({!r}) has no attendee name columns.'.format(position, worksheet.title))
        print('Headers: {}'.format(', '.join(header for header in headers if header)))
        sys.exit(1)

    first_index = _find_header(headers, 'Payer First Name')
    last_index = _find_header(headers, 'Payer Last Name')
    email_index = _find_header(headers, 'Payer Email')
    status_index = _find_header(headers, 'Payment Status')
    net_index = _find_header(headers, 'Net Quantity')
    seatmate_indexes = [index for index, header in enumerate(headers) if _is_seatmate_header(header)]

    row_count = 0
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if not any(_as_text(cell) for cell in row):
            continue
        status = _as_text(_at(row, status_index)).casefold()
        if status in SKIP_PAYMENT_STATUSES:
            continue
        if _is_zero(_at(row, net_index)):
            continue

        email = _at(row, email_index)
        payer = _norm('{} {}'.format(_as_text(_at(row, first_index)), _as_text(_at(row, last_index))))
        named = []
        unnamed_choices = []
        for name_index, choice_index in attendee_pairs:
            name = _norm(_at(row, name_index))
            choice = _norm(_at(row, choice_index))
            if name:
                named.append((name, choice))
            elif choice:
                unnamed_choices.append(choice)

        if not named:
            if not payer:
                continue
            named = [(payer, '')]
            extra = ''
            if unnamed_choices:
                extra = ' Choices with no name: {}.'.format(', '.join(unnamed_choices))
            book.warnings.append('No attendee names for payer "{}". Using the payer name.{}'.format(payer, extra))
            book.note_missing_attendee(payer, email, unnamed_choices)
            book.placeholder_keys.add(_key(payer))
        elif unnamed_choices:
            book.warnings.append(
                '{} has choices with no attendee name: {}.'.format(payer or named[0][0], ', '.join(unnamed_choices)))
            book.unclaimed_choices.append({
                'payer': payer or named[0][0],
                'choices': list(unnamed_choices),
            })

        party = []
        for name, choice in named:
            if ' ' not in name and payer and _key(name) != _key(payer):
                book.incomplete_names.append({'written': name, 'payer': payer, 'choice': choice})
                book.warnings.append(
                    'Attendee "{}" on {}\'s purchase has no last name, so they are seated as "{}".'.format(
                        name, payer, name))
            else:
                book.real_name_keys.add(_key(name))
            canonical = book.add(name, email, choice)
            if canonical and canonical not in party:
                party.append(canonical)
        party = book.link_party(party)
        seatmates = [_at(row, index) for index in seatmate_indexes]
        book.record_seatmate_cell(party, seatmates)
        row_count += 1
    print('Tab {} {!r}: {} purchase rows.'.format(position, worksheet.title, row_count))


def readGuestList(export_path=None, requests_path=None):
    """Read guests from a Cheddar Up export and pairs from a requests workbook.

    The export is chosen each run. Only the 5th and 6th tabs are used.
    The requests workbook has Requests and Antirequests tabs. When no path is
    given, Requests.xlsx in the working folder is used.
    """
    if not export_path:
        export_path = chooseExportWorkbook()
    export_path = os.path.abspath(export_path)
    if not requests_path:
        requests_path = os.path.abspath(REQUESTS_WORKBOOK)
    else:
        requests_path = os.path.abspath(requests_path)
    if os.path.normcase(export_path) == os.path.normcase(requests_path):
        print('The Cheddar Up export and the requests workbook have to be different files.')
        sys.exit(1)

    print('Reading export: {}'.format(export_path))
    book = GuestBook()
    workbook = _load_workbook(export_path, 'collection summary')
    try:
        sheets = workbook.worksheets
        if len(sheets) < max(GUEST_TAB_POSITIONS):
            found = ', '.join('{}. {}'.format(i, sheet.title) for i, sheet in enumerate(sheets, 1))
            print('ERROR: Guest lists are on tabs 5 and 6. This workbook has {} tabs: {}'.format(len(sheets), found))
            sys.exit(1)
        print('Using tabs 5 and 6 only. This workbook has {} tabs.'.format(len(sheets)))
        for position in GUEST_TAB_POSITIONS:
            _read_item_sheet(sheets[position - 1], position, book)
    finally:
        workbook.close()

    print('Reading {}'.format(requests_path))
    book.requests_name = os.path.basename(requests_path)
    requests_book = _load_workbook(requests_path, book.requests_name)
    try:
        _read_pair_sheet(requests_book, 'Requests', book, 'request')
        _read_pair_sheet(requests_book, 'Antirequests', book, 'antirequest')
    finally:
        requests_book.close()

    return book


def save_guest_pickle(book):
    with open('guest_data.p', 'wb') as handle:
        pickle.dump((book.guests, book.anti, book.emails, book.polls), handle)


def formatNameWithoutFamilyName(guest):
    if '-' in guest:
        guest = guest.split('-')[-1].strip()
    return guest


def _ensure_output_dir():
    if not os.path.exists('output'):
        os.makedirs('output')


def writeSeatingChart(tables, extraGuestData, timestamp):
    ## write seating chart into output file
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Seating Chart"

    for col_idx, heading in enumerate(['Table Number', 'Name', 'Number of People', 'Choice'], 1):
        sheet.cell(1, col_idx, heading)

    row_idx = 1
    table_num = 1
    for table in tables:
        for guest in sorted(table):
            row_idx += 1
            sheet.cell(row_idx, 1, table_num)
            sheet.cell(row_idx, 2, formatNameWithoutFamilyName(guest))
            sheet.cell(row_idx, 3, len(table))
            if guest in extraGuestData:
                sheet.cell(row_idx, 4, extraGuestData[guest])
        table_num += 1

    sheet.column_dimensions['A'].width = 16
    sheet.column_dimensions['B'].width = 28
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 18

    output_filename = 'output/Seating Chart' + '_' + timestamp + '.xlsx'
    print('saving seating chart...', output_filename)
    _ensure_output_dir()
    wb.save(output_filename)


def writeTables(tables, emails, polls, timestamp, writeEmails, extraColumn=None):
    wb = Workbook()
    i = 1

    output_filename = 'output/Table Rosters' + '_' + (
        'emails' if writeEmails else 'no emails') + '_' + timestamp + '.xlsx'
    footer_text = os.path.abspath(output_filename).replace('&', '&&')

    guest_font = Font(size=20, bold=True)
    guest_border = Border(bottom=Side(style='dashed'))
    title_font = Font(size=28, bold=True)
    title_alignment = Alignment(horizontal='center', vertical='center')
    header_font = Font(size=20, bold=True)
    header_border = Border(bottom=Side(style='thin'))
    footnote_font = Font(size=9)
    footnote_alignment = Alignment(horizontal='center')

    for table in tables:
        if i == 1:
            sheet = wb.active
        else:
            sheet = wb.create_sheet()
        sheet.title = "Table {}".format(i)
        sheet.page_setup.orientation = 'portrait'
        sheet.oddFooter.center.text = footer_text
        sheet.evenFooter.center.text = footer_text

        if writeEmails:
            sheet.cell(1, 1, 'Table #{} -- ({} people)'.format(i, len(table)))
            sheet.cell(1, 2, 'Email')
            sheet.column_dimensions['A'].width = 58
            sheet.column_dimensions['B'].width = 36
            data_row = 1
        else:
            sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
            title = sheet.cell(1, 1, 'Table #{}'.format(i))
            title.font = title_font
            title.alignment = title_alignment
            sheet.row_dimensions[1].height = 36
            sheet.column_dimensions['A'].width = 39
            sheet.column_dimensions['B'].width = 31

            for col_idx, heading in enumerate(['Name', 'Choice'], 1):
                header = sheet.cell(2, col_idx, heading)
                header.font = header_font
                header.border = header_border
            sheet.row_dimensions[2].height = 26
            data_row = 2

        for guest in sorted(table):
            data_row += 1
            display_name = guest.split('-')[-1].strip() if '-' in guest else guest
            name_cell = sheet.cell(data_row, 1, display_name)
            if not writeEmails:
                name_cell.font = guest_font
                name_cell.border = guest_border
                sheet.row_dimensions[data_row].height = 26
            try:
                if writeEmails:
                    sheet.cell(data_row, 2, emails[formatNameWithoutFamilyName(guest)])
                else:
                    choice_cell = sheet.cell(data_row, 2, extraColumn[guest])
                    choice_cell.font = guest_font
                    choice_cell.border = guest_border
            except:
                if writeEmails:
                    sheet.cell(data_row, 2, '')

        if not writeEmails:
            footnote_row = data_row + 1
            sheet.merge_cells(start_row=footnote_row, start_column=1, end_row=footnote_row, end_column=2)
            footnote = sheet.cell(footnote_row, 1, '({} people)'.format(len(table)))
            footnote.font = footnote_font
            footnote.alignment = footnote_alignment

        i += 1
    print('saving table rosters...', output_filename)
    _ensure_output_dir()
    wb.save(output_filename)
