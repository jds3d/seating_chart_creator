# -*- coding: utf-8 -*-
"""Load, save, and apply seating corrections.

Corrections.xlsx is the editable copy. Each sheet is one kind of decision.
"""

import os
from contextlib import contextmanager

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

import utils

CORRECTIONS_WORKBOOK = 'Corrections.xlsx'

_HEADER_FILL = PatternFill('solid', fgColor='1F4E79')
_HEADER_FONT = Font(bold=True, color='FFFFFF')

_SHEETS = {
    'missing': ('Missing attendees', ['Payer', 'Person', 'Choice']),
    'names': ('Names', ['As written', 'Correct name']),
    'emails': ('Emails', ['Guest', 'Email']),
    'choices': ('Meal choices', ['Guest', 'Choice']),
    'unclaimed': ('Unclaimed choices', ['Payer', 'Choice', 'Guest']),
    'duplicates': ('Duplicates', ['Name', 'Same as']),
    'seatmates': ('Seatmates', ['Written as', 'Sit with', 'Skip']),
    'pairs': ('Request pairs', ['Kind', 'Left as written', 'Right as written', 'Left guests', 'Right guests', 'Skip']),
}


def corrections_path(folder):
    return os.path.join(folder, CORRECTIONS_WORKBOOK)


def load_corrections(path):
    loaded = {key: [] for key in _SHEETS}
    if not path or not os.path.isfile(path):
        return loaded
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
    except PermissionError:
        raise
    try:
        for key, (title, headers) in _SHEETS.items():
            if title not in workbook.sheetnames:
                continue
            worksheet = workbook[title]
            rows = worksheet.iter_rows(values_only=True)
            header = [utils._as_text(cell) for cell in next(rows, [])]
            index = {name: position for position, name in enumerate(header)}
            if any(name not in index for name in headers):
                continue
            for row in rows:
                values = [_cell(row, index[name]) for name in headers]
                if not any(values):
                    continue
                loaded[key].append(dict(zip(headers, values)))
    finally:
        workbook.close()
    return loaded


def save_corrections(path, state, previous):
    workbook = Workbook()
    notes = workbook.active
    notes.title = 'Read me'
    lines = [
        'Seating Chart Creator saves decisions in this workbook.',
        'You can edit the sheets in Excel. Leave the header row in place.',
        'Separate more than one guest with a semicolon.',
        'Put yes in a Skip column when the text is not a seating request.',
        'Missing attendees: one row per person who should be seated when a payer left the names blank.',
        'Missing attendees, also: when an order paid for more tickets than it names people, one row per extra person '
        '(Payer, Person, Choice). Those people join the people already named on that order.',
        'Names: replace a short name with the full name, or correct a full name as written in Cheddar Up '
        '(for example a misspelling). Matching ignores case, periods, commas and text in parentheses.',
        'Emails and Meal choices: the value to keep.',
        'Seatmates: who a written request should sit with.',
        'Request pairs: how a row from the requests workbook should be matched.',
        'Duplicates: put the name to keep in Same as. Leave Same as blank to keep both people.',
    ]
    for row, line in enumerate(lines, 1):
        notes.cell(row, 1, line)
    notes.column_dimensions['A'].width = 110

    tables = {
        'missing': _missing_rows(state, previous),
        'names': _name_rows(state, previous),
        'emails': _email_rows(state, previous),
        'choices': _choice_rows(state, previous),
        'unclaimed': _unclaimed_rows(state, previous),
        'duplicates': _duplicate_rows(state, previous),
        'seatmates': _seatmate_rows(state, previous),
        'pairs': _pair_rows(state, previous),
    }
    for key, (title, headers) in _SHEETS.items():
        worksheet = workbook.create_sheet(title)
        for column, header in enumerate(headers, 1):
            cell = worksheet.cell(1, column, header)
            cell.fill = _HEADER_FILL
            cell.font = _HEADER_FONT
            worksheet.column_dimensions[utils_column(column)].width = 36 if column == 1 else 42
        for row_number, values in enumerate(tables[key], 2):
            for column, value in enumerate(values, 1):
                worksheet.cell(row_number, column, value)
        worksheet.freeze_panes = 'A2'
        worksheet.auto_filter.ref = 'A1:{}1'.format(utils_column(len(headers)))
    workbook.save(path)


def build_state(book, loaded):
    missing = _build_missing(book, loaded)
    state = {
        'missing': missing,
        'names': _build_names(book, loaded),
        'emails': _build_emails(book, loaded),
        'choices': _build_simple_conflicts(book.choice_conflicts, loaded.get('choices', []), 'Choice'),
        'unclaimed': _build_unclaimed(book, loaded, skip=_short_payers(missing)),
        'duplicates': _build_duplicates(book, loaded),
        'seatmates': _build_seatmates(book, loaded),
        'pairs': _build_pairs(book, loaded),
    }
    return state


def section_defs(state):
    titles = [
        ('missing', 'Missing attendees'),
        ('names', 'Names to correct'),
        ('emails', 'Two emails'),
        ('choices', 'Two meal choices'),
        ('unclaimed', 'Choices with no name'),
        ('duplicates', 'Possible duplicates'),
        ('seatmates', 'Seatmate requests'),
        ('pairs', 'Requests workbook'),
    ]
    found = []
    for key, title in titles:
        rows = state[key]
        if not rows:
            continue
        found.append((key, title, _pending(key, rows)))
    return found


def _pending(key, rows):
    if key == 'unclaimed':
        return sum(1 for item in rows.values() if any(not row['decided'] for row in item['rows']))
    return sum(1 for item in rows.values() if not item.get('decided'))


def has_sections(book):
    return bool(section_defs(build_state(book, {key: [] for key in _SHEETS})))


def preview_names(book, state):
    names = set(book.guests)
    for item in state['names'].values():
        written = item['written']
        correct = utils._person_name(item.get('name', ''))
        if not correct:
            continue
        names = {name for name in names if utils._key(name) != utils._key(written)}
        names.add(correct)
    for payer, item in state['missing'].items():
        entered = [utils._person_name(person.get('name', '')) for person in item['people']]
        entered = [name for name in entered if name]
        for name in entered:
            names.add(name)
        payer_key = utils._key(payer)
        if item.get('kind') != 'short' and payer_key in book.placeholder_keys and payer_key not in book.real_name_keys:
            if not any(utils._key(name) == payer_key for name in entered):
                names = {name for name in names if utils._key(name) != payer_key}
    for pair, item in state['duplicates'].items():
        if item.get('decided') and utils._person_name(item.get('same_as', '')):
            keep = utils._key(item['same_as'])
            drop = pair[0] if utils._key(pair[1]) == keep else pair[1]
            names = {name for name in names if utils._key(name) != utils._key(drop)}
    return sorted(names, key=str.casefold)


def refresh_open_seatmates(book, state):
    extra = preview_names(book, state)
    for item in state['seatmates'].values():
        if item['from_file'] or item['touched'] or item['decided']:
            continue
        fresh = _seatmate_guess(book, item['raw'], item['asked_by'], extra)
        item['parts'] = fresh['parts']
        item['sit_with'] = fresh['sit_with']
        item['decided'] = fresh['decided']


@contextmanager
def _temporary_names(book, names):
    added = []
    for name in names:
        name = utils._person_name(name)
        key = utils._key(name)
        if not key or key in book.by_key:
            continue
        book.by_key[key] = name
        book.guests[name] = set()
        added.append(name)
    try:
        yield
    finally:
        for name in added:
            book.guests.pop(name, None)
            if book.by_key.get(utils._key(name)) == name:
                book.by_key.pop(utils._key(name), None)


def validate(book, state):
    known = {utils._key(name): name for name in preview_names(book, state)}
    errors = []
    for payer, item in state['missing'].items():
        if not item['decided']:
            errors.append('Confirm the people for {}.'.format(payer))
            continue
        if not any(utils._person_name(person.get('name', '')) for person in item['people']):
            errors.append('Enter at least one person for {}.'.format(payer))
        elif any(not utils._person_name(person.get('name', '')) for person in item['people']):
            errors.append('Enter a name on every row for {}.'.format(payer))
    for item in state['names'].values():
        if not item['decided']:
            errors.append('Confirm the full name for {}.'.format(item['written']))
        elif ' ' not in utils._person_name(item.get('name', '')):
            errors.append('Give {} a first and last name.'.format(item['written']))
    for guest, item in state['emails'].items():
        if not item['decided']:
            errors.append('Choose an email for {}.'.format(guest))
    for guest, item in state['choices'].items():
        if not item['decided']:
            errors.append('Choose a meal for {}.'.format(guest))
    for payer, item in state['unclaimed'].items():
        for row in item['rows']:
            if not row['decided']:
                errors.append('Assign the {} choice on {}\'s purchase, or leave it unassigned and confirm it.'.format(
                    row['choice'], payer))
    for pair, item in state['duplicates'].items():
        if not item['decided']:
            errors.append('Say whether {} and {} are the same person.'.format(pair[0], pair[1]))
    for item in state['seatmates'].values():
        if not item['decided']:
            errors.append('Finish the seatmate request "{}".'.format(item['raw']))
            continue
        if item['skip']:
            continue
        if not item['sit_with']:
            errors.append('Choose who "{}" sits with, or mark it as not a seating request.'.format(item['raw']))
            continue
        for name in item['sit_with']:
            if utils._key(name) not in known:
                errors.append('"{}" is not on the guest list. Add that person under Missing attendees, or choose someone else.'.format(name))
    for item in state['pairs'].values():
        if not item['decided']:
            errors.append('Match the request "{}" / "{}".'.format(item['left'], item['right']))
            continue
        if item['skip']:
            continue
        for name in item['left_list'] + item['right_list']:
            if utils._key(name) not in known:
                errors.append('"{}" is not on the guest list.'.format(name))
        if not item['left_list'] or not item['right_list']:
            errors.append('Choose a guest on both sides of "{}" / "{}", or skip it.'.format(item['left'], item['right']))
    return errors


def apply_all(book, state):
    _apply_duplicates(book, state)
    _apply_names(book, state)
    _apply_missing(book, state)
    _apply_emails(book, state)
    _apply_choices(book, state)
    _apply_unclaimed(book, state)
    _apply_seatmates(book, state)
    _apply_pairs(book, state)


def _apply_duplicates(book, state):
    for pair, item in state['duplicates'].items():
        if not item.get('decided'):
            continue
        keep = utils._person_name(item.get('same_as', ''))
        if not keep:
            continue
        drop = pair[0] if utils._key(pair[1]) == utils._key(keep) else pair[1]
        book.merge_guest(drop, keep)


def _apply_names(book, state):
    for item in state['names'].values():
        if item.get('decided') and utils._person_name(item.get('name', '')):
            book.rename_guest(item['written'], item['name'])


def _apply_missing(book, state):
    for payer, item in state['missing'].items():
        if not item.get('decided'):
            continue
        email = item.get('email', '')
        for issue in book.missing_attendees:
            if utils._key(issue['payer']) == utils._key(payer):
                email = issue['email']
                break
        people = []
        for person in item['people']:
            name = utils._person_name(person.get('name', ''))
            if not name:
                continue
            canonical = book.add(name, email, '')
            if person.get('choice'):
                book.set_choice(canonical, person['choice'])
            if canonical not in people:
                people.append(canonical)
        if item.get('kind') == 'short':
            # The order already names some people. The new ones join that party.
            named = [book.by_key[utils._key(name)] for name in item.get('named', []) if utils._key(name) in book.by_key]
            if len(people) + len(named) > 1:
                book.link_party(named + [person for person in people if person not in named])
            continue
        if len(people) > 1:
            book.link_party(people)
        payer_key = utils._key(payer)
        kept = any(utils._key(person) == payer_key for person in people)
        if not kept and payer_key in book.placeholder_keys and payer_key not in book.real_name_keys:
            book.remove_guest(payer)


def _apply_emails(book, state):
    for guest, item in state['emails'].items():
        if item.get('decided') and item.get('email'):
            book.set_email(guest, item['email'])


def _apply_choices(book, state):
    for guest, item in state['choices'].items():
        if item.get('decided') and item.get('choice'):
            book.set_choice(guest, item['choice'])


def _apply_unclaimed(book, state):
    for item in state['unclaimed'].values():
        for row in item['rows']:
            if row.get('decided') and row.get('guest') and row.get('choice'):
                book.set_choice(row['guest'], row['choice'])


def _apply_seatmates(book, state):
    book.skipped_seatmates = []
    parties = {}
    for raw, party in book.raw_seatmates:
        parties.setdefault(raw, []).append(party)
    for item in state['seatmates'].values():
        if not item.get('decided'):
            continue
        raw = item['raw']
        if item.get('skip'):
            book.skipped_seatmates.append(raw)
            continue
        chosen = []
        for name in item['sit_with']:
            canonical = book.by_key.get(utils._key(name))
            if canonical and canonical not in chosen:
                chosen.append(canonical)
        groups = parties.get(raw) or [[]]
        for party in groups:
            people = []
            for person in list(party) + chosen:
                if person and person not in people:
                    people.append(person)
            for person in people:
                for other in people:
                    if person != other and person in book.guests and other in book.guests:
                        book.guests[person].add(other)
            for guest in chosen:
                note_key = (raw, guest)
                if note_key not in book._applied_keys:
                    book._applied_keys.add(note_key)
                    book.applied_seatmates.append((raw, guest, utils._key(guest) != utils._key(raw)))


def _apply_pairs(book, state):
    for item in state['pairs'].values():
        if item.get('decided'):
            if item.get('skip'):
                continue
            left = [book.by_key[utils._key(name)] for name in item['left_list'] if utils._key(name) in book.by_key]
            right = [book.by_key[utils._key(name)] for name in item['right_list'] if utils._key(name) in book.by_key]
            if left and right:
                book._link_people(left, right, item['kind'])
            continue
        if item['left'] and item['right']:
            book.link_request_row(item['left'], item['right'], item['kind'])


def _build_missing(book, loaded):
    saved = {}
    for row in loaded.get('missing', []):
        saved.setdefault(utils._key(row['Payer']), []).append(row)
    items = {}
    for issue in book.missing_attendees:
        payer = issue['payer']
        rows = saved.get(utils._key(payer))
        if rows:
            people = [{'name': utils._person_name(row['Person']), 'choice': row['Choice']} for row in rows]
            decided = True
        else:
            choices = issue['choices'] or ['']
            people = [{'name': utils._person_name(payer) if index == 0 else '', 'choice': choice} for index, choice in enumerate(choices)]
            decided = False
        items[utils._person_name(payer)] = {
            'payer': utils._person_name(payer),
            'email': issue['email'],
            'choices': list(issue['choices']),
            'people': people,
            'decided': decided,
            'kind': 'blank',
            'named': [],
            'extra': 0,
        }
    for order in _short_order_issues(book):
        payer = utils._person_name(order['payer'])
        if utils._key(payer) in {utils._key(name) for name in items}:
            continue
        rows = saved.get(utils._key(payer))
        if rows:
            people = [{'name': utils._person_name(row['Person']), 'choice': row['Choice']} for row in rows]
            decided = True
        else:
            choices = order['choices']
            people = [{'name': '', 'choice': choices[index] if index < len(choices) else ''}
                      for index in range(max(order['extra'], len(choices)))]
            decided = False
        items[payer] = {
            'payer': payer,
            'email': order['email'],
            'choices': list(order['choices']),
            'people': people,
            'decided': decided,
            'kind': 'short',
            'named': list(order['names']),
            'extra': order['extra'],
            'tickets': order['tickets'],
        }
    return items


def _short_order_issues(book):
    """One entry per payer whose orders paid for more tickets than they name people."""
    found = {}
    for order in book.short_orders():
        payer = utils._person_name(order['payer'])
        if not payer:
            continue
        key = utils._key(payer)
        if key not in found:
            found[key] = dict(order, payer=payer, names=list(order['names']), choices=list(order['choices']))
            continue
        merged = found[key]
        merged['tickets'] += order['tickets']
        merged['extra'] += order['extra']
        merged['names'] += [name for name in order['names'] if name not in merged['names']]
        merged['choices'] += [choice for choice in order['choices'] if choice not in merged['choices']]
    return list(found.values())


def _short_payers(missing):
    return {utils._key(payer) for payer, item in missing.items() if item.get('kind') == 'short'}


def _build_names(book, loaded):
    saved = {utils._key(row['As written']): row['Correct name'] for row in loaded.get('names', [])}
    items = {}
    for issue in book.incomplete_names:
        written = issue['written']
        key = utils._key(written)
        if key in items:
            continue
        if key in saved and saved[key]:
            name = saved[key]
            decided = True
        else:
            payer_tokens = utils._key(issue['payer']).split()
            name = written
            if payer_tokens and utils._key(written) != payer_tokens[-1]:
                name = '{} {}'.format(written, issue['payer'].split()[-1])
            decided = False
        name = utils._person_name(name)
        items[key] = {
            'written': utils._person_name(written),
            'payer': utils._person_name(issue['payer']),
            'name': name,
            'decided': decided,
        }
    # A full name written in Cheddar Up can also be corrected here.
    for key, correct in saved.items():
        if key in items or not correct or key not in book.by_key:
            continue
        written = book.by_key[key]
        items[key] = {
            'written': written,
            'payer': '',
            'name': utils._person_name(correct),
            'decided': True,
        }
    return items


def _build_emails(book, loaded):
    saved = {utils._key(row['Guest']): row['Email'] for row in loaded.get('emails', [])}
    items = {}
    for issue in book.email_conflicts:
        guest = issue['guest']
        key = utils._key(guest)
        if key in items:
            continue
        chosen = saved.get(key, '')
        items[key] = {
            'guest': guest,
            'first': issue['first'],
            'second': issue['second'],
            'email': chosen,
            'decided': bool(chosen),
        }
    return items


def _build_simple_conflicts(conflicts, rows, value_header):
    saved = {utils._key(row['Guest']): row[value_header] for row in rows}
    items = {}
    for issue in conflicts:
        guest = issue['guest']
        key = utils._key(guest)
        if key in items:
            continue
        chosen = saved.get(key, '')
        items[key] = {
            'guest': guest,
            'first': issue['first'],
            'second': issue['second'],
            'choice': chosen,
            'decided': bool(chosen),
        }
    return items


def _build_unclaimed(book, loaded, skip=()):
    saved = {}
    for row in loaded.get('unclaimed', []):
        saved.setdefault(utils._key(row['Payer']), []).append(row)
    items = {}
    for issue in book.unclaimed_choices:
        payer = issue['payer']
        key = utils._key(payer)
        if key in skip:
            # The Missing attendees question for this order already asks for these meals.
            continue
        if key in items:
            for choice in issue['choices']:
                if choice not in [row['choice'] for row in items[key]['rows']]:
                    items[key]['rows'].append({'choice': choice, 'guest': '', 'decided': False})
            continue
        rows = saved.get(key)
        if rows:
            built = [{'choice': row['Choice'], 'guest': row['Guest'], 'decided': True} for row in rows]
        else:
            built = [{'choice': choice, 'guest': '', 'decided': False} for choice in issue['choices']]
        items[payer] = {'payer': payer, 'rows': built}
    return items


def _build_duplicates(book, loaded):
    saved = {}
    for row in loaded.get('duplicates', []):
        saved[utils._key(row['Name'])] = row['Same as']
    items = {}
    for left, right in book.possible_duplicates:
        key = tuple(sorted((left, right), key=str.casefold))
        if key in items:
            continue
        same = saved.get(utils._key(left), None)
        if same is None:
            same = saved.get(utils._key(right), None)
        items[key] = {
            'pair': key,
            'same_as': same or '',
            'decided': same is not None,
        }
    return items


def _build_seatmates(book, loaded):
    saved = {}
    for row in loaded.get('seatmates', []):
        saved[row['Written as']] = row
    grouped = {}
    order = []
    for raw, party in book.raw_seatmates:
        if raw not in grouped:
            grouped[raw] = []
            order.append(raw)
        for person in party:
            if person not in grouped[raw]:
                grouped[raw].append(person)
    items = {}
    for raw in order:
        if raw in saved:
            row = saved[raw]
            item = {
                'raw': raw,
                'asked_by': grouped[raw],
                'sit_with': _split_names(row['Sit with']),
                'skip': _is_yes(row['Skip']),
                'parts': [],
                'decided': True,
                'touched': True,
                'from_file': True,
            }
        else:
            item = _seatmate_guess(book, raw, grouped[raw], [])
        items[raw] = item
    return items


def _seatmate_guess(book, raw, asked_by, extra_names):
    with _temporary_names(book, extra_names):
        explained = book.explain_people(raw)
    parts = []
    sit_with = []
    for item in explained:
        guest = item['guest'] or ''
        parts.append({
            'part': item['part'],
            'guest': guest,
            'absent': False,
            'approximate': item['approximate'],
        })
        if guest and guest not in sit_with:
            sit_with.append(guest)
    needs_answer = any(not part['guest'] for part in parts) or any(part['approximate'] for part in parts)
    return {
        'raw': raw,
        'asked_by': asked_by,
        'sit_with': sit_with,
        'skip': False,
        'parts': parts,
        'decided': not needs_answer,
        'touched': False,
        'from_file': False,
    }


def _build_pairs(book, loaded):
    saved = {}
    for row in loaded.get('pairs', []):
        saved[(row['Kind'], row['Left as written'], row['Right as written'])] = row
    items = {}
    for left, right, kind in book.raw_pair_rows:
        key = (kind, left, right)
        if key in saved:
            row = saved[key]
            items[key] = {
                'kind': kind,
                'left': left,
                'right': right,
                'left_list': _split_names(row['Left guests']),
                'right_list': _split_names(row['Right guests']),
                'skip': _is_yes(row['Skip']),
                'decided': True,
            }
            continue
        left_matched, left_unmatched = book.resolve_people(left)
        right_matched, right_unmatched = book.resolve_people(right)
        resolved = bool(left_matched and right_matched and not left_unmatched and not right_unmatched and left and right)
        items[key] = {
            'kind': kind,
            'left': left,
            'right': right,
            'left_list': left_matched,
            'right_list': right_matched,
            'skip': False,
            'decided': resolved,
        }
    return items


def _missing_rows(state, previous):
    rows = []
    seen = set()
    for payer, item in state['missing'].items():
        if not item.get('decided'):
            continue
        seen.add(utils._key(payer))
        for person in item['people']:
            name = utils._person_name(person.get('name', ''))
            if name:
                rows.append((utils._person_name(payer), name, person.get('choice', '')))
    for row in previous.get('missing', []):
        if utils._key(row['Payer']) not in seen:
            rows.append((row['Payer'], row['Person'], row['Choice']))
    return rows


def _name_rows(state, previous):
    rows = []
    seen = set()
    for item in state['names'].values():
        if not item.get('decided'):
            continue
        seen.add(utils._key(item['written']))
        rows.append((utils._person_name(item['written']), utils._person_name(item.get('name', ''))))
    for row in previous.get('names', []):
        if utils._key(row['As written']) not in seen:
            rows.append((row['As written'], row['Correct name']))
    return rows


def _email_rows(state, previous):
    rows = []
    seen = set()
    for item in state['emails'].values():
        if not item.get('decided'):
            continue
        seen.add(utils._key(item['guest']))
        rows.append((item['guest'], item.get('email', '')))
    for row in previous.get('emails', []):
        if utils._key(row['Guest']) not in seen:
            rows.append((row['Guest'], row['Email']))
    return rows


def _choice_rows(state, previous):
    rows = []
    seen = set()
    for item in state['choices'].values():
        if not item.get('decided'):
            continue
        seen.add(utils._key(item['guest']))
        rows.append((item['guest'], item.get('choice', '')))
    for row in previous.get('choices', []):
        if utils._key(row['Guest']) not in seen:
            rows.append((row['Guest'], row['Choice']))
    return rows


def _unclaimed_rows(state, previous):
    rows = []
    seen = set()
    for payer, item in state['unclaimed'].items():
        if not all(row.get('decided') for row in item['rows']):
            continue
        seen.add(utils._key(payer))
        for row in item['rows']:
            rows.append((payer, row['choice'], row.get('guest', '')))
    for row in previous.get('unclaimed', []):
        if utils._key(row['Payer']) not in seen:
            rows.append((row['Payer'], row['Choice'], row['Guest']))
    return rows


def _duplicate_rows(state, previous):
    rows = []
    seen = set()
    for pair, item in state['duplicates'].items():
        if not item.get('decided'):
            continue
        seen.add(utils._key(pair[0]))
        rows.append((pair[0], item.get('same_as', '')))
    for row in previous.get('duplicates', []):
        if utils._key(row['Name']) not in seen:
            rows.append((row['Name'], row['Same as']))
    return rows


def _seatmate_rows(state, previous):
    rows = []
    seen = set()
    for item in state['seatmates'].values():
        if not item.get('decided'):
            continue
        seen.add(item['raw'])
        rows.append((item['raw'], '' if item.get('skip') else '; '.join(item['sit_with']), 'yes' if item.get('skip') else ''))
    for row in previous.get('seatmates', []):
        if row['Written as'] not in seen:
            rows.append((row['Written as'], row['Sit with'], row['Skip']))
    return rows


def _pair_rows(state, previous):
    rows = []
    seen = set()
    for item in state['pairs'].values():
        if not item.get('decided'):
            continue
        seen.add((item['kind'], item['left'], item['right']))
        rows.append((
            item['kind'],
            item['left'],
            item['right'],
            '' if item.get('skip') else '; '.join(item['left_list']),
            '' if item.get('skip') else '; '.join(item['right_list']),
            'yes' if item.get('skip') else '',
        ))
    for row in previous.get('pairs', []):
        key = (row['Kind'], row['Left as written'], row['Right as written'])
        if key not in seen:
            rows.append((row['Kind'], row['Left as written'], row['Right as written'], row['Left guests'], row['Right guests'], row['Skip']))
    return rows


def _cell(row, index):
    if index >= len(row):
        return ''
    return utils._as_text(row[index])


def _split_names(value):
    text = utils._as_text(value)
    if not text:
        return []
    return [utils._person_name(part) for part in text.split(';') if utils._person_name(part)]


def _is_yes(value):
    return utils._as_text(value).casefold() in ('yes', 'y', 'true', '1')


def utils_column(number):
    name = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name
