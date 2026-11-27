# -*- coding: utf-8 -*-
"""Build Seating Chart Creator.exe in this project folder.

From this project folder:

    python package_exe.py

The exe is written next to Requests.xlsx and Corrections.xlsx.
Those workbooks stay where they are, so a git pull can replace the exe
without resetting corrections. Charts are written to output\\.
"""

import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
VENV_PYTHON = os.path.join(ROOT, '.venv', 'Scripts', 'python.exe')
APP_NAME = 'Seating Chart Creator'
EXE_PATH = os.path.join(ROOT, APP_NAME + '.exe')
HOW_TO = """Double-click "Seating Chart Creator.exe".

Keep Requests.xlsx in this same folder. The program reads it every time you make a chart.

To get the newest Cheddar Up export, click "Download latest from Cheddar Up".
Cheddar Up opens in your browser. Sign in if it asks, open the TVNJC collection, then Manage > Reports > Collection Summary.
Leave the file in Downloads. The window picks it up as soon as it finishes, copies it into the "Cheddar Up exports" folder, and offers to make the seating chart.
If the export is already in Downloads, the button offers to use it right away.

You can still click Browse to pick a Cheddar Up file yourself. The window keeps the last one you used.

Finished seating charts and table rosters are saved in the output folder.

When a name or seatmate request needs a decision, the window asks you to fix it.
Those answers are saved in Corrections.xlsx, in this same folder. You can edit that file in Excel.
Updating the program does not replace Corrections.xlsx, Requests.xlsx, or the saved Cheddar Up file.

To get a newer program, close Seating Chart Creator, then update this folder with git pull.
Then double-click the exe again.

Python does not need to be installed.
"""


def _same_file(left, right):
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


def _relaunch_in_venv():
    # os.execv splits this path at the space in "Personal Code" on Windows.
    command = [VENV_PYTHON, os.path.abspath(__file__), *sys.argv[1:]]
    raise SystemExit(subprocess.call(command))


def _use_project_venv():
    if os.path.isfile(VENV_PYTHON) and _same_file(sys.executable, VENV_PYTHON):
        return
    if not os.path.isfile(VENV_PYTHON):
        print('Creating .venv with Python 3.13')
        subprocess.check_call(['py', '-3.13', '-m', 'venv', os.path.join(ROOT, '.venv')])
    _relaunch_in_venv()


def _install_build_tools():
    subprocess.check_call([
        sys.executable, '-m', 'pip', 'install', '-r', os.path.join(ROOT, 'requirements.txt'), 'pyinstaller',
    ])


def _build_exe():
    dist = os.path.join(ROOT, 'dist')
    work = os.path.join(ROOT, 'build')
    seating = os.path.join(ROOT, 'Seating Chart Creator.py')
    rosters = os.path.join(ROOT, 'Create Table Rosters From Seating Chart.py')
    command = [
        sys.executable, '-m', 'PyInstaller',
        '--noconfirm',
        '--clean',
        '--windowed',
        '--onefile',
        '--noupx',
        '--name', APP_NAME,
        '--distpath', dist,
        '--workpath', work,
        '--specpath', work,
        '--add-data', seating + os.pathsep + '.',
        '--add-data', rosters + os.pathsep + '.',
        '--collect-all', 'openpyxl',
        os.path.join(ROOT, 'gui.py'),
    ]
    subprocess.check_call(command)
    exe = os.path.join(dist, APP_NAME + '.exe')
    if not os.path.isfile(exe):
        raise SystemExit('PyInstaller finished without writing {}'.format(exe))
    return exe


def _place_exe(exe_path):
    """Copy the new exe beside the workbooks. Do not replace those workbooks."""
    os.makedirs(os.path.join(ROOT, 'output'), exist_ok=True)
    requests = os.path.join(ROOT, 'Requests.xlsx')
    template = os.path.join(ROOT, 'Requests_Template.xlsx')
    if not os.path.isfile(requests):
        if os.path.isfile(template):
            shutil.copy2(template, requests)
            print('Requests.xlsx was missing, so the template was copied into the project folder.')
        else:
            raise SystemExit('Put Requests.xlsx in the project folder before packaging.')
    if not _same_file(exe_path, EXE_PATH):
        try:
            shutil.copy2(exe_path, EXE_PATH)
        except PermissionError:
            raise SystemExit(
                'Close "Seating Chart Creator.exe" and run this again. '
                'The new exe was built, but the one in this folder is still open.'
            )
    with open(os.path.join(ROOT, 'How to run.txt'), 'w', encoding='utf-8') as instructions:
        instructions.write(HOW_TO)


def main():
    os.chdir(ROOT)
    _use_project_venv()
    print('Installing openpyxl and PyInstaller into .venv')
    _install_build_tools()
    print('Building the exe. This takes a few minutes.')
    exe_path = _build_exe()
    _place_exe(exe_path)
    print()
    print('Ready: {}'.format(EXE_PATH))
    print('Commit that exe. A git pull replaces it and leaves Corrections.xlsx and Requests.xlsx alone.')


if __name__ == '__main__':
    main()
