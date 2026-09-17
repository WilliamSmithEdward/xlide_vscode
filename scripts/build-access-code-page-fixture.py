#!/usr/bin/env python
"""Builds tests/fixtures/binaries/AccessCodePageFixture.accdb.

    python scripts/build-access-code-page-fixture.py <blank.accdb> <out.accdb>

A form whose controls are named outside ASCII, which is what shows the code
page a design's TypeInfo stream is written in, and what Access does with a
name that page cannot hold.

What it holds, on form Names:

  Caf<e-acute>         stored as 43 61 66 E9
  Em<em dash>Dash      stored with 97
  <euro>uro            stored with 80, which is cp1252 (cp1251 has it at 88)
  Na<i-diaeresis>ve <em dash> x   both names: underscores for its spaces
  <three Cyrillic letters>        NO entry, and no ordinal spent on it
  Plain

Measured on Access 16.0 on a machine whose ANSI code page is 1252. The names
are single bytes in that page. A name with even one character the page cannot
hold exactly gets no entry at all, and no best fit is tried: an A-macron, a
fullwidth A and a Greek omega are all left out rather than listed as A or O.
The same bytes come out of a database created with the Cyrillic collation, and
of one whose PROJECTCODEPAGE was patched to 1251 first: VBA went on reading the
project as cp1252, and Access wrote 1252 back on its next save. So on one
machine the page is the machine's, whatever the file says. What a machine with
another ANSI page writes could not be measured here.

Safety: works on a scratch copy of the blank database, in a temporary folder,
in an Access instance it starts itself. It refuses to run while Access is
already running, and quits through COM.

Requires Windows, Access, and pywin32. The output is a committed test fixture,
so the tests (and CI) need none of the three.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time

import pythoncom
import win32com.client

AC_FORM, AC_SAVE_YES, AC_QUIT_SAVE_NONE = 2, 1, 2
# AcSection and AcCommand, from the type library.
AC_DETAIL = 0
AC_CMD_COMPILE_AND_SAVE_ALL_MODULES = 126
TEXT_BOX = 109
TEXT_BOXES = [
    'Caf\u00e9',
    'Em\u2014Dash',
    '\u20acuro',
    'Na\u00efve \u2014 x',
    '\u0418\u043c\u044f',
    'Plain',
]
# Only the controls named in letters are reached through Me: a compile error
# in Access is a dialog, and nothing here may need one dismissed.
FORM_CODE = (
    'Public Sub Touch()\r\n'
    '    Me.Caf\u00e9 = 1\r\n'
    '    Me.Plain = 2\r\n'
    'End Sub\r\n'
)


def build_form(app):
    form = app.CreateForm()
    made = form.Name
    for wanted in TEXT_BOXES:
        app.CreateControl(made, TEXT_BOX, AC_DETAIL).Name = wanted
    form.Module.InsertText(FORM_CODE)
    app.DoCmd.Save(AC_FORM, made)
    app.DoCmd.Close(AC_FORM, made, AC_SAVE_YES)
    app.DoCmd.Rename('Names', AC_FORM, made)


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    blank, out_path = sys.argv[1:3]
    running = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq MSACCESS.EXE', '/NH'], capture_output=True, text=True,
    ).stdout
    if 'MSACCESS' in running.upper():
        print('Access is already running; not touching a user instance.')
        sys.exit(3)
    # Access takes a moment to let go of the file after it quits, so a folder
    # that cannot be removed yet is left to the system rather than fought over.
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        # Access opens a database read-only when its path has forward slashes.
        scratch = os.path.abspath(os.path.join(folder, 'scratch.accdb'))
        shutil.copyfile(blank, scratch)
        app = win32com.client.DispatchEx('Access.Application')
        try:
            app.Visible = True
            app.OpenCurrentDatabase(scratch)
            build_form(app)
            app.DoCmd.RunCommand(AC_CMD_COMPILE_AND_SAVE_ALL_MODULES)
            compiled = bool(app.IsCompiled)
        finally:
            try:
                app.CloseCurrentDatabase()
            except pythoncom.com_error:
                pass
            app.Quit(AC_QUIT_SAVE_NONE)
            del app
        if not compiled:
            print('The code behind the form did not compile; no fixture written.')
            sys.exit(1)
        for attempt in range(20):
            try:
                shutil.copyfile(scratch, out_path)
                break
            except PermissionError:
                if attempt == 19:
                    raise
                time.sleep(0.5)
    print(f'{out_path}: {os.path.getsize(out_path)} bytes, form Names, compiled')


if __name__ == '__main__':
    main()
