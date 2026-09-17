#!/usr/bin/env python
"""Builds tests/fixtures/binaries/AccessControlNamesFixture.accdb.

    python scripts/build-access-control-names-fixture.py <blank.accdb> <out.accdb>

A database whose designs name things the way real ones do: not as VBA
identifiers. The form wizard names a control after its field, so `Order Date`
is ordinary, and VBA reaches it as `Me.Order_Date`. Access keeps both names in
the design's TypeInfo stream, the identifier first:

    ident  ordinal  "Order_Date" NUL "Order Date" NUL

and leaves the second empty where the two are equal ("Plain" NUL NUL). An
ActiveX control's 36-byte tail follows both names.

What it holds:

  form Names     text boxes `Order Date`, `Qty-1`, `2ndBox`, `Tax (VAT)` and
                 `Plain`; an ActiveX control `Web View`; the form header
                 renamed `Top Part`, which Access moves to the end of the
                 stream with its ordinal kept; and code that reaches every
                 control through Me, so "it still compiles" is a real check.
  report Totals  a text box `Line Total`.

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

AC_FORM, AC_REPORT, AC_SAVE_YES, AC_QUIT_SAVE_NONE = 2, 3, 1, 2
# AcSection and AcCommand, from the type library.
AC_DETAIL, AC_HEADER = 0, 1
AC_CMD_FORM_HDR_FTR, AC_CMD_COMPILE_AND_SAVE_ALL_MODULES = 36, 126
TEXT_BOX, CUSTOM_CONTROL = 109, 119
TEXT_BOXES = ['Order Date', 'Qty-1', '2ndBox', 'Tax (VAT)', 'Plain']
FORM_CODE = (
    'Public Sub Touch()\r\n'
    '    Me.Order_Date = 1\r\n'
    '    Me.Qty_1 = 2\r\n'
    '    Me.Ctl2ndBox = 3\r\n'
    '    Me.Tax__VAT_ = 4\r\n'
    '    Me.Plain = 5\r\n'
    '    Me.Web_View.Visible = True\r\n'
    '    Me.Top_Part.Visible = True\r\n'
    'End Sub\r\n'
)


def section_of(design, index):
    """The design's section at an AcSection index, or None when it has none."""
    # `Section` is an indexed property; late-bound pywin32 calls it as a
    # method and gets "Member not found", so ask for the property outright.
    section_id = design._oleobj_.GetIDsOfNames('Section')
    try:
        return win32com.client.Dispatch(
            design._oleobj_.Invoke(section_id, 0, pythoncom.DISPATCH_PROPERTYGET, True, index))
    except pythoncom.com_error:
        return None


def build_form(app):
    form = app.CreateForm()
    made = form.Name
    for wanted in TEXT_BOXES:
        app.CreateControl(made, TEXT_BOX, AC_DETAIL).Name = wanted
    app.CreateControl(made, CUSTOM_CONTROL, AC_DETAIL).Name = 'Web View'
    # The command is a toggle, so it is sent only where the header is absent.
    if section_of(form, AC_HEADER) is None:
        app.DoCmd.RunCommand(AC_CMD_FORM_HDR_FTR)
    section_of(form, AC_HEADER).Name = 'Top Part'
    module = form.Module
    module.CreateEventProc('Click', 'Order_Date')
    module.InsertText(FORM_CODE)
    app.DoCmd.Save(AC_FORM, made)
    app.DoCmd.Close(AC_FORM, made, AC_SAVE_YES)
    app.DoCmd.Rename('Names', AC_FORM, made)


def build_report(app):
    report = app.CreateReport()
    made = report.Name
    app.CreateReportControl(made, TEXT_BOX, AC_DETAIL).Name = 'Line Total'
    app.DoCmd.Save(AC_REPORT, made)
    app.DoCmd.Close(AC_REPORT, made, AC_SAVE_YES)
    app.DoCmd.Rename('Totals', AC_REPORT, made)


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
            build_report(app)
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
    print(f'{out_path}: {os.path.getsize(out_path)} bytes, form Names and report Totals, compiled')


if __name__ == '__main__':
    main()
