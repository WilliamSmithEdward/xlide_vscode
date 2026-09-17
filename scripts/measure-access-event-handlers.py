#!/usr/bin/env python
"""Measures the event procedures Access itself writes.

    python scripts/measure-access-event-handlers.py <blank.accdb> <MSACC.OLB> <out.json>

Builds a scratch form and a scratch report holding one of every control type
Access will create there, switches on every section a design can have, then
asks Access, through Module.CreateEventProc, to write a handler for every
event name the type library knows - for the design, each section and each
control. What Access accepts is the object's real event set, and the line it
writes is the handler's real declaration, `ByVal` and all.

The output (tests/fixtures/access/eventHandlerOracle.json) is what the
generated Access model is held to: tests/accessDesignMeType renders every
handler from the model and compares it, character for character, with the
line Access wrote. `skipped` records what Access refused, which is evidence
too - it will not put a navigation control or a web browser on a report.

Safety: works on a scratch copy of a blank database, in a temporary folder,
in an Access instance it starts itself. It refuses to run while Access is
already running, saves nothing, and quits through COM.

Requires Windows, Access, and pywin32. The output is a committed test fixture,
so the tests (and CI) need none of the three.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

import pythoncom
import win32com.client

AC_FORM, AC_REPORT, AC_SAVE_NO, AC_QUIT_SAVE_NONE = 2, 3, 2, 2
# AcSection, from the type library: 0 to 8 are every section a design can have.
AC_DETAIL, AC_HEADER, AC_PAGE_HEADER = 0, 1, 3
AC_SECTIONS = range(0, 9)
IMPLTYPEFLAG_FSOURCE = 0x2
# AcCommand, from the type library. A design's own header and footer and its
# page header and footer are menu commands; a report group's are a call.
AC_CMD_FORM_HDR_FTR, AC_CMD_REPORT_HDR_FTR, AC_CMD_PAGE_HDR_FTR = 36, 37, 182
CONTROL_TYPES = {
    100: 'Label', 101: 'Rectangle', 102: 'Line', 103: 'Image', 104: 'CommandButton',
    105: 'OptionButton', 106: 'CheckBox', 107: 'OptionGroup', 108: 'BoundObjectFrame',
    109: 'TextBox', 110: 'ListBox', 111: 'ComboBox', 112: 'Subform', 114: 'ObjectFrame',
    118: 'PageBreak', 122: 'ToggleButton', 123: 'Tab', 126: 'Attachment',
    128: 'WebBrowser', 129: 'NavigationControl', 133: 'Chart', 134: 'EdgeBrowser',
}
# The types an option group can hold. The library gives each a class of its
# own there, so they are measured apart from the same control standing alone.
IN_OPTION_TYPES = {105: 'OptionButton', 106: 'CheckBox', 122: 'ToggleButton'}
OPTION_GROUP, TAB = 107, 123


def failure(err):
    return str(err.excepinfo[2] if err.excepinfo else err)[:160]


def event_names(olb):
    """Every event name any source interface of the library declares."""
    tlb = pythoncom.LoadTypeLib(olb)
    names = set()
    for index in range(tlb.GetTypeInfoCount()):
        if tlb.GetTypeInfoType(index) != pythoncom.TKIND_COCLASS:
            continue
        info = tlb.GetTypeInfo(index)
        for impl in range(info.GetTypeAttr().cImplTypes):
            if not info.GetImplTypeFlags(impl) & IMPLTYPEFLAG_FSOURCE:
                continue
            source = info.GetRefTypeInfo(info.GetRefTypeOfImplType(impl))
            for func in range(source.GetTypeAttr().cFuncs):
                names.add(source.GetNames(source.GetFuncDesc(func).memid)[0])
    return sorted(names)


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


def measure(module, object_name, events):
    """{event: declaration line} for every event Access will write a handler for."""
    out = {}
    for event in events:
        try:
            line = module.CreateEventProc(event, object_name)
        except pythoncom.com_error:
            continue
        out[event] = module.Lines(line, 1).strip()
    return out


def build(app, kind, events):
    """One design of `kind`, with everything measurable on it."""
    result = {'controls': {}, 'sections': {}, 'skipped': {}}
    design = app.CreateForm() if kind == 'form' else app.CreateReport()
    name = design.Name
    create = app.CreateControl if kind == 'form' else app.CreateReportControl
    made = []
    for code, type_name in CONTROL_TYPES.items():
        try:
            control = create(name, code, AC_DETAIL)
            made.append((control.Name, type_name, False))
            if code == OPTION_GROUP:
                for inner_code, inner_name in IN_OPTION_TYPES.items():
                    inner = create(name, inner_code, AC_DETAIL, control.Name)
                    made.append((inner.Name, inner_name, True))
            if code == TAB:
                for page in control.Pages:
                    made.append((page.Name, 'Page', False))
                    break
        except pythoncom.com_error as err:
            result['skipped'][type_name] = failure(err)
    # Both commands are toggles, and which sections a new design starts with
    # differs (a report has its page header and footer, a form has neither),
    # so each is switched on only where the design does not have it yet.
    own = AC_CMD_FORM_HDR_FTR if kind == 'form' else AC_CMD_REPORT_HDR_FTR
    for label, command, index in (
        ('own header and footer', own, AC_HEADER),
        ('page header and footer', AC_CMD_PAGE_HDR_FTR, AC_PAGE_HEADER),
    ):
        if section_of(design, index) is not None:
            continue
        try:
            app.DoCmd.RunCommand(command)
        except pythoncom.com_error as err:
            result['skipped'][label] = failure(err)
    if kind == 'report':
        try:
            app.CreateGroupLevel(name, '=1', True, True)
        except pythoncom.com_error as err:
            result['skipped']['group header and footer'] = failure(err)
    module = design.Module
    result['self'] = measure(module, 'Form' if kind == 'form' else 'Report', events)
    for index in AC_SECTIONS:
        section = section_of(design, index)
        if section is not None:
            result['sections'][f'{index}:{section.Name}'] = measure(module, section.Name, events)
    for control_name, type_name, in_option in made:
        key = f"{type_name}{' in OptionGroup' if in_option else ''}"
        result['controls'][key] = measure(module, control_name, events)
    app.DoCmd.Close(AC_FORM if kind == 'form' else AC_REPORT, name, AC_SAVE_NO)
    return result


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(2)
    blank, olb, out_path = sys.argv[1:4]
    running = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq MSACCESS.EXE', '/NH'], capture_output=True, text=True,
    ).stdout
    if 'MSACCESS' in running.upper():
        print('Access is already running; not touching a user instance.')
        sys.exit(3)
    events = event_names(olb)
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
            result = {
                'access': str(app.Version),
                'eventNamesTried': len(events),
                'form': build(app, 'form', events),
                'report': build(app, 'report', events),
            }
        finally:
            try:
                app.CloseCurrentDatabase()
            except pythoncom.com_error:
                pass
            app.Quit(AC_QUIT_SAVE_NONE)
            del app
    with open(out_path, 'w', encoding='utf-8', newline='\n') as handle:
        json.dump(result, handle, indent=1, sort_keys=True)
        handle.write('\n')
    print(f"{out_path}: form {len(result['form']['self'])} events, {len(result['form']['sections'])} sections, "
          f"{len(result['form']['controls'])} control kinds; report {len(result['report']['self'])} events, "
          f"{len(result['report']['sections'])} sections, {len(result['report']['controls'])} control kinds")


if __name__ == '__main__':
    main()
