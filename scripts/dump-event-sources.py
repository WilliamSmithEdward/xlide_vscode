#!/usr/bin/env python
"""Extracts the events of every class in an Office type library.

    python scripts/dump-event-sources.py <library.exe|.dll|.olb> <out.json>

The reference dumps this repository generates from (reference/<host>/json)
name an event and its parameters, but not how a parameter is PASSED, and VBA
refuses a handler whose ByVal does not match the event's: Access writes
`Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)` and
`Form_Unload(Cancel As Integer)`, and either one declared the other way is a
compile error. The dumps also lost three classes outright - TextBox, CheckBox
and ComboBox share a file name, case aside, with their own interfaces. So the
events are read straight from the library:

  classes      each class's DEFAULT source interface, which is the one the
               VBE binds handlers against (a Form's is _FormEvents2, not the
               older _FormEvents it also implements);
  interfaces   each of those interfaces' events in library order, hidden ones
               left out, every parameter with its type and whether it is
               passed by value.

Measured against Access 16.0 itself: Module.CreateEventProc wrote 722 handlers
for a form, a report, every kind of section and 26 kinds of control
(scripts/measure-access-event-handlers.py, kept as
tests/fixtures/access/eventHandlerOracle.json), and every one matched what
this renders, character for character. The three events the library hides on
a form (RecordExit, BeginBatchEdit, UndoBatchEdit) are the three Access
refused to write.

Requires Windows, the Office library, and pywin32. The output belongs with the
rest of the reference corpus (reference/<host>/eventSources.json), which is a
generator input and is not committed; the model generated from it is, and the
handler oracle is what holds that model to Access on every machine.
"""
import json
import os
import sys

import pythoncom

IMPLTYPEFLAG_FDEFAULT = 0x1
IMPLTYPEFLAG_FSOURCE = 0x2
# oaidl.h. An event the library hides or restricts is one the VBE never lists.
FUNCFLAG_FRESTRICTED = 0x1
FUNCFLAG_FHIDDEN = 0x40

VBA_TYPES = {
    pythoncom.VT_UI1: 'Byte',
    pythoncom.VT_I2: 'Integer',
    pythoncom.VT_I4: 'Long',
    pythoncom.VT_INT: 'Long',
    pythoncom.VT_R4: 'Single',
    pythoncom.VT_R8: 'Double',
    pythoncom.VT_CY: 'Currency',
    pythoncom.VT_DATE: 'Date',
    pythoncom.VT_BSTR: 'String',
    pythoncom.VT_BOOL: 'Boolean',
    pythoncom.VT_VARIANT: 'Variant',
    pythoncom.VT_DISPATCH: 'Object',
    pythoncom.VT_UNKNOWN: 'IUnknown',
}


OBJECT_KINDS = (pythoncom.TKIND_INTERFACE, pythoncom.TKIND_DISPATCH, pythoncom.TKIND_COCLASS)


def is_object_type(info, tdesc):
    """Whether a type description names a class or an interface."""
    if not (isinstance(tdesc, tuple) and tdesc[0] == pythoncom.VT_USERDEFINED):
        return False
    return info.GetRefTypeInfo(tdesc[1]).GetTypeAttr().typekind in OBJECT_KINDS


def vba_type(info, tdesc):
    """(VBA type name, passed by reference) for one parameter's type."""
    if isinstance(tdesc, tuple):
        vt, inner = tdesc
        if vt == pythoncom.VT_PTR:
            # A pointer to a value is ByRef. A pointer to an object type IS the
            # object, passed by value; only a pointer to that pointer is ByRef.
            name, inner_by_ref = vba_type(info, inner)
            return name, inner_by_ref or not is_object_type(info, inner)
        if vt == pythoncom.VT_USERDEFINED:
            return info.GetRefTypeInfo(inner).GetDocumentation(-1)[0], False
        raise ValueError(f'unhandled type description {tdesc!r}')
    if tdesc not in VBA_TYPES:
        raise ValueError(f'unhandled variant type {tdesc}')
    return VBA_TYPES[tdesc], False


def events_of(info):
    out = []
    for index in range(info.GetTypeAttr().cFuncs):
        desc = info.GetFuncDesc(index)
        if desc.wFuncFlags & (FUNCFLAG_FHIDDEN | FUNCFLAG_FRESTRICTED):
            continue
        names = info.GetNames(desc.memid)
        params = []
        for position, arg in enumerate(desc.args):
            type_name, by_ref = vba_type(info, arg[0])
            param = {'name': names[position + 1], 'type': type_name}
            if not by_ref:
                param['byVal'] = True
            params.append(param)
        out.append({'name': names[0], 'params': params})
    return out


def dump(lib_path):
    tlb = pythoncom.LoadTypeLib(lib_path)
    classes = {}
    interfaces = {}
    for index in range(tlb.GetTypeInfoCount()):
        if tlb.GetTypeInfoType(index) != pythoncom.TKIND_COCLASS:
            continue
        info = tlb.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        for impl in range(attr.cImplTypes):
            flags = info.GetImplTypeFlags(impl)
            if flags & IMPLTYPEFLAG_FSOURCE and flags & IMPLTYPEFLAG_FDEFAULT:
                source = info.GetRefTypeInfo(info.GetRefTypeOfImplType(impl))
                source_name = source.GetDocumentation(-1)[0]
                classes[tlb.GetDocumentation(index)[0]] = source_name
                if source_name not in interfaces:
                    interfaces[source_name] = events_of(source)
    name, doc = tlb.GetDocumentation(-1)[:2]
    lib = tlb.GetLibAttr()
    return {
        # The library's own name and version, not where this machine keeps it.
        'library': f'{name} {lib[3]}.{lib[4]} ({os.path.basename(lib_path)}): {doc}',
        'classes': dict(sorted(classes.items())),
        'interfaces': dict(sorted(interfaces.items())),
    }


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    result = dump(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8', newline='\n') as handle:
        json.dump(result, handle, indent=1, sort_keys=False)
        handle.write('\n')
    events = sum(len(entry) for entry in result['interfaces'].values())
    print(f"{sys.argv[2]}: {len(result['classes'])} classes raise events through "
          f"{len(result['interfaces'])} interfaces, {events} events")
