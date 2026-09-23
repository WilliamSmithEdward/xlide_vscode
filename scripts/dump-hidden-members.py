#!/usr/bin/env python
"""Extracts the hidden/restricted attribute from an Office type library.

    python scripts/dump-hidden-members.py <library.exe|.dll|.olb> <out.json>

The reference dumps this repository generates from (reference/<host>/json)
carry a member's name, type, signature and documentation, but NOT whether the
library marks it hidden or restricted - so the host object model could not tell
`_CodeName` or `ActiveMenuBar` from a member a developer can actually write
(issue #56). This reads that attribute straight from the library.

It also records what each flagged member IS - property or method, type,
access, parameters - because the dumps leave many hidden members out entirely,
and the generator adds those to the model from here (Workbook.Title and
Worksheet.OnEntry compile and run, and were reported as missing).

Requires Windows, the Office library, and pywin32. The output lives with the
reference dumps it complements, local and uncommitted like them; the generated
model is what is committed, so CI never needs any of the three.
"""
import json
import sys

import pythoncom

# oaidl.h. THREE flags matter, not one: the Object Browser's "hidden" is
# FHIDDEN, the dispatch plumbing VBA cannot call is FRESTRICTED, and the
# default/enumerator members (_CodeName, _Default, _Evaluate, _NewEnum) carry
# only FNONBROWSABLE - which is why keying on FHIDDEN alone missed exactly the
# members issue #56 opened with.
FUNCFLAG_FRESTRICTED = 0x1
FUNCFLAG_FHIDDEN = 0x40
FUNCFLAG_FNONBROWSABLE = 0x400
VARFLAG_FREADONLY = 0x1
VARFLAG_FRESTRICTED = 0x8
VARFLAG_FHIDDEN = 0x40
VARFLAG_FNONBROWSABLE = 0x400
TYPEFLAG_FHIDDEN = 0x10
TYPEFLAG_FRESTRICTED = 0x200


def flag_names(flags, hidden_bit, restricted_bit, nonbrowsable_bit):
    names = []
    if flags & hidden_bit:
        names.append('hidden')
    if flags & restricted_bit:
        names.append('restricted')
    if flags & nonbrowsable_bit:
        names.append('nonbrowsable')
    return names


# What the library says a member IS, for the flagged members the reference
# dumps leave out entirely. A hidden member is still a member: `Workbook.Title`
# and `Worksheet.OnEntry` compile and run, and a model without them reported
# member-not-found on every closed surface that names one. The generator adds
# these, marked hidden, where the dump has no entry of its own.
INVOKE_FUNC = 1
INVOKE_PROPERTYGET = 2
PARAMFLAG_FOPT = 0x10
PARAMFLAG_FHASDEFAULT = 0x20
VT_NAMES = {
    2: 'Integer', 3: 'Long', 4: 'Single', 5: 'Double', 6: 'Currency', 7: 'Date',
    8: 'String', 9: 'Object', 10: 'Long', 11: 'Boolean', 12: 'Variant', 13: 'Object',
    16: 'Byte', 17: 'Byte', 18: 'Integer', 19: 'Long', 20: 'LongLong', 21: 'LongLong',
    22: 'Long', 23: 'Long', 24: 'void', 25: 'void', 27: 'Variant',
}
VT_PTR = 26
VT_USERDEFINED = 29


def vba_type_name(info, typedesc):
    """A VBA-spelled type name for a TYPEDESC, as the reference dumps spell one."""
    if isinstance(typedesc, int):
        return VT_NAMES.get(typedesc, 'Variant')
    vt, inner = typedesc
    if vt == VT_PTR:
        return vba_type_name(info, inner)
    if vt == VT_USERDEFINED:
        try:
            name = info.GetRefTypeInfo(inner).GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return 'Variant'
        # The library names the interface behind a coclass: `_Workbook`.
        return name[1:] if name.startswith('_') else name
    return VT_NAMES.get(vt, 'Variant')


def member_details(info, attr):
    """Kind, type, access and parameters for each member, keyed by name."""
    out = {}
    puts = set()
    for i in range(attr.cFuncs):
        try:
            desc = info.GetFuncDesc(i)
        except pythoncom.com_error:
            continue
        if desc.invkind not in (INVOKE_FUNC, INVOKE_PROPERTYGET):
            puts.add(desc.memid)
    for i in range(attr.cFuncs):
        try:
            desc = info.GetFuncDesc(i)
            names = info.GetNames(desc.memid)
        except pythoncom.com_error:
            continue
        name = names[0]
        if desc.invkind not in (INVOKE_FUNC, INVOKE_PROPERTYGET):
            # A write-only property: keep it, writable, if nothing else did.
            out.setdefault(name, {'kind': 'property', 'type': 'Variant', 'access': 'read/write'})
            continue
        parameters = []
        for index, arg in enumerate(desc.args):
            flags = arg[1]
            parameters.append({
                'name': names[index + 1] if index + 1 < len(names) else f'Arg{index + 1}',
                'type': vba_type_name(info, arg[0]),
                'optional': bool(flags & (PARAMFLAG_FOPT | PARAMFLAG_FHASDEFAULT)),
            })
        returns = vba_type_name(info, desc.rettype[0])
        if desc.invkind == INVOKE_PROPERTYGET:
            out[name] = {
                'kind': 'property',
                'type': returns,
                'access': 'read/write' if desc.memid in puts else 'read-only',
                **({'parameters': parameters} if parameters else {}),
            }
        else:
            out[name] = {'kind': 'method', 'returns': returns, 'parameters': parameters}
    for i in range(attr.cVars):
        try:
            desc = info.GetVarDesc(i)
            name = info.GetNames(desc.memid)[0]
        except pythoncom.com_error:
            continue
        out[name] = {
            'kind': 'property',
            'type': vba_type_name(info, desc.elemdescVar[0]),
            'access': 'read-only' if desc.wVarFlags & VARFLAG_FREADONLY else 'read/write',
        }
    return out


def dump(lib_path):
    tlb = pythoncom.LoadTypeLib(lib_path)
    members = {}
    details = {}
    hidden_types = []
    for index in range(tlb.GetTypeInfoCount()):
        try:
            type_name = tlb.GetDocumentation(index)[0]
            info = tlb.GetTypeInfo(index)
            attr = info.GetTypeAttr()
        except pythoncom.com_error:
            continue
        if attr.wTypeFlags & (TYPEFLAG_FHIDDEN | TYPEFLAG_FRESTRICTED):
            hidden_types.append(type_name)
        flagged = {}
        for i in range(attr.cFuncs):
            try:
                desc = info.GetFuncDesc(i)
                name = info.GetNames(desc.memid)[0]
            except pythoncom.com_error:
                continue
            names = flag_names(
                desc.wFuncFlags, FUNCFLAG_FHIDDEN, FUNCFLAG_FRESTRICTED, FUNCFLAG_FNONBROWSABLE)
            if names:
                flagged.setdefault(name, sorted(set(names + flagged.get(name, []))))
        for i in range(attr.cVars):
            try:
                desc = info.GetVarDesc(i)
                name = info.GetNames(desc.memid)[0]
            except pythoncom.com_error:
                continue
            names = flag_names(
                desc.wVarFlags, VARFLAG_FHIDDEN, VARFLAG_FRESTRICTED, VARFLAG_FNONBROWSABLE)
            if names:
                flagged.setdefault(name, sorted(set(names + flagged.get(name, []))))
        if flagged:
            members.setdefault(type_name, {}).update(flagged)
            described = member_details(info, attr)
            details.setdefault(type_name, {}).update(
                {name: described[name] for name in flagged if name in described})
    return {
        'library': lib_path,
        'hiddenTypes': sorted(hidden_types),
        'members': {name: dict(sorted(flags.items())) for name, flags in sorted(members.items())},
        'details': {name: dict(sorted(found.items())) for name, found in sorted(details.items()) if found},
    }


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    result = dump(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8', newline='\n') as handle:
        json.dump(result, handle, indent=1, sort_keys=False)
        handle.write('\n')
    print(f"{sys.argv[2]}: {len(result['members'])} types carry flagged members, "
          f"{len(result['hiddenTypes'])} hidden types")
