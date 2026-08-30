#!/usr/bin/env python
"""Extracts the hidden/restricted attribute from an Office type library.

    python scripts/dump-hidden-members.py <library.exe|.dll|.olb> <out.json>

The reference dumps this repository generates from (reference/<host>/json)
carry a member's name, type, signature and documentation, but NOT whether the
library marks it hidden or restricted - so the host object model could not tell
`_CodeName` or `ActiveMenuBar` from a member a developer can actually write
(issue #56). This reads that attribute straight from the library.

Requires Windows, the Office library, and pywin32. The output is committed, so
the generators (and CI) never need any of the three.
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


def dump(lib_path):
    tlb = pythoncom.LoadTypeLib(lib_path)
    members = {}
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
    return {
        'library': lib_path,
        'hiddenTypes': sorted(hidden_types),
        'members': {name: dict(sorted(flags.items())) for name, flags in sorted(members.items())},
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
