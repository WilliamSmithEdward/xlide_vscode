#!/usr/bin/env python
"""Extracts what VBA sees of a type library's plain (vtable) interfaces.

    python scripts/dump-vtable-members.py <library.exe|.dll|.olb> <out.json>

The reference dumps this repository generates from (reference/<host>/json)
record an interface the way the library declares it. For a dispatch or dual
interface that is what VBA shows. For a plain interface (TKIND_INTERFACE) it
is not: every function returns HRESULT, the value is its last [out, retval]
parameter, and an [lcid] parameter is filled in by the caller. The dumps kept
all three as written, so the Office model typed DocumentProperties.Count as
HRESULT, and gave DocumentProperties.Add the parameters lcid and ppIDocProp,
which VBA never shows.

This reads each plain interface's functions and writes, per member, the type
VBA gives it (the retval parameter's), the parameters VBA shows, and whether a
property can be assigned. The generator applies it to the members the dump
typed HRESULT.

Requires Windows, the Office library, and pywin32. The output lives with the
reference dumps it corrects (reference/<host>/vtable.json), local and
uncommitted like them; the generated model is what is committed.
"""
import json
import sys

import pythoncom

TKIND_INTERFACE = 3
TKIND_ALIAS = 6
INVOKE_FUNC = 1
INVOKE_PROPERTYGET = 2
PARAMFLAG_FLCID = 0x4
PARAMFLAG_FRETVAL = 0x8
PARAMFLAG_FOPT = 0x10
PARAMFLAG_FHASDEFAULT = 0x20
VT_NAMES = {
    2: 'Integer', 3: 'Long', 4: 'Single', 5: 'Double', 6: 'Currency', 7: 'Date',
    8: 'String', 9: 'Object', 10: 'Long', 11: 'Boolean', 12: 'Variant', 13: 'Object',
    16: 'Byte', 17: 'Byte', 18: 'Integer', 19: 'Long', 20: 'LongLong', 21: 'LongLong',
    22: 'Long', 23: 'Long', 24: 'void', 25: 'void', 27: 'Variant',
}
VT_HRESULT = 25
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
            ref = info.GetRefTypeInfo(inner)
            attr = ref.GetTypeAttr()
            if attr.typekind == TKIND_ALIAS:
                return vba_type_name(ref, attr.tdescAlias)
            name = ref.GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return 'Variant'
        # The library names the interface behind a coclass: `_Workbook`.
        return name[1:] if name.startswith('_') else name
    return VT_NAMES.get(vt, 'Variant')


def interface_members(info, attr):
    """Each member as VBA sees it, keyed by name."""
    out = {}
    assignable = set()
    for i in range(attr.cFuncs):
        desc = info.GetFuncDesc(i)
        if desc.invkind not in (INVOKE_FUNC, INVOKE_PROPERTYGET):
            assignable.add(desc.memid)
    for i in range(attr.cFuncs):
        desc = info.GetFuncDesc(i)
        if desc.invkind not in (INVOKE_FUNC, INVOKE_PROPERTYGET):
            continue
        names = info.GetNames(desc.memid)
        # A function the library declares with its own return type (Parent
        # As IDispatch) has no retval parameter: that type is the value.
        value = 'void' if desc.rettype[0] == VT_HRESULT else vba_type_name(info, desc.rettype[0])
        parameters = []
        for index, arg in enumerate(desc.args):
            typedesc, flags = arg[0], arg[1]
            if flags & PARAMFLAG_FRETVAL:
                value = vba_type_name(info, typedesc)
                continue
            if flags & PARAMFLAG_FLCID:
                continue
            parameters.append({
                'name': names[index + 1] if index + 1 < len(names) else f'Arg{index + 1}',
                'type': vba_type_name(info, typedesc),
                'optional': bool(flags & (PARAMFLAG_FOPT | PARAMFLAG_FHASDEFAULT)),
            })
        if desc.invkind == INVOKE_PROPERTYGET:
            out[names[0]] = {
                'kind': 'property',
                'type': value,
                'access': 'read/write' if desc.memid in assignable else 'read-only',
                **({'parameters': parameters} if parameters else {}),
            }
        else:
            out[names[0]] = {'kind': 'method', 'returns': value, 'parameters': parameters}
    return out


def dump(lib_path):
    tlb = pythoncom.LoadTypeLib(lib_path)
    interfaces = {}
    for index in range(tlb.GetTypeInfoCount()):
        info = tlb.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        if attr.typekind != TKIND_INTERFACE:
            continue
        members = interface_members(info, attr)
        if members:
            interfaces[tlb.GetDocumentation(index)[0]] = dict(sorted(members.items()))
    return {'library': lib_path, 'interfaces': dict(sorted(interfaces.items()))}


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    result = dump(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8', newline='\n') as handle:
        json.dump(result, handle, indent=1)
        handle.write('\n')
    print(f"{sys.argv[2]}: {len(result['interfaces'])} plain interfaces")
