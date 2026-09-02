#!/usr/bin/env python3
"""Dump the type libraries inside msvbvm60.dll into reference/vb6/json.

The VB6 runtime carries two type libraries as resources: resource 1 is `VBA`
(the VBA6 language runtime: Strings, Math, Collection, ErrObject ...) and
resource 3 is `VBRUN` (the runtime's own objects and every intrinsic
constant: vbNormal, vbKeyReturn, DataObject, AmbientProperties ...). `VB.OLB`,
which declares App, Screen, Printer and the intrinsic controls, ships only
with the VB6 IDE and is not available here; that library is transcribed from
documentation instead (scripts/transcribe-vb6-docs.mjs).

The dumps take the shape pyVBAReference writes for the Office libraries, so
scripts/generate-host-object-model.mjs reads them unchanged, plus a
`libraryId` (VBA | VBRUN) the generator maps to a namespace prefix and a
`source` naming the DLL and its version.

    python scripts/dump-vb6-typelib.py [path\\to\\msvbvm60.dll]

Windows only (pythoncom). Idempotent: re-running rewrites the same files.
"""
import io
import json
import os
import sys

import pythoncom  # type: ignore

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, 'reference', 'vb6', 'json')

DEFAULT_DLL = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'SysWOW64', 'msvbvm60.dll')
RESOURCES = [(1, 'VBA'), (3, 'VBRUN')]

VT_NAMES = {
    pythoncom.VT_EMPTY: '', pythoncom.VT_NULL: 'Null', pythoncom.VT_I2: 'Integer',
    pythoncom.VT_I4: 'Long', pythoncom.VT_R4: 'Single', pythoncom.VT_R8: 'Double',
    pythoncom.VT_CY: 'Currency', pythoncom.VT_DATE: 'Date', pythoncom.VT_BSTR: 'String',
    pythoncom.VT_DISPATCH: 'Object', pythoncom.VT_ERROR: 'Error', pythoncom.VT_BOOL: 'Boolean',
    pythoncom.VT_VARIANT: 'Variant', pythoncom.VT_UNKNOWN: 'IUnknown', pythoncom.VT_DECIMAL: 'Decimal',
    pythoncom.VT_I1: 'Byte', pythoncom.VT_UI1: 'Byte', pythoncom.VT_UI2: 'Integer',
    pythoncom.VT_UI4: 'Long', pythoncom.VT_I8: 'LongLong', pythoncom.VT_UI8: 'LongLong',
    pythoncom.VT_INT: 'Long', pythoncom.VT_UINT: 'Long', pythoncom.VT_VOID: 'void',
    pythoncom.VT_HRESULT: 'HRESULT', pythoncom.VT_LPSTR: 'String', pythoncom.VT_LPWSTR: 'String',
}

FUNCFLAG_FRESTRICTED = 0x1
FUNCFLAG_FHIDDEN = 0x40
VARFLAG_FREADONLY = 0x1
VARFLAG_FHIDDEN = 0x40
VARFLAG_FRESTRICTED = 0x80
TYPEFLAG_FRESTRICTED = 0x1
TYPEFLAG_FHIDDEN = 0x10
IMPLTYPEFLAG_FDEFAULT = 0x1
IMPLTYPEFLAG_FSOURCE = 0x2
PARAMFLAG_FOPT = 0x10

TKIND_ENUM, TKIND_RECORD, TKIND_MODULE, TKIND_INTERFACE, TKIND_DISPATCH, TKIND_COCLASS, TKIND_ALIAS = range(7)
KIND_LABEL = {
    TKIND_INTERFACE: 'Interface',
    TKIND_DISPATCH: 'Dispatch Interface',
    TKIND_COCLASS: 'Class',
    TKIND_MODULE: 'Module',
    TKIND_ENUM: 'Enumeration',
}


def resolve_type(tdesc, info):
    if tdesc is None:
        return ''
    if isinstance(tdesc, int):
        return VT_NAMES.get(tdesc, 'VT_%d' % tdesc)
    vt = tdesc[0]
    if vt == pythoncom.VT_PTR:
        return resolve_type(tdesc[1], info)
    if vt == pythoncom.VT_SAFEARRAY:
        return 'SAFEARRAY(%s)' % resolve_type(tdesc[1], info)
    if vt == pythoncom.VT_CARRAY:
        return '%s()' % resolve_type(tdesc[1], info)
    if vt == pythoncom.VT_USERDEFINED:
        try:
            return info.GetRefTypeInfo(tdesc[1]).GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return 'Unknown'
    return VT_NAMES.get(vt, 'VT_%d' % vt)


def ret_typedesc(fd):
    rt = getattr(fd, 'rettype', None)
    if rt is None:
        return None
    try:
        return rt[0]
    except (TypeError, IndexError):
        return rt


def signature_and_params(info, fd, names):
    fname = names[0]
    argnames = list(names[1:])
    args = getattr(fd, 'args', None) or ()
    parts, params = [], []
    for idx in range(len(args)):
        aname = argnames[idx] if idx < len(argnames) else 'Arg%d' % (idx + 1)
        tdesc, pflags = None, 0
        try:
            elem = args[idx]
            tdesc = elem[0]
            if len(elem) > 1 and isinstance(elem[1], int):
                pflags = elem[1]
        except (IndexError, TypeError):
            pass
        atype = resolve_type(tdesc, info)
        optional = bool(pflags & PARAMFLAG_FOPT)
        part = aname + (' As %s' % atype if atype else '')
        parts.append('[%s]' % part if optional else part)
        params.append({'name': aname, 'type': atype, 'optional': optional, 'description': ''})
    rettype = resolve_type(ret_typedesc(fd), info)
    sig = '%s(%s)' % (fname, ', '.join(parts))
    if rettype and rettype not in ('void', 'HRESULT', 'Null'):
        sig += ' As %s' % rettype
    return sig, params, rettype


def doc_of(info, memid):
    try:
        return info.GetDocumentation(memid)[1] or ''
    except pythoncom.com_error:
        return ''


def interface_members(info):
    """(properties, methods) of a dispatch or custom interface."""
    attr = info.GetTypeAttr()
    access, ptype, pdoc, order, methods = {}, {}, {}, [], []
    for i in range(attr.cFuncs):
        try:
            fd = info.GetFuncDesc(i)
        except pythoncom.com_error:
            continue
        if fd.wFuncFlags & (FUNCFLAG_FHIDDEN | FUNCFLAG_FRESTRICTED):
            continue
        names = info.GetNames(fd.memid)
        if not names:
            continue
        fname = names[0]
        doc = doc_of(info, fd.memid)
        if fd.invkind == pythoncom.INVOKE_PROPERTYGET:
            access.setdefault(fname, set()).add('get')
            if fname not in order:
                order.append(fname)
            ptype[fname] = resolve_type(ret_typedesc(fd), info)
            if doc and fname not in pdoc:
                pdoc[fname] = doc
        elif fd.invkind in (pythoncom.INVOKE_PROPERTYPUT, pythoncom.INVOKE_PROPERTYPUTREF):
            access.setdefault(fname, set()).add('set')
            if fname not in order:
                order.append(fname)
            if doc and fname not in pdoc:
                pdoc[fname] = doc
            if fname not in ptype and getattr(fd, 'args', None):
                try:
                    ptype[fname] = resolve_type(fd.args[-1][0], info)
                except (IndexError, TypeError):
                    pass
        else:
            sig, params, rettype = signature_and_params(info, fd, names)
            method = {'name': fname, 'kind': 'method', 'signature': sig, 'description': doc}
            if rettype and rettype not in ('void', 'HRESULT'):
                method['returns'] = rettype
            if params:
                method['parameters'] = params
            methods.append(method)
    for i in range(attr.cVars):
        try:
            vd = info.GetVarDesc(i)
        except pythoncom.com_error:
            continue
        if vd.wVarFlags & (VARFLAG_FHIDDEN | VARFLAG_FRESTRICTED):
            continue
        names = info.GetNames(vd.memid)
        if not names:
            continue
        vname = names[0]
        access.setdefault(vname, set()).add('get')
        if not (vd.wVarFlags & VARFLAG_FREADONLY):
            access[vname].add('set')
        if vname not in order:
            order.append(vname)
        try:
            ptype[vname] = resolve_type(vd.elemdescVar[0], info)
        except (AttributeError, IndexError, TypeError):
            pass
        d = doc_of(info, vd.memid)
        if d and vname not in pdoc:
            pdoc[vname] = d
    properties = []
    for name in order:
        acc = access.get(name, set())
        label = 'read/write' if {'get', 'set'} <= acc else ('read-only' if 'get' in acc else 'write-only')
        properties.append({'name': name, 'kind': 'property', 'type': ptype.get(name, ''),
                           'access': label, 'description': pdoc.get(name, '')})
    return properties, methods


def event_members(info):
    attr = info.GetTypeAttr()
    events = []
    for i in range(attr.cFuncs):
        try:
            fd = info.GetFuncDesc(i)
        except pythoncom.com_error:
            continue
        if fd.wFuncFlags & (FUNCFLAG_FHIDDEN | FUNCFLAG_FRESTRICTED):
            continue
        names = info.GetNames(fd.memid)
        if not names:
            continue
        sig, params, _ = signature_and_params(info, fd, names)
        event = {'name': names[0], 'kind': 'event', 'signature': sig, 'description': doc_of(info, fd.memid)}
        if params:
            event['parameters'] = params
        events.append(event)
    return events


def enum_constants(info):
    attr = info.GetTypeAttr()
    out = []
    for i in range(attr.cVars):
        try:
            vd = info.GetVarDesc(i)
        except pythoncom.com_error:
            continue
        if vd.wVarFlags & (VARFLAG_FHIDDEN | VARFLAG_FRESTRICTED):
            continue
        names = info.GetNames(vd.memid)
        if not names:
            continue
        entry = {'name': names[0], 'value': getattr(vd, 'value', None), 'description': doc_of(info, vd.memid)}
        try:
            ctype = resolve_type(vd.elemdescVar[0], info)
            if ctype and attr.typekind == TKIND_MODULE:
                entry['type'] = ctype
        except (AttributeError, IndexError, TypeError):
            pass
        out.append(entry)
    return out


def coclass_interfaces(info):
    attr = info.GetTypeAttr()
    default, source = None, None
    for j in range(attr.cImplTypes):
        try:
            flags = info.GetImplTypeFlags(j)
            ref = info.GetRefTypeInfo(info.GetRefTypeOfImplType(j))
        except pythoncom.com_error:
            continue
        if flags & IMPLTYPEFLAG_FSOURCE:
            if source is None or flags & IMPLTYPEFLAG_FDEFAULT:
                source = ref
        elif default is None or flags & IMPLTYPEFLAG_FDEFAULT:
            default = ref
    return default, source


def file_version(dll):
    try:
        import win32api  # type: ignore
        info = win32api.GetFileVersionInfo(dll, '\\')
        ms, ls = info['FileVersionMS'], info['FileVersionLS']
        return '%d.%d.%d.%d' % (ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF)
    except Exception:  # noqa: BLE001 - the version is a label, not a gate
        return 'unknown'


def dump_library(dll, resource, library_id, version):
    tlb = pythoncom.LoadTypeLib('%s\\%d' % (dll, resource))
    lib_name, lib_doc = tlb.GetDocumentation(-1)[:2]
    source = '%s resource %d (%s %s, file version %s)' % (os.path.basename(dll), resource, lib_name, lib_doc or '', version)
    written, skipped = [], []
    for i in range(tlb.GetTypeInfoCount()):
        info = tlb.GetTypeInfo(i)
        attr = info.GetTypeAttr()
        name, doc = tlb.GetDocumentation(i)[:2]
        kind = attr.typekind
        if name.startswith('_') or attr.wTypeFlags & (TYPEFLAG_FHIDDEN | TYPEFLAG_FRESTRICTED) or kind not in KIND_LABEL:
            skipped.append('%s (%s)' % (name, KIND_LABEL.get(kind, 'kind %d' % kind)))
            continue
        entry = {
            'name': name,
            'kind': KIND_LABEL[kind],
            'guid': str(attr.iid) if kind != TKIND_ENUM else '',
            'library': '%s (%s)' % (lib_name, library_id),
            'libraryId': library_id,
            'source': source,
            'description': doc or '',
            'remarks': '',
            'example': '',
        }
        if kind == TKIND_ENUM:
            entry['constants'] = enum_constants(info)
        elif kind == TKIND_MODULE:
            _props, funcs = interface_members(info)
            entry['methods'] = funcs
            entry['constants'] = enum_constants(info)
        elif kind == TKIND_COCLASS:
            default, src = coclass_interfaces(info)
            props, methods = interface_members(default) if default is not None else ([], [])
            entry['properties'], entry['methods'] = props, methods
            entry['events'] = event_members(src) if src is not None else []
        else:
            props, methods = interface_members(info)
            entry['properties'], entry['methods'], entry['events'] = props, methods, []
        with io.open(os.path.join(OUT_DIR, name + '.json'), 'w', encoding='utf-8', newline='\n') as fh:
            json.dump(entry, fh, indent=2)
            fh.write('\n')
        written.append(name)
    return lib_name, written, skipped


def main():
    dll = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DLL
    if not os.path.exists(dll):
        print('msvbvm60.dll not found at', dll)
        return 1
    os.makedirs(OUT_DIR, exist_ok=True)
    version = file_version(dll)
    index = {'dll': os.path.basename(dll), 'fileVersion': version, 'libraries': []}
    for resource, library_id in RESOURCES:
        lib_name, written, skipped = dump_library(dll, resource, library_id, version)
        index['libraries'].append({'resource': resource, 'libraryId': library_id, 'name': lib_name,
                                   'types': written, 'skipped': skipped})
        print('%s (resource %d): %d types written, %d skipped' % (lib_name, resource, len(written), len(skipped)))
    with io.open(os.path.join(OUT_DIR, '_index.json'), 'w', encoding='utf-8', newline='\n') as fh:
        json.dump(index, fh, indent=2)
        fh.write('\n')
    return 0


if __name__ == '__main__':
    sys.exit(main())
