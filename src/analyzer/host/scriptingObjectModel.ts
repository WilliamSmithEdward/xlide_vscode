// Portable snapshots of two referenced Windows libraries. Metadata is read
// from COM by scripts/dump-scripting-typelibs.py, then generated mechanically.
// These models are selected by project-reference GUID, never added implicitly.
import type { HostObjectModel } from './excelObjectModel';
import { scriptingReferenceData } from './scriptingObjectModelData';
import { regexpReferenceData } from './regexpObjectModelData';

let scripting: HostObjectModel | undefined;
let regexp: HostObjectModel | undefined;

export function getScriptingObjectModel(): HostObjectModel {
    return scripting ??= {
        ...scriptingReferenceData(),
        source: 'Microsoft Scripting Runtime 1.0 via pyVBAReference and registered COM type library',
        hostName: 'Scripting',
        globals: {},
    };
}

export function getRegExpObjectModel(): HostObjectModel {
    return regexp ??= {
        ...regexpReferenceData(),
        source: 'Microsoft VBScript Regular Expressions 5.5 via pyVBAReference and registered COM type library',
        hostName: 'VBScript_RegExp_55',
        globals: {},
    };
}
