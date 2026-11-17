// stdole (OLE Automation, stdole2.tlb), which the default references of a new
// Excel, Word, PowerPoint and Access file all hold. What it offers without a
// qualifier: the StdFunctions module's LoadPicture and SavePicture, and the
// members of its OLE_TRISTATE and LoadPictureConstants enums.
//
// Measured in Excel, Word, PowerPoint and Access 16.0 (build 20527) through
// the VBE oracle: each name compiles bare under Option Explicit. The
// signatures are the VBE's Quick Info for the two functions. In Access,
// Access's own Application.LoadPicture(FileName) wins over stdole's.
//
// A project without stdole could not be measured, since the oracle only adds
// references, so these names are known whatever the reference list says.

/** stdole's two functions, by lowercased name, with the VBE's Quick Info signature. */
export const STDOLE_FUNCTION_SIGNATURES: ReadonlyMap<string, string> = new Map([
	['loadpicture', 'LoadPicture([filename], [widthDesired As Long], [heightDesired As Long], [flags As LoadPictureConstants = Default]) As IPictureDisp'],
	['savepicture', 'SavePicture(Picture As IPictureDisp, filename As String)'],
]);

/** The members of stdole's OLE_TRISTATE and LoadPictureConstants enums, lowercased. */
export const STDOLE_CONSTANT_NAMES: ReadonlySet<string> = new Set([
	'unchecked', 'checked', 'gray',
	'default', 'monochrome', 'vgacolor', 'color',
]);

/** The hosts whose `hostName` the measurements above cover; a model with none is Excel's. */
export const STDOLE_MEASURED_HOSTS: ReadonlySet<string> = new Set(['Excel', 'Word', 'PowerPoint', 'Access']);
