// Which host types VBA resolves a member against while compiling.
//
// A COM interface marked NONEXTENSIBLE can gain no members at run time, so a
// name that is not on it can never resolve and VBA rejects it as a compile
// error. Without that flag the object is extensible: VBA compiles the call
// and asks IDispatch for the name when it runs. Excel leans on that heavily -
// `Application.Match` is on no interface in the library at all, it is a
// worksheet function Excel resolves dynamically, and it is ordinary VBA.
//
// So a member list being complete is not enough to report an absent member:
// absence is only provable where the interface is closed. Both facts are
// needed, and they come from different places - the member list from the
// reference dump, this flag from the type library's own TYPEFLAGS.
//
// Measured from the registered libraries with LoadRegTypeLib (Office 16),
// reading wTypeFlags & TYPEFLAG_FNONEXTENSIBLE (0x80) for every interface and
// dispinterface, and confirmed against the VBE for seven receivers, which
// agreed with the flag in every case:
//
//   Application, Range, Dim r As Range, Dim wb As Workbook, Font   accepted
//   Dim ws As Worksheet, Dim c As Chart, Worksheets (a Sheets)     rejected
//
// Of Excel's 747 interfaces only 27 are closed, and these are the ones the
// model carries a type for. The rest of the object model is open, which is
// why `Range("A1").Whatever` compiles and only fails when it runs.

import {
	ACCESS_CLOSED_TYPE_NAMES,
	POWERPOINT_CLOSED_TYPE_NAMES,
	WORD_CLOSED_TYPE_NAMES,
} from './closedTypeNames';

/**
 * Excel types whose interface is NONEXTENSIBLE, so the VBE refuses a member
 * that is not on it. `Worksheet` and `Chart` are the two a user meets: a typo
 * on a worksheet variable is a compile error, the same typo on a Range is not.
 */
const EXCEL_CLOSED_TYPES: ReadonlySet<string> = new Set([
	'Adjustments',
	'CalloutFormat',
	'Chart',
	'ColorFormat',
	'CubeField',
	'CubeFields',
	'DefaultWebOptions',
	'DiagramNode',
	'DiagramNodeChildren',
	'DiagramNodes',
	'FillFormat',
	'Global',
	'LineFormat',
	'Model3DFormat',
	'PictureFormat',
	'PublishObject',
	'ShadowFormat',
	'ShapeNode',
	'ShapeNodes',
	'Sheets',
	'TextEffectFormat',
	'TextFrame2',
	'ThreeDFormat',
	'TreeviewControl',
	'WebOptions',
	'Workbooks',
	'Worksheet',
]);

/**
 * Model types that stand where the library returns a closed type. The
 * Worksheets property of Application, Global and Workbook returns `Sheets`
 * in the type library, and Sheets is closed, so the VBE refuses
 * `Worksheets.Whatever`. The model returns `Worksheets` there instead, which
 * keeps `Worksheets(1)` a Worksheet where Sheets would give a Worksheet or a
 * Chart. The library's Worksheets interface is open but lists the same 29
 * members as Sheets (both read from the registered library), so the model's
 * Worksheets is closed exactly as far as Sheets is. A variable declared
 * `As Worksheets` gets the same answer; the VBE has not been asked about that
 * form, and the collection it holds lacks the name either way.
 */
const EXCEL_TYPES_STANDING_FOR: ReadonlyMap<string, string> = new Map([
	['Worksheets', 'Sheets'],
]);

/**
 * Whether VBA resolves a member against this Excel type while compiling, so
 * a name the model does not carry is genuinely absent rather than deferred.
 *
 * Takes the model's own key, qualified (`Excel.Range`) or bare, since the
 * default host is Excel. A type the set does not name is extensible, which is
 * the safe answer for one nobody has measured: nothing is reported for it.
 */
export function hostTypeResolvesWhenCompiling(qualifiedName: string): boolean {
	const dot = qualifiedName.indexOf('.');
	const library = dot > 0 ? qualifiedName.slice(0, dot).toLowerCase() : 'excel';
	const displayName = dot > 0 ? qualifiedName.slice(dot + 1) : qualifiedName;
	switch (library) {
		case 'excel':
			return EXCEL_CLOSED_TYPES.has(EXCEL_TYPES_STANDING_FOR.get(displayName) ?? displayName);
		case 'word':
			return WORD_CLOSED_TYPES.has(displayName.toLowerCase());
		case 'powerpoint':
			return POWERPOINT_CLOSED_TYPES.has(displayName.toLowerCase());
		case 'access':
			return ACCESS_CLOSED_TYPES.has(displayName.toLowerCase());
		case 'office':
			// The shared Office library (CommandBar and friends): its model
			// carries no hidden members, so absence is never proved there
			// whatever the flag says.
			return false;
		default:
			// A library nobody has measured keeps the answer it had: closed,
			// which its (never exhaustive) model cannot act on anyway.
			return true;
	}
}

/**
 * The other hosts, read the same way from their registered libraries (Word
 * 8.7, PowerPoint 2.12, Access 9.0) by the issue #127 sweep on 2026-09-26,
 * and held there by tests/hostTypeExtensibility.test.ts. Word Range,
 * Selection, Paragraph and Table are closed and Document is open; PowerPoint
 * Slide, Shape and TextRange are closed; Access's controls (TextBox,
 * ComboBox, ...) are closed while Form, Report, Control and the project
 * objects are open, which is what lets `f.CustomerID` compile.
 */
const WORD_CLOSED_TYPES: ReadonlySet<string> = new Set(WORD_CLOSED_TYPE_NAMES.map((name) => name.toLowerCase()));
const POWERPOINT_CLOSED_TYPES: ReadonlySet<string> = new Set(POWERPOINT_CLOSED_TYPE_NAMES.map((name) => name.toLowerCase()));
const ACCESS_CLOSED_TYPES: ReadonlySet<string> = new Set(ACCESS_CLOSED_TYPE_NAMES.map((name) => name.toLowerCase()));

/** The closed type names, for the test that holds them to the type library. */
export const EXCEL_CLOSED_TYPE_NAMES: readonly string[] = [...EXCEL_CLOSED_TYPES];
