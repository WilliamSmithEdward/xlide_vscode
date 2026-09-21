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
	// Only Excel has been measured against its type library. Another host's
	// types keep the answer they had, which is what they have always been
	// analyzed under - and Word and PowerPoint are closed almost throughout,
	// so the flag would change little there anyway.
	return library === 'excel' ? EXCEL_CLOSED_TYPES.has(displayName) : true;
}

/** The closed type names, for the test that holds them to the type library. */
export const EXCEL_CLOSED_TYPE_NAMES: readonly string[] = [...EXCEL_CLOSED_TYPES];
