// Canonical VBA keyword table.
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025),
// section 3.3.5.2 "Reserved Identifiers and IDENTIFIER" (PDF pages 34-36).
//
// VBA identifiers are case-insensitive (MS-VBAL section 3.3.5.2). This table
// maps a lowercased keyword to the canonical capitalization that the VBA editor
// (VBE) emits. Canonical casing below is taken verbatim from the spec grammar,
// with one documented exception: the spec writes literal-identifiers in lower
// case ("true"/"false"/"nothing"/"empty"/"null", section 3.3.5.2) but the VBE
// always renders them capitalized (True/False/Nothing/Empty/Null). We follow
// the VBE convention because this table drives editor presentation.
//
// Two kinds of keyword are tracked:
//   1. Reserved identifiers - the closed set defined by MS-VBAL section 3.3.5.2.
//      An identifier in this set can never be a user-defined name (it is not an
//      <IDENTIFIER>). Used by the lexer/parser to distinguish reserved tokens.
//   2. Contextual keywords - words that are NOT reserved identifiers per the
//      spec but that the VBE still capitalizes because they are meaningful in a
//      specific statement (e.g. "Explicit" in Option Explicit, "Lib"/"Alias" in
//      Declare, "Step" in For). These are VBE-convention casing, not section
//      3.3.5.2 reserved words, and are marked as such below.

// ---------------------------------------------------------------------------
// 1. Reserved identifiers (MS-VBAL section 3.3.5.2) - canonical casing verbatim.
// ---------------------------------------------------------------------------

/** statement-keyword (MS-VBAL section 3.3.5.2). */
export const STATEMENT_KEYWORDS: readonly string[] = [
	'Call', 'Case', 'Close', 'Const', 'Declare', 'DefBool', 'DefByte', 'DefCur',
	'DefDate', 'DefDbl', 'DefInt', 'DefLng', 'DefLngLng', 'DefLngPtr', 'DefObj',
	'DefSng', 'DefStr', 'DefVar', 'Dim', 'Do', 'Else', 'ElseIf', 'End', 'EndIf',
	'Enum', 'Erase', 'Event', 'Exit', 'For', 'Friend', 'Function', 'Get',
	'Global', 'GoSub', 'GoTo', 'If', 'Implements', 'Input', 'Let', 'Lock',
	'Loop', 'LSet', 'Next', 'On', 'Open', 'Option', 'Print', 'Private', 'Public',
	'Put', 'RaiseEvent', 'ReDim', 'Resume', 'Return', 'RSet', 'Seek', 'Select',
	'Set', 'Static', 'Stop', 'Sub', 'Type', 'Unlock', 'Wend', 'While', 'With',
	'Write',
];

/** rem-keyword (MS-VBAL section 3.3.5.2). */
export const REM_KEYWORD = 'Rem';

/** marker-keyword (MS-VBAL section 3.3.5.2). */
export const MARKER_KEYWORDS: readonly string[] = [
	'Any', 'As', 'ByRef', 'ByVal', 'Case', 'Each', 'Else', 'In', 'New', 'Shared',
	'Until', 'WithEvents', 'Write', 'Optional', 'ParamArray', 'Preserve', 'Spc',
	'Tab', 'Then', 'To',
];

/** operator-identifier (MS-VBAL section 3.3.5.2). */
export const OPERATOR_IDENTIFIERS: readonly string[] = [
	'AddressOf', 'And', 'Eqv', 'Imp', 'Is', 'Like', 'New', 'Mod', 'Not', 'Or',
	'TypeOf', 'Xor',
];

/** reserved-name (MS-VBAL section 3.3.5.2). */
export const RESERVED_NAMES: readonly string[] = [
	'Abs', 'CBool', 'CByte', 'CCur', 'CDate', 'CDbl', 'CDec', 'CInt', 'CLng',
	'CLngLng', 'CLngPtr', 'CSng', 'CStr', 'CVar', 'CVErr', 'Date', 'Debug',
	'DoEvents', 'Fix', 'Int', 'Len', 'LenB', 'Me', 'PSet', 'Scale', 'Sgn',
	'String',
];

/** special-form (MS-VBAL section 3.3.5.2). */
export const SPECIAL_FORMS: readonly string[] = [
	'Array', 'Circle', 'Input', 'InputB', 'LBound', 'Scale', 'UBound',
];

/** reserved-type-identifier (MS-VBAL section 3.3.5.2). */
export const RESERVED_TYPE_IDENTIFIERS: readonly string[] = [
	'Boolean', 'Byte', 'Currency', 'Date', 'Double', 'Integer', 'Long',
	'LongLong', 'LongPtr', 'Single', 'String', 'Variant',
];

// literal-identifier (MS-VBAL section 3.3.5.2). Spec grammar writes these lower
// case; the VBE renders them capitalized, which is the form used here.
export const LITERAL_IDENTIFIERS: readonly string[] = [
	'True', 'False', 'Nothing', 'Empty', 'Null',
];

/** future-reserved (MS-VBAL section 3.3.5.2). */
export const FUTURE_RESERVED: readonly string[] = [
	'CDecl', 'Decimal', 'DefDec',
];

// reserved-for-implementation-use (MS-VBAL section 3.3.5.2). These are reserved
// identifiers for declaration validation. They are intentionally not keyword-
// cased by the lexer because exported source metadata uses them as raw
// Attribute-line spellings such as `Attribute VB_Name = "Module1"`.
export const RESERVED_FOR_IMPLEMENTATION_USE: readonly string[] = [
	'Attribute', 'LINEINPUT', 'VB_Base', 'VB_Control', 'VB_Creatable',
	'VB_Customizable', 'VB_Description', 'VB_Exposed', 'VB_Ext_KEY',
	'VB_GlobalNameSpace', 'VB_HelpID', 'VB_Invoke_Func', 'VB_Invoke_Property',
	'VB_Invoke_PropertyPut', 'VB_Invoke_PropertyPutRef', 'VB_MemberFlags',
	'VB_Name', 'VB_PredeclaredId', 'VB_ProcData', 'VB_TemplateDerived',
	'VB_UserMemId', 'VB_VarDescription', 'VB_VarHelpID', 'VB_VarMemberFlags',
	'VB_VarProcData', 'VB_VarUserMemId',
];

// ---------------------------------------------------------------------------
// 2. Contextual keywords - NOT reserved identifiers per MS-VBAL section
//    3.3.5.2, but capitalized by the VBE in their statement context. Casing
//    follows VBE convention (source-of-truth hierarchy level 2), not the
//    closed reserved-identifier grammar. Grouped by the statement they appear
//    in for traceability.
// ---------------------------------------------------------------------------
export const CONTEXTUAL_KEYWORDS: readonly string[] = [
	// Option statements: Option Explicit / Base / Compare Binary / Compare Text.
	'Explicit', 'Base', 'Compare', 'Binary', 'Text',
	// Declare statement: Declare [PtrSafe] ... Lib "..." Alias "...". PtrSafe is
	// the 64-bit form every modern Declare carries, and the keywords reference
	// lists it (issue #41).
	'Lib', 'Alias', 'PtrSafe',
	// Property statement keyword (note: not in section 3.3.5.2; Get/Let/Set are).
	'Property',
	// For ... Step.
	'Step',
	// On Error / Error statement / Error type-coercion form.
	'Error',
	// Open ... For <mode> As: Output / Append / Random / Read (Input/Lock/Write
	// are reserved and live above).
	'Output', 'Append', 'Random', 'Read',
	// Object: used as a type name; MS-VBAL section 3.3.5.3 notes "object" is not
	// a reserved-identifier but is used as if it were a reserved-type-identifier.
	'Object',
];

// ---------------------------------------------------------------------------
// 3. Derived lookup structures.
// ---------------------------------------------------------------------------

const KEYWORD_CASING_LISTS: readonly (readonly string[])[] = [
	STATEMENT_KEYWORDS,
	[REM_KEYWORD],
	MARKER_KEYWORDS,
	OPERATOR_IDENTIFIERS,
	RESERVED_NAMES,
	SPECIAL_FORMS,
	RESERVED_TYPE_IDENTIFIERS,
	LITERAL_IDENTIFIERS,
	FUTURE_RESERVED,
];

const RESERVED_IDENTIFIER_LISTS: readonly (readonly string[])[] = [
	...KEYWORD_CASING_LISTS,
	RESERVED_FOR_IMPLEMENTATION_USE,
];

/**
 * Lowercased names of every reserved identifier defined by MS-VBAL section
 * 3.3.5.2. A name in this set is never an <IDENTIFIER>.
 */
export const RESERVED_IDENTIFIERS: ReadonlySet<string> = new Set(
	RESERVED_IDENTIFIER_LISTS.flat().map((w) => w.toLowerCase()),
);

function buildKeywordMap(): Record<string, string> {
	const map: Record<string, string> = {};
	const add = (word: string): void => {
		const key = word.toLowerCase();
		// Reserved-identifier casing wins over contextual casing on conflict, but
		// there are no conflicting spellings between the two sets today.
		if (!(key in map)) {
			map[key] = word;
		}
	};
	for (const list of KEYWORD_CASING_LISTS) {
		for (const word of list) {
			add(word);
		}
	}
	for (const word of CONTEXTUAL_KEYWORDS) {
		add(word);
	}
	return map;
}

/**
 * Lowercase keyword -> canonical capitalization. Covers all single-token
 * reserved identifiers (MS-VBAL section 3.3.5.2) plus VBE-convention contextual
 * keywords. Does not include reserved-for-implementation-use names, which remain
 * raw Attribute-line metadata spellings even though `isReservedIdentifier`
 * treats them as reserved for declaration validation.
 */
export const VBA_KEYWORDS: Readonly<Record<string, string>> = buildKeywordMap();

/**
 * Returns the canonical capitalization for a keyword, or undefined if the word
 * is not a known VBA keyword. Matching is case-insensitive per MS-VBAL section
 * 3.3.5.2.
 */
export function canonicalKeyword(word: string): string | undefined {
	return VBA_KEYWORDS[word.toLowerCase()];
}

/** True if the word is a reserved identifier (MS-VBAL section 3.3.5.2). */
export function isReservedIdentifier(word: string): boolean {
	return RESERVED_IDENTIFIERS.has(word.toLowerCase());
}

