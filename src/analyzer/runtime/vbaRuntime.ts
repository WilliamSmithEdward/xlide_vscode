// Built-in VBA runtime metadata (MS-VBAL Phase 9).
//
// A curated, verified subset of the VBA runtime library: the intrinsic
// functions, statements, constants, and global objects available in every VBA
// project regardless of host (MsgBox, Left, CLng, Now, Array, RGB, Err, ...).
// Signatures are transcribed from the Microsoft VBA language reference
// (learn.microsoft.com/office/vba/language) and MS-VBAL; they are never
// LLM-invented.
//
// Names that collide with intrinsic data types or are otherwise context
// ambiguous (Date, Time, String, Error) are intentionally omitted so that
// hovering a type in an `As` position is never mistaken for a function call.
//
// Pure analyzer data: depends on nothing and is reused by hover, identifier
// completion, and (later) signature help. See Phase 9 in
// docs/xlide_vba_language_service_roadmap.md.

/** A built-in VBA runtime function or statement. */
export interface VbaRuntimeFunction {
	/** Canonical VBE name (correct capitalization). */
	name: string;
	/** Full signature line, e.g. `Left(String, Length) As String`. */
	signature: string;
	/** Declared return type, when the entry is a function. */
	returns?: string;
	/** `function` returns a value; `statement` is invoked for effect only. */
	kind: 'function' | 'statement';
	/** Provenance marker - always verified against Microsoft VBA docs. */
	source: 'verified';
	/** Explicitly curated parameter types. Missing entries are treated unknown. */
	params?: readonly VbaRuntimeParam[];
	/** Explicit `Call` statement compatibility. Omitted means allowed. */
	explicitCall?: 'allowed' | 'forbidden';
}

/** Explicit metadata for one runtime parameter. Never inferred from names. */
export interface VbaRuntimeParam {
	name: string;
	type?: string;
	optional?: boolean;
	paramArray?: boolean;
}

/** A built-in VBA runtime constant or enum member. */
export interface VbaRuntimeConstant {
	name: string;
	type?: string;
	value?: string | number;
	source: 'verified';
}

export interface VbaRuntimeObjectMember {
	name: string;
	kind: 'property' | 'method';
	returns?: string;
	signature?: string;
	writable?: boolean;
	writeType?: string;
}

/** A built-in global VBA runtime object, such as `Err`. */
export interface VbaRuntimeObject {
	name: string;
	type: string;
	source: 'verified';
	exhaustive: true;
	members: readonly VbaRuntimeObjectMember[];
}

function fn(
	name: string,
	signature: string,
	returns?: string,
	params?: readonly VbaRuntimeParam[],
	options?: Pick<VbaRuntimeFunction, 'explicitCall'>,
): VbaRuntimeFunction {
	return { name, signature, returns, kind: 'function', source: 'verified', params, ...options };
}

function stmt(
	name: string,
	signature: string,
	params?: readonly VbaRuntimeParam[],
	options?: Pick<VbaRuntimeFunction, 'explicitCall'>,
): VbaRuntimeFunction {
	return { name, signature, kind: 'statement', source: 'verified', params, ...options };
}

function c(name: string, type?: string, value?: string | number): VbaRuntimeConstant {
	return { name, type, value, source: 'verified' };
}

function prop(
	name: string,
	type: string,
	options: Pick<VbaRuntimeObjectMember, 'writable'> = { writable: true },
): VbaRuntimeObjectMember {
	return {
		name,
		kind: 'property',
		signature: `${name} As ${type}`,
		writeType: type,
		...options,
	};
}

function method(name: string, signature: string): VbaRuntimeObjectMember {
	return { name, kind: 'method', signature };
}

const LEFT_PARAMS = [
	{ name: 'String', type: 'String' },
	{ name: 'Length', type: 'Long' },
] as const;
const RIGHT_PARAMS = [
	{ name: 'String', type: 'String' },
	{ name: 'Length', type: 'Long' },
] as const;
const MID_PARAMS = [
	{ name: 'String', type: 'String' },
	{ name: 'Start', type: 'Long' },
	{ name: 'Length', type: 'Long', optional: true },
] as const;
const SPACE_PARAMS = [
	{ name: 'Number', type: 'Long' },
] as const;
const STRING_PARAMS = [
	{ name: 'Number', type: 'Long' },
	{ name: 'Character', type: 'Variant' },
] as const;
const REPLACE_PARAMS = [
	{ name: 'Expression', type: 'String' },
	{ name: 'Find', type: 'String' },
	{ name: 'Replace', type: 'String' },
	{ name: 'Start', type: 'Long', optional: true },
	{ name: 'Count', type: 'Long', optional: true },
	{ name: 'Compare', type: 'VbCompareMethod', optional: true },
] as const;

/** The verified built-in VBA runtime functions and statements. */
export const VBA_RUNTIME_FUNCTIONS: VbaRuntimeFunction[] = [
	// -- Interaction --------------------------------------------------------
	fn(
		'MsgBox',
		'MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context]) As VbMsgBoxResult',
		'VbMsgBoxResult',
	),
	fn(
		'InputBox',
		'InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context]) As String',
		'String',
	),
	fn('Environ', 'Environ(Expression) As String', 'String'),
	fn('Shell', 'Shell(PathName, [WindowStyle As VbAppWinStyle = vbMinimizedFocus]) As Double', 'Double'),
	fn('DoEvents', 'DoEvents() As Integer', 'Integer', undefined, { explicitCall: 'forbidden' }),
	fn('Erl', 'Erl() As Integer', 'Integer', undefined, { explicitCall: 'forbidden' }),
	// The FUNCTION form. `Error 5` - the statement that raises one - is grammar
	// handled by the parser; `Error` and `Error(5)` read the message text for a
	// code, defaulting to the current Err.Number.
	fn('Error', 'Error([ErrorNumber]) As String', 'String'),
	fn('Error$', 'Error$([ErrorNumber]) As String', 'String'),
	fn('CreateObject', 'CreateObject(Class, [ServerName]) As Object', 'Object'),
	fn('GetObject', 'GetObject([PathName], [Class]) As Object', 'Object'),
	stmt('Beep', 'Beep'),
	stmt('Randomize', 'Randomize [Number]'),
	stmt('Load', 'Load Object'),
	stmt('Unload', 'Unload Object'),
	fn('IMEStatus', 'IMEStatus() As Integer', 'Integer'),
	// Documented VBA functions that answer only on Mac hosts. They are part of
	// the language reference either way, and a Windows-only analyzer reporting
	// them as undeclared is a false positive on portable code (issue #41).
	fn('MacID', 'MacID(Constant) As Long', 'Long'),
	fn('MacScript', 'MacScript(Script) As Variant', 'Variant'),

	// -- String functions ---------------------------------------------------
	fn('Len', 'Len(Expression) As Long', 'Long'),
	fn('Left', 'Left(String, Length) As String', 'String', LEFT_PARAMS),
	fn('Left$', 'Left$(String, Length) As String', 'String', LEFT_PARAMS),
	fn('Right', 'Right(String, Length) As String', 'String', RIGHT_PARAMS),
	fn('Right$', 'Right$(String, Length) As String', 'String', RIGHT_PARAMS),
	fn('Mid', 'Mid(String, Start, [Length]) As String', 'String', MID_PARAMS),
	fn('Mid$', 'Mid$(String, Start, [Length]) As String', 'String', MID_PARAMS),
	fn('Trim', 'Trim(String) As String', 'String'),
	fn('LTrim', 'LTrim(String) As String', 'String'),
	fn('RTrim', 'RTrim(String) As String', 'String'),
	fn('UCase', 'UCase(String) As String', 'String'),
	fn('LCase', 'LCase(String) As String', 'String'),
	fn('Replace', 'Replace(Expression, Find, Replace, [Start = 1], [Count = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String', 'String', REPLACE_PARAMS),
	fn('Replace$', 'Replace$(Expression, Find, Replace, [Start = 1], [Count = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String', 'String', REPLACE_PARAMS),
	fn('InStr', 'InStr([Start], String1, String2, [Compare As VbCompareMethod]) As Long', 'Long'),
	fn('InStrRev', 'InStrRev(StringCheck, StringMatch, [Start = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As Long', 'Long'),
	fn('Split', 'Split(Expression, [Delimiter], [Limit = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String()', 'String()'),
	fn('Join', 'Join(SourceArray, [Delimiter]) As String', 'String'),
	fn('StrComp', 'StrComp(String1, String2, [Compare As VbCompareMethod]) As Integer', 'Integer'),
	fn('StrConv', 'StrConv(String, Conversion As VbStrConv, [LCID]) As String', 'String'),
	fn('String$', 'String$(Number, Character) As String', 'String', STRING_PARAMS),
	fn('Space', 'Space(Number) As String', 'String', SPACE_PARAMS),
	fn('Space$', 'Space$(Number) As String', 'String', SPACE_PARAMS),
	fn('Format', 'Format(Expression, [Format], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As String', 'String'),
	fn('Chr', 'Chr(CharCode) As String', 'String'),
	fn('ChrW', 'ChrW(CharCode) As String', 'String'),
	fn('Asc', 'Asc(String) As Integer', 'Integer'),
	fn('AscW', 'AscW(String) As Integer', 'Integer'),

	// -- Byte-string functions (the ...B variants operate on byte positions) --
	fn('LenB', 'LenB(Expression) As Long', 'Long'),
	fn('LeftB', 'LeftB(String, Length) As String', 'String'),
	fn('LeftB$', 'LeftB$(String, Length) As String', 'String'),
	fn('RightB', 'RightB(String, Length) As String', 'String'),
	fn('RightB$', 'RightB$(String, Length) As String', 'String'),
	fn('MidB', 'MidB(String, Start, [Length]) As String', 'String'),
	fn('MidB$', 'MidB$(String, Start, [Length]) As String', 'String'),
	fn('InStrB', 'InStrB([Start], String1, String2, [Compare As VbCompareMethod]) As Long', 'Long'),
	fn('AscB', 'AscB(String) As Integer', 'Integer'),
	fn('ChrB', 'ChrB(CharCode) As String', 'String'),
	fn('ChrB$', 'ChrB$(CharCode) As String', 'String'),

	// -- Pointer / memory (hidden but always available in VBA) --------------
	fn('VarPtr', 'VarPtr(Variable) As LongPtr', 'LongPtr'),
	fn('StrPtr', 'StrPtr(StringVariable) As LongPtr', 'LongPtr'),
	fn('ObjPtr', 'ObjPtr(ObjectVariable) As LongPtr', 'LongPtr'),

	// -- Type conversion ----------------------------------------------------
	fn('CBool', 'CBool(Expression) As Boolean', 'Boolean'),
	fn('CByte', 'CByte(Expression) As Byte', 'Byte'),
	fn('CCur', 'CCur(Expression) As Currency', 'Currency'),
	fn('CDate', 'CDate(Expression) As Date', 'Date'),
	fn('CDbl', 'CDbl(Expression) As Double', 'Double'),
	fn('CDec', 'CDec(Expression) As Variant', 'Variant'),
	fn('CInt', 'CInt(Expression) As Integer', 'Integer'),
	fn('CLng', 'CLng(Expression) As Long', 'Long'),
	fn('CLngLng', 'CLngLng(Expression) As LongLong', 'LongLong'),
	fn('CSng', 'CSng(Expression) As Single', 'Single'),
	fn('CStr', 'CStr(Expression) As String', 'String'),
	fn('CVar', 'CVar(Expression) As Variant', 'Variant'),
	fn('Val', 'Val(String) As Double', 'Double'),
	fn('Hex', 'Hex(Number) As String', 'String'),
	fn('Oct', 'Oct(Number) As String', 'String'),

	// CVDate is the legacy conversion function the type-conversion reference
	// still documents beside CDate.
	fn('CVDate', 'CVDate(Expression) As Date', 'Date'),

	// -- Math ---------------------------------------------------------------
	fn('Abs', 'Abs(Number) As Variant', 'Variant'),
	fn('Int', 'Int(Number) As Variant', 'Variant'),
	fn('Fix', 'Fix(Number) As Variant', 'Variant'),
	fn('Sgn', 'Sgn(Number) As Integer', 'Integer'),
	fn('Sqr', 'Sqr(Number) As Double', 'Double'),
	fn('Exp', 'Exp(Number) As Double', 'Double'),
	fn('Log', 'Log(Number) As Double', 'Double'),
	fn('Sin', 'Sin(Number) As Double', 'Double'),
	fn('Cos', 'Cos(Number) As Double', 'Double'),
	fn('Tan', 'Tan(Number) As Double', 'Double'),
	fn('Atn', 'Atn(Number) As Double', 'Double'),
	fn('Round', 'Round(Number, [NumDigitsAfterDecimal]) As Double', 'Double'),
	fn('Rnd', 'Rnd([Number]) As Single', 'Single'),

	// -- Date / time --------------------------------------------------------
	fn('Now', 'Now() As Date', 'Date'),
	// `Time` is both a function and a statement (`Time = #12:00:00 PM#`); the
	// name is what the identifier rules need, and the read form is the one
	// with a signature to show (issue #41).
	fn('Time', 'Time() As Date', 'Date'),
	fn('Timer', 'Timer() As Single', 'Single'),
	fn('Year', 'Year(Date) As Integer', 'Integer'),
	fn('Month', 'Month(Date) As Integer', 'Integer'),
	fn('Day', 'Day(Date) As Integer', 'Integer'),
	fn('Hour', 'Hour(Time) As Integer', 'Integer'),
	fn('Minute', 'Minute(Time) As Integer', 'Integer'),
	fn('Second', 'Second(Time) As Integer', 'Integer'),
	fn('Weekday', 'Weekday(Date, [FirstDayOfWeek As VbDayOfWeek = vbSunday]) As Integer', 'Integer'),
	fn('MonthName', 'MonthName(Month, [Abbreviate As Boolean = False]) As String', 'String'),
	fn('WeekdayName', 'WeekdayName(Weekday, [Abbreviate As Boolean = False], [FirstDayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek]) As String', 'String'),
	fn('DateAdd', 'DateAdd(Interval, Number, Date) As Date', 'Date'),
	fn('DateDiff', 'DateDiff(Interval, Date1, Date2, [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Long', 'Long'),
	fn('DatePart', 'DatePart(Interval, Date, [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Integer', 'Integer'),
	fn('DateSerial', 'DateSerial(Year, Month, Day) As Date', 'Date'),
	fn('TimeSerial', 'TimeSerial(Hour, Minute, Second) As Date', 'Date'),
	fn('DateValue', 'DateValue(Date) As Date', 'Date'),
	fn('TimeValue', 'TimeValue(Time) As Date', 'Date'),

	// -- Arrays / inspection ------------------------------------------------
	fn('Array', 'Array(ArgList) As Variant', 'Variant', [
		{ name: 'ArgList', type: 'Variant', optional: true, paramArray: true },
	]),
	fn('LBound', 'LBound(ArrayName, [Dimension = 1]) As Long', 'Long'),
	fn('UBound', 'UBound(ArrayName, [Dimension = 1]) As Long', 'Long'),
	fn('IsArray', 'IsArray(VarName) As Boolean', 'Boolean'),
	fn('IsDate', 'IsDate(Expression) As Boolean', 'Boolean'),
	fn('IsEmpty', 'IsEmpty(Expression) As Boolean', 'Boolean'),
	fn('IsError', 'IsError(Expression) As Boolean', 'Boolean'),
	fn('IsNull', 'IsNull(Expression) As Boolean', 'Boolean'),
	fn('IsNumeric', 'IsNumeric(Expression) As Boolean', 'Boolean'),
	fn('IsObject', 'IsObject(Expression) As Boolean', 'Boolean'),
	fn('VarType', 'VarType(VarName) As VbVarType', 'VbVarType'),
	fn('TypeName', 'TypeName(VarName) As String', 'String'),

	// -- Selection / colour -------------------------------------------------
	fn('RGB', 'RGB(Red, Green, Blue) As Long', 'Long'),
	fn('QBColor', 'QBColor(Color) As Long', 'Long'),
	fn('IIf', 'IIf(Expression, TruePart, FalsePart) As Variant', 'Variant'),
	fn('Choose', 'Choose(Index, ArgList) As Variant', 'Variant'),
	fn('Switch', 'Switch(ArgList) As Variant', 'Variant'),
	fn('IsMissing', 'IsMissing(ArgName) As Boolean', 'Boolean'),
	fn('CallByName', 'CallByName(Object, ProcName As String, CallType As VbCallType, [Args]) As Variant', 'Variant', [
		{ name: 'Object', type: 'Object' },
		{ name: 'ProcName', type: 'String' },
		{ name: 'CallType', type: 'VbCallType' },
		{ name: 'Args', type: 'Variant', optional: true, paramArray: true },
	]),

	// -- Additional string functions ---------------------------------------
	fn('StrReverse', 'StrReverse(Expression) As String', 'String'),
	fn('Str', 'Str(Number) As String', 'String'),
	fn('Filter', 'Filter(SourceArray, Match, [Include As Boolean = True], [Compare As VbCompareMethod = vbBinaryCompare]) As String()', 'String()'),
	fn('FormatCurrency', 'FormatCurrency(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit As VbTriState = vbUseDefault], [UseParensForNegativeNumbers As VbTriState = vbUseDefault], [GroupDigits As VbTriState = vbUseDefault]) As String', 'String'),
	fn('FormatNumber', 'FormatNumber(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit As VbTriState = vbUseDefault], [UseParensForNegativeNumbers As VbTriState = vbUseDefault], [GroupDigits As VbTriState = vbUseDefault]) As String', 'String'),
	fn('FormatPercent', 'FormatPercent(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit As VbTriState = vbUseDefault], [UseParensForNegativeNumbers As VbTriState = vbUseDefault], [GroupDigits As VbTriState = vbUseDefault]) As String', 'String'),
	fn('FormatDateTime', 'FormatDateTime(Date, [NamedFormat As VbDateTimeFormat = vbGeneralDate]) As String', 'String'),

	// -- Conversion helpers -------------------------------------------------
	fn('CVErr', 'CVErr(ErrorNumber) As Variant', 'Variant'),

	// -- Interaction / registry --------------------------------------------
	fn('Command', 'Command() As String', 'String'),
	fn('Partition', 'Partition(Number, Start, Stop, Interval) As String', 'String'),
	fn('GetSetting', 'GetSetting(AppName, Section, Key, [Default]) As String', 'String'),
	fn('GetAllSettings', 'GetAllSettings(AppName, Section) As Variant', 'Variant'),
	stmt('SaveSetting', 'SaveSetting AppName, Section, Key, Setting'),
	stmt('DeleteSetting', 'DeleteSetting AppName, [Section], [Key]'),
	stmt('AppActivate', 'AppActivate Title, [Wait As Boolean]'),
	stmt('SendKeys', 'SendKeys String, [Wait]'),

	// -- File system --------------------------------------------------------
	fn('Dir', 'Dir([PathName], [Attributes As VbFileAttribute = vbNormal]) As String', 'String'),
	fn('FreeFile', 'FreeFile([RangeNumber]) As Integer', 'Integer'),
	fn('EOF', 'EOF(FileNumber) As Boolean', 'Boolean'),
	fn('LOF', 'LOF(FileNumber) As Long', 'Long'),
	fn('Loc', 'Loc(FileNumber) As Long', 'Long'),
	fn('Seek', 'Seek(FileNumber) As Long', 'Long'),
	fn('FileLen', 'FileLen(PathName) As Long', 'Long'),
	fn('FileDateTime', 'FileDateTime(PathName) As Date', 'Date'),
	fn('GetAttr', 'GetAttr(PathName) As VbFileAttribute', 'VbFileAttribute'),
	fn('FileAttr', 'FileAttr(FileNumber, ReturnType) As Long', 'Long'),
	fn('CurDir', 'CurDir([Drive]) As String', 'String'),
	stmt('ChDir', 'ChDir Path'),
	stmt('ChDrive', 'ChDrive Drive'),
	stmt('MkDir', 'MkDir Path'),
	stmt('RmDir', 'RmDir Path'),
	stmt('Kill', 'Kill PathName'),
	stmt('FileCopy', 'FileCopy Source, Destination'),
	// Two words, one identifier as far as the rules are concerned.
	stmt('Line', 'Line Input #FileNumber, VarName'),
	stmt('SetAttr', 'SetAttr PathName, Attributes As VbFileAttribute'),
	stmt('Name', 'Name OldPathName As NewPathName'),
	stmt('Reset', 'Reset'),
	stmt('Width', 'Width #FileNumber, Width'),

	// -- Financial ----------------------------------------------------------
	fn('PV', 'PV(Rate, NPer, Pmt, [FV = 0], [Type = 0]) As Double', 'Double'),
	fn('FV', 'FV(Rate, NPer, Pmt, [PV = 0], [Type = 0]) As Double', 'Double'),
	fn('Pmt', 'Pmt(Rate, NPer, PV, [FV = 0], [Type = 0]) As Double', 'Double'),
	fn('IPmt', 'IPmt(Rate, Per, NPer, PV, [FV = 0], [Type = 0]) As Double', 'Double'),
	fn('PPmt', 'PPmt(Rate, Per, NPer, PV, [FV = 0], [Type = 0]) As Double', 'Double'),
	fn('NPer', 'NPer(Rate, Pmt, PV, [FV = 0], [Type = 0]) As Double', 'Double'),
	fn('Rate', 'Rate(NPer, Pmt, PV, [FV = 0], [Type = 0], [Guess = 0.1]) As Double', 'Double'),
	fn('NPV', 'NPV(Rate, ValueArray()) As Double', 'Double'),
	fn('IRR', 'IRR(ValueArray(), [Guess = 0.1]) As Double', 'Double'),
	fn('MIRR', 'MIRR(ValueArray(), FinanceRate, ReinvestRate) As Double', 'Double'),
	fn('SLN', 'SLN(Cost, Salvage, Life) As Double', 'Double'),
	fn('SYD', 'SYD(Cost, Salvage, Life, Period) As Double', 'Double'),
	fn('DDB', 'DDB(Cost, Salvage, Life, Period, [Factor = 2]) As Double', 'Double'),
];

/** The verified built-in VBA runtime constants and enum members. */
export const VBA_RUNTIME_CONSTANTS: VbaRuntimeConstant[] = [
	// Documented in the VBA constants reference (learn.microsoft.com/en-us/
	// office/vba/language/reference/constants-visual-basic-for-applications)
	// but absent until issue #41: every one of these read as an undeclared
	// variable under Option Explicit. Values are the type library's where it
	// declares an enumeration, and the reference tables' where it does not.
	// CallType constants (VbCallType)
	c('vbGet', 'VbCallType', 2),
	c('vbLet', 'VbCallType', 4),
	c('vbMethod', 'VbCallType', 1),
	c('vbSet', 'VbCallType', 8),

	// Comparison constants (VbCompareMethod)
	c('vbUseCompareOption', 'VbCompareMethod', -1),

	// QueryClose constants (VbQueryClose)
	c('vbAppTaskManager', 'VbQueryClose', 3),
	c('vbAppWindows', 'VbQueryClose', 2),
	c('vbFormCode', 'VbQueryClose', 1),
	c('vbFormControlMenu', 'VbQueryClose', 0),

	// Form constants
	c('vbModal', undefined, 1),
	c('vbModeless', undefined, 0),

	// Colour constants
	c('vbBlack', undefined, 0),
	c('vbBlue', undefined, 16711680),
	c('vbCyan', undefined, 16776960),
	c('vbGreen', undefined, 65280),
	c('vbMagenta', undefined, 16711935),
	c('vbRed', undefined, 255),
	c('vbWhite', undefined, 16777215),
	c('vbYellow', undefined, 65535),

	// System colour constants
	c('vb3DDKShadow', undefined, -2147483627),
	c('vb3DHighlight', undefined, -2147483628),
	c('vb3DLight', undefined, -2147483626),
	c('vbActiveBorder', undefined, -2147483638),
	c('vbActiveTitleBar', undefined, -2147483646),
	c('vbApplicationWorkspace', undefined, -2147483636),
	c('vbButtonFace', undefined, -2147483633),
	c('vbButtonShadow', undefined, -2147483632),
	c('vbButtonText', undefined, -2147483630),
	c('vbDesktop', undefined, -2147483647),
	c('vbGrayText', undefined, -2147483631),
	c('vbHighlight', undefined, -2147483635),
	c('vbHighlightText', undefined, -2147483634),
	c('vbInactiveBorder', undefined, -2147483637),
	c('vbInactiveCaptionText', undefined, -2147483629),
	c('vbInactiveTitleBar', undefined, -2147483645),
	c('vbInfoBackground', undefined, -2147483624),
	c('vbInfoText', undefined, -2147483625),
	c('vbMenuBar', undefined, -2147483644),
	c('vbMenuText', undefined, -2147483641),
	c('vbScrollBars', undefined, -2147483648),
	c('vbTitleBarText', undefined, -2147483639),
	c('vbWindowBackground', undefined, -2147483643),
	c('vbWindowFrame', undefined, -2147483642),
	c('vbWindowText', undefined, -2147483640),

	// Keycode constants - no enumeration is documented for these
	c('vbKey0', undefined, 48),
	c('vbKey1', undefined, 49),
	c('vbKey2', undefined, 50),
	c('vbKey3', undefined, 51),
	c('vbKey4', undefined, 52),
	c('vbKey5', undefined, 53),
	c('vbKey6', undefined, 54),
	c('vbKey7', undefined, 55),
	c('vbKey8', undefined, 56),
	c('vbKey9', undefined, 57),
	c('vbKeyA', undefined, 65),
	c('vbKeyAdd', undefined, 107),
	c('vbKeyB', undefined, 66),
	c('vbKeyBack', undefined, 8),
	c('vbKeyC', undefined, 67),
	c('vbKeyCancel', undefined, 3),
	c('vbKeyCapital', undefined, 20),
	c('vbKeyClear', undefined, 12),
	c('vbKeyControl', undefined, 17),
	c('vbKeyD', undefined, 68),
	c('vbKeyDecimal', undefined, 110),
	c('vbKeyDelete', undefined, 46),
	c('vbKeyDivide', undefined, 111),
	c('vbKeyDown', undefined, 40),
	c('vbKeyE', undefined, 69),
	c('vbKeyEnd', undefined, 35),
	c('vbKeyEscape', undefined, 27),
	c('vbKeyExecute', undefined, 43),
	c('vbKeyF', undefined, 70),
	c('vbKeyF1', undefined, 112),
	c('vbKeyF10', undefined, 121),
	c('vbKeyF11', undefined, 122),
	c('vbKeyF12', undefined, 123),
	c('vbKeyF13', undefined, 124),
	c('vbKeyF14', undefined, 125),
	c('vbKeyF15', undefined, 126),
	c('vbKeyF16', undefined, 127),
	c('vbKeyF2', undefined, 113),
	c('vbKeyF3', undefined, 114),
	c('vbKeyF4', undefined, 115),
	c('vbKeyF5', undefined, 116),
	c('vbKeyF6', undefined, 117),
	c('vbKeyF7', undefined, 118),
	c('vbKeyF8', undefined, 119),
	c('vbKeyF9', undefined, 120),
	c('vbKeyG', undefined, 71),
	c('vbKeyH', undefined, 72),
	c('vbKeyHelp', undefined, 47),
	c('vbKeyHome', undefined, 36),
	c('vbKeyI', undefined, 73),
	c('vbKeyInsert', undefined, 45),
	c('vbKeyJ', undefined, 74),
	c('vbKeyK', undefined, 75),
	c('vbKeyL', undefined, 76),
	c('vbKeyLButton', undefined, 1),
	c('vbKeyLeft', undefined, 37),
	c('vbKeyM', undefined, 77),
	c('vbKeyMButton', undefined, 4),
	c('vbKeyMenu', undefined, 18),
	c('vbKeyMultiply', undefined, 106),
	c('vbKeyN', undefined, 78),
	c('vbKeyNumlock', undefined, 144),
	c('vbKeyNumpad0', undefined, 96),
	c('vbKeyNumpad1', undefined, 97),
	c('vbKeyNumpad2', undefined, 98),
	c('vbKeyNumpad3', undefined, 99),
	c('vbKeyNumpad4', undefined, 100),
	c('vbKeyNumpad5', undefined, 101),
	c('vbKeyNumpad6', undefined, 102),
	c('vbKeyNumpad7', undefined, 103),
	c('vbKeyNumpad8', undefined, 104),
	c('vbKeyNumpad9', undefined, 105),
	c('vbKeyO', undefined, 79),
	c('vbKeyP', undefined, 80),
	c('vbKeyPageDown', undefined, 34),
	c('vbKeyPageUp', undefined, 33),
	c('vbKeyPause', undefined, 19),
	c('vbKeyPrint', undefined, 42),
	c('vbKeyQ', undefined, 81),
	c('vbKeyR', undefined, 82),
	c('vbKeyRButton', undefined, 2),
	c('vbKeyReturn', undefined, 13),
	c('vbKeyRight', undefined, 39),
	c('vbKeyS', undefined, 83),
	c('vbKeySelect', undefined, 41),
	c('vbKeySeparator', undefined, 108),
	c('vbKeyShift', undefined, 16),
	c('vbKeySnapshot', undefined, 44),
	c('vbKeySpace', undefined, 32),
	c('vbKeySubtract', undefined, 109),
	c('vbKeyT', undefined, 84),
	c('vbKeyTab', undefined, 9),
	c('vbKeyU', undefined, 85),
	c('vbKeyUp', undefined, 38),
	c('vbKeyV', undefined, 86),
	c('vbKeyW', undefined, 87),
	c('vbKeyX', undefined, 88),
	c('vbKeyY', undefined, 89),
	c('vbKeyZ', undefined, 90),

	c('vbObjectError', 'Long', -2147221504),
	c('vbNullString', 'String'),
	c('vbNullChar', 'String'),
	c('vbCrLf', 'String'),
	c('vbNewLine', 'String'),
	c('vbCr', 'String'),
	c('vbLf', 'String'),
	c('vbBack', 'String'),
	c('vbFormFeed', 'String'),
	c('vbTab', 'String'),
	c('vbVerticalTab', 'String'),

	c('vbOKOnly', 'VbMsgBoxStyle', 0),
	c('vbOKCancel', 'VbMsgBoxStyle', 1),
	c('vbAbortRetryIgnore', 'VbMsgBoxStyle', 2),
	c('vbYesNoCancel', 'VbMsgBoxStyle', 3),
	c('vbYesNo', 'VbMsgBoxStyle', 4),
	c('vbRetryCancel', 'VbMsgBoxStyle', 5),
	c('vbCritical', 'VbMsgBoxStyle', 16),
	c('vbQuestion', 'VbMsgBoxStyle', 32),
	c('vbExclamation', 'VbMsgBoxStyle', 48),
	c('vbInformation', 'VbMsgBoxStyle', 64),
	c('vbDefaultButton1', 'VbMsgBoxStyle', 0),
	c('vbDefaultButton2', 'VbMsgBoxStyle', 256),
	c('vbDefaultButton3', 'VbMsgBoxStyle', 512),
	c('vbDefaultButton4', 'VbMsgBoxStyle', 768),
	c('vbApplicationModal', 'VbMsgBoxStyle', 0),
	c('vbSystemModal', 'VbMsgBoxStyle', 4096),
	c('vbMsgBoxHelpButton', 'VbMsgBoxStyle', 16384),
	c('vbMsgBoxRight', 'VbMsgBoxStyle', 524288),
	c('vbMsgBoxRtlReading', 'VbMsgBoxStyle', 1048576),
	c('vbMsgBoxSetForeground', 'VbMsgBoxStyle', 65536),

	c('vbOK', 'VbMsgBoxResult', 1),
	c('vbCancel', 'VbMsgBoxResult', 2),
	c('vbAbort', 'VbMsgBoxResult', 3),
	c('vbRetry', 'VbMsgBoxResult', 4),
	c('vbIgnore', 'VbMsgBoxResult', 5),
	c('vbYes', 'VbMsgBoxResult', 6),
	c('vbNo', 'VbMsgBoxResult', 7),

	c('vbBinaryCompare', 'VbCompareMethod', 0),
	c('vbTextCompare', 'VbCompareMethod', 1),
	c('vbDatabaseCompare', 'VbCompareMethod', 2),
	c('BinaryCompare', 'VbCompareMethod', 0),
	c('TextCompare', 'VbCompareMethod', 1),
	c('DatabaseCompare', 'VbCompareMethod', 2),

	c('vbUseDefault', 'VbTriState', -2),
	c('vbTrue', 'VbTriState', -1),
	c('vbFalse', 'VbTriState', 0),

	c('vbGeneralDate', 'VbDateTimeFormat', 0),
	c('vbLongDate', 'VbDateTimeFormat', 1),
	c('vbShortDate', 'VbDateTimeFormat', 2),
	c('vbLongTime', 'VbDateTimeFormat', 3),
	c('vbShortTime', 'VbDateTimeFormat', 4),

	c('vbUseSystemDayOfWeek', 'VbDayOfWeek', 0),
	c('vbSunday', 'VbDayOfWeek', 1),
	c('vbMonday', 'VbDayOfWeek', 2),
	c('vbTuesday', 'VbDayOfWeek', 3),
	c('vbWednesday', 'VbDayOfWeek', 4),
	c('vbThursday', 'VbDayOfWeek', 5),
	c('vbFriday', 'VbDayOfWeek', 6),
	c('vbSaturday', 'VbDayOfWeek', 7),

	c('vbUseSystem', 'VbFirstWeekOfYear', 0),
	c('vbFirstJan1', 'VbFirstWeekOfYear', 1),
	c('vbFirstFourDays', 'VbFirstWeekOfYear', 2),
	c('vbFirstFullWeek', 'VbFirstWeekOfYear', 3),

	c('vbHide', 'VbAppWinStyle', 0),
	c('vbNormalFocus', 'VbAppWinStyle', 1),
	c('vbMinimizedFocus', 'VbAppWinStyle', 2),
	c('vbMaximizedFocus', 'VbAppWinStyle', 3),
	c('vbNormalNoFocus', 'VbAppWinStyle', 4),
	c('vbMinimizedNoFocus', 'VbAppWinStyle', 6),

	c('vbNormal', 'VbFileAttribute', 0),
	c('vbReadOnly', 'VbFileAttribute', 1),
	c('vbHidden', 'VbFileAttribute', 2),
	c('vbSystem', 'VbFileAttribute', 4),
	c('vbVolume', 'VbFileAttribute', 8),
	c('vbDirectory', 'VbFileAttribute', 16),
	c('vbArchive', 'VbFileAttribute', 32),
	c('vbAlias', 'VbFileAttribute', 64),

	c('VbMethod', 'VbCallType', 1),
	c('VbGet', 'VbCallType', 2),
	c('VbLet', 'VbCallType', 4),
	c('VbSet', 'VbCallType', 8),

	c('vbEmpty', 'VbVarType', 0),
	c('vbNull', 'VbVarType', 1),
	c('vbInteger', 'VbVarType', 2),
	c('vbLong', 'VbVarType', 3),
	c('vbSingle', 'VbVarType', 4),
	c('vbDouble', 'VbVarType', 5),
	c('vbCurrency', 'VbVarType', 6),
	c('vbDate', 'VbVarType', 7),
	c('vbString', 'VbVarType', 8),
	c('vbObject', 'VbVarType', 9),
	c('vbError', 'VbVarType', 10),
	c('vbBoolean', 'VbVarType', 11),
	c('vbVariant', 'VbVarType', 12),
	c('vbDataObject', 'VbVarType', 13),
	c('vbDecimal', 'VbVarType', 14),
	c('vbByte', 'VbVarType', 17),
	c('vbLongLong', 'VbVarType', 20),
	c('vbUserDefinedType', 'VbVarType', 36),
	c('vbArray', 'VbVarType', 8192),

	c('vbUpperCase', 'VbStrConv', 1),
	c('vbLowerCase', 'VbStrConv', 2),
	c('vbProperCase', 'VbStrConv', 3),
	c('vbWide', 'VbStrConv', 4),
	c('vbNarrow', 'VbStrConv', 8),
	c('vbKatakana', 'VbStrConv', 16),
	c('vbHiragana', 'VbStrConv', 32),
	c('vbUnicode', 'VbStrConv', 64),
	c('vbFromUnicode', 'VbStrConv', 128),

	c('vbIMENoOp', 'VbIMEStatus', 0),
	c('vbIMEModeNoControl', 'VbIMEStatus', 0),
	c('vbIMEOn', 'VbIMEStatus', 1),
	c('vbIMEModeOn', 'VbIMEStatus', 1),
	c('vbIMEOff', 'VbIMEStatus', 2),
	c('vbIMEModeOff', 'VbIMEStatus', 2),
	c('vbIMEDisable', 'VbIMEStatus', 3),
	c('vbIMEModeDisable', 'VbIMEStatus', 3),
	c('vbIMEHiragana', 'VbIMEStatus', 4),
	c('vbIMEModeHiragana', 'VbIMEStatus', 4),
	c('vbIMEKatakanaDbl', 'VbIMEStatus', 5),
	c('vbIMEModeKatakana', 'VbIMEStatus', 5),
	c('vbIMEKatakanaSng', 'VbIMEStatus', 6),
	c('vbIMEModeKatakanaHalf', 'VbIMEStatus', 6),
	c('vbIMEAlphaDbl', 'VbIMEStatus', 7),
	c('vbIMEModeAlphaFull', 'VbIMEStatus', 7),
	c('vbIMEAlphaSng', 'VbIMEStatus', 8),
	c('vbIMEModeAlpha', 'VbIMEStatus', 8),
	c('vbIMEModeHangulFull', 'VbIMEStatus', 9),
	c('vbIMEModeHangul', 'VbIMEStatus', 10),

	c('vbCalGreg', 'VbCalendar', 0),
	c('vbCalHijri', 'VbCalendar', 1),
];

/** The verified built-in VBA runtime global objects. */
export const VBA_RUNTIME_OBJECTS: VbaRuntimeObject[] = [
	{
		name: 'Debug',
		type: 'VBA.DebugObject',
		source: 'verified',
		exhaustive: true,
		members: [
			method('Assert', 'Assert(BooleanExpression As Boolean)'),
			method('Print', 'Print(ParamArray OutputList As Variant)'),
		],
	},
	{
		name: 'Err',
		type: 'VBA.ErrObject',
		source: 'verified',
		exhaustive: true,
		members: [
			method('Clear', 'Clear()'),
			prop('Description', 'String'),
			prop('HelpContext', 'Long'),
			prop('HelpFile', 'String'),
			prop('LastDLLError', 'Long', { writable: false }),
			prop('Number', 'Long'),
			method('Raise', 'Raise(Number As Long, [Source], [Description], [HelpFile], [HelpContext])'),
			prop('Source', 'String'),
		],
	},
];

const BY_LOWER = new Map<string, VbaRuntimeFunction>(
	VBA_RUNTIME_FUNCTIONS.map((f) => [f.name.toLowerCase(), f]),
);
const CONSTANTS_BY_LOWER = new Map<string, VbaRuntimeConstant>(
	VBA_RUNTIME_CONSTANTS.map((constant) => [constant.name.toLowerCase(), constant]),
);
const OBJECTS_BY_LOWER = new Map<string, VbaRuntimeObject>(
	VBA_RUNTIME_OBJECTS.map((object) => [object.name.toLowerCase(), object]),
);
const OBJECTS_BY_TYPE_LOWER = new Map<string, VbaRuntimeObject>(
	VBA_RUNTIME_OBJECTS.map((object) => [object.type.toLowerCase(), object]),
);

/** Resolves a built-in VBA runtime function/statement by name (case-insensitive). */
export function resolveRuntimeFunction(
	name: string,
): VbaRuntimeFunction | undefined {
	return BY_LOWER.get(name.toLowerCase());
}

/** Whether this runtime entry may be the target of an explicit `Call` statement. */
export function runtimeAllowsExplicitCall(fn: VbaRuntimeFunction): boolean {
	return fn.explicitCall !== 'forbidden';
}

/** Resolves a built-in VBA runtime constant by name (case-insensitive). */
export function resolveRuntimeConstant(
	name: string,
): VbaRuntimeConstant | undefined {
	return CONSTANTS_BY_LOWER.get(name.toLowerCase());
}

/** Resolves a built-in VBA runtime global object by name, such as `Err`. */
export function resolveRuntimeObject(
	name: string,
): VbaRuntimeObject | undefined {
	return OBJECTS_BY_LOWER.get(name.toLowerCase());
}

/** Resolves a built-in VBA runtime object type, such as `VBA.ErrObject`. */
export function resolveRuntimeObjectType(
	typeName: string,
): VbaRuntimeObject | undefined {
	return OBJECTS_BY_TYPE_LOWER.get(typeName.toLowerCase());
}
