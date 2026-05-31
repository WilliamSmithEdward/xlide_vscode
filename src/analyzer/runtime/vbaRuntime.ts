// Built-in VBA runtime metadata (MS-VBAL Phase 9).
//
// A curated, verified subset of the VBA runtime library: the intrinsic
// functions and statements available in every VBA project regardless of host
// (MsgBox, Left, CLng, Now, Array, RGB, ...). Signatures are transcribed from
// the Microsoft VBA language reference (learn.microsoft.com/office/vba/language)
// and MS-VBAL; they are never LLM-invented.
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
}

/** Explicit metadata for one runtime parameter. Never inferred from names. */
export interface VbaRuntimeParam {
	name: string;
	type?: string;
	optional?: boolean;
	paramArray?: boolean;
}

function fn(
	name: string,
	signature: string,
	returns?: string,
	params?: readonly VbaRuntimeParam[],
): VbaRuntimeFunction {
	return { name, signature, returns, kind: 'function', source: 'verified', params };
}

function stmt(
	name: string,
	signature: string,
	params?: readonly VbaRuntimeParam[],
): VbaRuntimeFunction {
	return { name, signature, kind: 'statement', source: 'verified', params };
}

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
	fn('DoEvents', 'DoEvents() As Integer', 'Integer'),
	fn('CreateObject', 'CreateObject(Class, [ServerName]) As Object', 'Object'),
	fn('GetObject', 'GetObject([PathName], [Class]) As Object', 'Object'),
	stmt('Beep', 'Beep'),
	stmt('Randomize', 'Randomize [Number]'),

	// -- String functions ---------------------------------------------------
	fn('Len', 'Len(Expression) As Long', 'Long'),
	fn('Left', 'Left(String, Length) As String', 'String', [
		{ name: 'String', type: 'String' },
		{ name: 'Length', type: 'Long' },
	]),
	fn('Right', 'Right(String, Length) As String', 'String'),
	fn('Mid', 'Mid(String, Start, [Length]) As String', 'String'),
	fn('Trim', 'Trim(String) As String', 'String'),
	fn('LTrim', 'LTrim(String) As String', 'String'),
	fn('RTrim', 'RTrim(String) As String', 'String'),
	fn('UCase', 'UCase(String) As String', 'String'),
	fn('LCase', 'LCase(String) As String', 'String'),
	fn('Replace', 'Replace(Expression, Find, Replace, [Start = 1], [Count = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String', 'String'),
	fn('InStr', 'InStr([Start], String1, String2, [Compare As VbCompareMethod]) As Long', 'Long'),
	fn('InStrRev', 'InStrRev(StringCheck, StringMatch, [Start = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As Long', 'Long'),
	fn('Split', 'Split(Expression, [Delimiter], [Limit = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String()', 'String()'),
	fn('Join', 'Join(SourceArray, [Delimiter]) As String', 'String'),
	fn('StrComp', 'StrComp(String1, String2, [Compare As VbCompareMethod]) As Integer', 'Integer'),
	fn('StrConv', 'StrConv(String, Conversion As VbStrConv, [LCID]) As String', 'String'),
	fn('Space', 'Space(Number) As String', 'String'),
	fn('Format', 'Format(Expression, [Format], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As String', 'String'),
	fn('Chr', 'Chr(CharCode) As String', 'String'),
	fn('ChrW', 'ChrW(CharCode) As String', 'String'),
	fn('Asc', 'Asc(String) As Integer', 'Integer'),
	fn('AscW', 'AscW(String) As Integer', 'Integer'),

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
	fn('Array', 'Array(ArgList) As Variant', 'Variant'),
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
	fn('CallByName', 'CallByName(Object, ProcName As String, CallType As VbCallType, [Args]) As Variant', 'Variant'),

	// -- Additional string functions ---------------------------------------
	fn('StrReverse', 'StrReverse(Expression) As String', 'String'),
	fn('LenB', 'LenB(Expression) As Long', 'Long'),
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
	fn('CurDir', 'CurDir([Drive]) As String', 'String'),
	stmt('ChDir', 'ChDir Path'),
	stmt('ChDrive', 'ChDrive Drive'),
	stmt('MkDir', 'MkDir Path'),
	stmt('RmDir', 'RmDir Path'),
	stmt('Kill', 'Kill PathName'),
	stmt('FileCopy', 'FileCopy Source, Destination'),
	stmt('SetAttr', 'SetAttr PathName, Attributes As VbFileAttribute'),

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

const BY_LOWER = new Map<string, VbaRuntimeFunction>(
	VBA_RUNTIME_FUNCTIONS.map((f) => [f.name.toLowerCase(), f]),
);

/** Resolves a built-in VBA runtime function/statement by name (case-insensitive). */
export function resolveRuntimeFunction(
	name: string,
): VbaRuntimeFunction | undefined {
	return BY_LOWER.get(name.toLowerCase());
}
