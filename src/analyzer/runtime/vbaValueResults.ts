// Which of VBA's own functions give a value, not an object or a Variant, to
// `Is` and `TypeOf` (issue #1416, measured in Excel, Word, PowerPoint and
// Access 16.0, build 20430: 2,972 compile-only cases, the same in every host).
//
// The runtime table's return types are not used for this. It types `Right`,
// `Mid`, `LTrim` and others as String where the type library and the VBE
// give a Variant; only their `$` forms return a String. These sets are the
// measured verdicts, so they hold only the names that were measured.
//
// Names are lowercase, with the `$` of a `$` form.

/**
 * `VBA.X(...) Is Nothing` and `TypeOf VBA.X(...) Is Collection` are Type
 * mismatch: X's result is a value. Bare `X(...)` is the same, but see
 * `TYPEOF_KEYWORD_OPERANDS` for bare TypeOf.
 */
export const VALUE_RESULT_FUNCTIONS: ReadonlySet<string> = new Set([
	'asc', 'ascb', 'ascw', 'atn', 'cbool', 'cbyte', 'ccur', 'cdate', 'cdbl', 'chrb$', 'cint', 'clng', 'clnglng',
	'command$', 'cos', 'csng', 'cstr', 'ddb', 'dir', 'doevents', 'eof', 'erl', 'error$', 'exp', 'fileattr',
	'filelen', 'formatcurrency', 'formatdatetime', 'formatnumber', 'formatpercent', 'freefile', 'fv', 'getattr',
	'getsetting', 'hex$', 'imestatus', 'inputbox', 'instrrev', 'ipmt', 'irr', 'isarray', 'isdate', 'isempty',
	'iserror', 'ismissing', 'isnull', 'isnumeric', 'isobject', 'join', 'lcase$', 'left$', 'leftb$', 'loc', 'lof',
	'log', 'ltrim$', 'macid', 'macscript', 'mid$', 'midb$', 'mirr', 'monthname', 'msgbox', 'nper', 'npv',
	'objptr', 'oct$', 'pmt', 'ppmt', 'pv', 'qbcolor', 'rate', 'replace', 'replace$', 'rgb', 'right$', 'rightb$',
	'rnd', 'rtrim$', 'seek', 'shell', 'sin', 'sln', 'space$', 'sqr', 'str$', 'string$', 'strptr', 'strreverse',
	'syd', 'tan', 'time$', 'timer', 'trim$', 'typename', 'ucase$', 'val', 'varptr', 'vartype', 'weekdayname',
]);

/** Bare only: `LBound(a)` and `UBound(a)` are values. `VBA.LBound` is no member of VBA. */
export const BARE_VALUE_RESULT_FUNCTIONS: ReadonlySet<string> = new Set(['lbound', 'ubound']);

/**
 * Bare only, under `Is`: a value when every argument is one (`Len(s)`,
 * `InStr("ab", "a")`, `Abs(n + 1)`), and a Variant when any argument is a
 * Variant, `Null` or an object. Qualified, `VBA.Len(s)`, they are Variants,
 * as the type library declares them, and so is `TypeOf InStr(s, "a")`.
 */
export const ARGUMENT_TYPED_FUNCTIONS: ReadonlySet<string> = new Set([
	'abs', 'fix', 'int', 'sgn', 'len', 'lenb', 'instr', 'instrb', 'strcomp',
]);

/**
 * Bare `TypeOf X ... Is Collection` is a Syntax error for these: the VBE
 * reads them as keywords there. Qualified, `TypeOf VBA.CDbl(1)`, they follow
 * `VALUE_RESULT_FUNCTIONS`, and `CVar`, `CDec`, `Array` and `CVErr` compile.
 */
export const TYPEOF_KEYWORD_OPERANDS: ReadonlySet<string> = new Set([
	'abs', 'array', 'cbool', 'cbyte', 'ccur', 'cdate', 'cdbl', 'cdec', 'cint', 'clng', 'clnglng', 'csng', 'cstr',
	'cvar', 'cverr', 'doevents', 'fix', 'int', 'lbound', 'len', 'lenb', 'seek', 'sgn', 'string$', 'ubound',
]);
