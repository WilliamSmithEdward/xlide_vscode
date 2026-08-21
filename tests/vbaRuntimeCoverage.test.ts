// Tripwire: every function the VBA language reference documents is a name the
// analyzer knows, and the statements that read like bare identifiers are too
// (issue #41).
//
// A name the analyzer does not know is reported as an undeclared variable
// under Option Explicit, so a gap here is a red squiggle on correct code -
// which is how this was found: a user hit Time, Unload, and fmMatchEntry.
//
// The name list is transcribed from the reference index
// (learn.microsoft.com/en-us/office/vba/language/reference/
// functions-visual-basic-for-applications, crawled 2026-08-21), minus its
// four non-function index entries (Derived math, Data type summary, Keywords
// by task, and the language-reference link). Kept as a checked-in list so the
// suite needs no network.
import { describe, expect, it } from 'vitest';
import {
    analyzeModule,
    resolveHostConstant,
    resolveRuntimeConstant,
    resolveRuntimeFunction,
    resolveRuntimeObject,
} from '../src/analyzer';
import { canonicalKeyword } from '../src/analyzer/lexer/keywordTable';
import { vbaRuntimeDescription } from '../src/analyzer/runtime/vbaRuntimeDocs';
import { resolveHover, resolveIdentifierCompletions, resolveSignatureHelp } from '../src/analyzer';
import { getExcelObjectModel } from '../src/analyzer/host/excelObjectModel';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';

const DOCUMENTED_FUNCTIONS: readonly string[] = [
    'Asc', 'Chr', 'CVErr', 'Format', 'Hex', 'Oct', 'Str', 'Val', 'Abs', 'Atn', 'Cos', 'Exp',
    'Int', 'Fix', 'Log', 'Rnd', 'Sgn', 'Sin', 'Sqr', 'Tan', 'Array', 'CallByName', 'Choose',
    'Command', 'CreateObject', 'CurDir', 'Date', 'DateAdd', 'DateDiff', 'DatePart', 'DateSerial',
    'DateValue', 'Day', 'DDB', 'Dir', 'DoEvents', 'Environ', 'EOF', 'Error', 'FileAttr',
    'FileDateTime', 'FileLen', 'Filter', 'FormatCurrency', 'FormatDateTime', 'FormatNumber',
    'FormatPercent', 'FreeFile', 'FV', 'GetAllSettings', 'GetAttr', 'GetObject', 'GetSetting',
    'Hour', 'IIf', 'IMEStatus', 'Input', 'InputBox', 'InStr', 'InStrRev', 'IPmt', 'IRR',
    'IsArray', 'IsDate', 'IsEmpty', 'IsError', 'IsMissing', 'IsNull', 'IsNumeric', 'IsObject',
    'Join', 'LBound', 'LCase', 'Left', 'Len', 'Loc', 'LOF', 'LTrim', 'RTrim', 'Trim', 'MacID',
    'MacScript', 'Mid', 'Minute', 'MIRR', 'Month', 'MonthName', 'MsgBox', 'Now', 'NPer', 'NPV',
    'Partition', 'Pmt', 'PPmt', 'PV', 'QBColor', 'Rate', 'Replace', 'RGB', 'Right', 'Round',
    'Second', 'Seek', 'Shell', 'SLN', 'Space', 'Spc', 'Split', 'StrComp', 'StrConv', 'String',
    'StrReverse', 'Switch', 'SYD', 'Tab', 'Time', 'Timer', 'TimeSerial', 'TimeValue', 'TypeName',
    'UBound', 'UCase', 'VarType', 'Weekday', 'WeekdayName', 'Year',
];

/**
 * Statements whose name reaches the analyzer as a bare identifier, so an
 * unknown one is reported as an undeclared variable. `Name`, `Reset` and
 * `Width #` are deliberately absent: their statement forms already parse, so
 * they are not names the identifier rules ever see, and adding them would
 * weaken a real "Variable not defined" on a bare `x = Name`.
 */
const DOCUMENTED_STATEMENTS: readonly string[] = [
    'Load', 'Unload', 'Time', 'Line', 'Name', 'Reset', 'Width',
];

/**
 * The type-conversion functions (language/concepts/getting-started/
 * type-conversion-functions) and the data types and keywords the reference
 * documents, crawled 2026-08-21.
 */
const TYPE_CONVERSION_FUNCTIONS: readonly string[] = [
    'CBool', 'CByte', 'CCur', 'CDate', 'CDbl', 'CDec', 'CInt', 'CLng', 'CLngLng',
    'CLngPtr', 'CSng', 'CStr', 'CVar', 'CVDate', 'CVErr',
];

const DOCUMENTED_KEYWORDS: readonly string[] = [
    'As', 'Binary', 'ByRef', 'ByVal', 'Date', 'Else', 'Empty', 'Error', 'False',
    'For', 'Friend', 'Get', 'Input', 'Is', 'Len', 'Let', 'Lock', 'Me', 'Mid',
    'New', 'Next', 'Nothing', 'Null', 'On', 'Option', 'Optional', 'ParamArray',
    'Print', 'Private', 'Property', 'PtrSafe', 'Public', 'Resume', 'Seek', 'Set',
    'Static', 'Step', 'String', 'Then', 'Time', 'To', 'True', 'WithEvents',
];

function isKnownName(name: string): boolean {
    return Boolean(resolveRuntimeFunction(name))
        || Boolean(resolveRuntimeObject(name))
        || Boolean(canonicalKeyword(name));
}

describe('VBA language reference coverage', () => {
    it('knows every documented VBA function', () => {
        const unknown = DOCUMENTED_FUNCTIONS.filter((name) => !isKnownName(name));
        expect(unknown).toEqual([]);
        expect(DOCUMENTED_FUNCTIONS).toHaveLength(126);
    });

    it('knows every documented statement', () => {
        expect(DOCUMENTED_STATEMENTS.filter((name) => !isKnownName(name))).toEqual([]);
    });

    it('knows every documented type-conversion function', () => {
        expect(TYPE_CONVERSION_FUNCTIONS.filter((name) => !isKnownName(name))).toEqual([]);
    });

    it('knows every documented keyword', () => {
        expect(DOCUMENTED_KEYWORDS.filter((name) => !isKnownName(name))).toEqual([]);
    });

    it('knows every documented VBA constant', () => {
        // The constants reference, both the enumerations the type library
        // declares and the groups it does not (Keycode, Color, System Color,
        // Form). A sample per family, with the whole set pinned by count.
        const sample = [
            'vbKeyA', 'vbKeyF12', 'vbKeyNumpad0', 'vbKeyEscape', 'vbBlack', 'vbRed',
            'vbButtonFace', 'vb3DLight', 'vbModal', 'vbModeless', 'vbGet', 'vbSet',
            'vbUseCompareOption', 'vbFormControlMenu', 'vbAppTaskManager',
        ];
        const model = getExcelObjectModel();
        const unknown = sample.filter(
            (name) => !resolveRuntimeConstant(name) && !resolveHostConstant(name, model),
        );
        expect(unknown).toEqual([]);
    });

    it('knows the MSForms enum constants in every host', () => {
        for (const getModel of [getExcelObjectModel, getWordObjectModel, getPowerPointObjectModel, getAccessObjectModel]) {
            const model = getModel();
            for (const name of ['fmMatchEntryComplete', 'fmAlignmentLeft', 'fmActionCut']) {
                expect(resolveHostConstant(name, model), `${String(model.hostName)}: ${name}`).toBeDefined();
            }
        }
    });
});

describe('built-ins describe themselves, not just their spelling (issue #41)', () => {
    it('carries a reference summary for every documented function', () => {
        const undocumented = DOCUMENTED_FUNCTIONS.filter((name) => !vbaRuntimeDescription(name));
        expect(undocumented).toEqual([]);
    });

    it('describes the documented statements too', () => {
        const undocumented = DOCUMENTED_STATEMENTS.filter((name) => !vbaRuntimeDescription(name));
        expect(undocumented).toEqual([]);
    });

    it('describes every type-conversion function', () => {
        // CVErr has its own page; the rest share the type-conversion page.
        const undocumented = TYPE_CONVERSION_FUNCTIONS.filter((name) => !vbaRuntimeDescription(name));
        expect(undocumented).toEqual([]);
    });

    it('puts the summary in the hover, the completion and the call tip', () => {
        const source = 'Sub T()\n    x = Left(\"ab\", 1)\nEnd Sub\n';
        const hover = resolveHover(source, source.indexOf('Left') + 2, {});
        expect(hover?.documentation).toContain('characters from the left side of a string');

        const typed = 'Sub T()\n    Le\nEnd Sub\n';
        const item = resolveIdentifierCompletions(typed, typed.indexOf('Le') + 2, { moduleName: 'M' })
            .find((candidate) => candidate.name === 'Left');
        expect(item?.documentation).toContain('characters from the left side of a string');

        const tip = resolveSignatureHelp(source, source.indexOf('ab') + 1, {});
        expect(tip?.documentation).toContain('characters from the left side of a string');
    });

    it('invents no prose for the undocumented hidden built-ins', () => {
        // The reference documents these only in passing on a parent page, so
        // they keep their signature and gain no made-up description.
        for (const name of ['VarPtr', 'StrPtr', 'ObjPtr', 'LenB', 'MidB']) {
            expect(vbaRuntimeDescription(name), name).toBeUndefined();
        }
    });
});

describe('the reported false positives are gone (issue #41)', () => {
    const analyze = (source: string): string[] =>
        analyzeModule(source, { host: 'excel', knownIdentifiers: new Set(['userform1']) })
            .map((finding) => `${finding.code}: ${finding.message}`);

    it.each([
        ['Time function', 'Option Explicit\nSub T()\n    Dim t As Date\n    t = Time\nEnd Sub\n'],
        ['Time call form', 'Option Explicit\nSub T()\n    Dim t As Date\n    t = Time()\nEnd Sub\n'],
        ['Time statement', 'Option Explicit\nSub T()\n    Time = #12:00:00 PM#\nEnd Sub\n'],
        ['Unload statement', 'Option Explicit\nSub T()\n    Unload UserForm1\nEnd Sub\n'],
        ['Load statement', 'Option Explicit\nSub T()\n    Load UserForm1\nEnd Sub\n'],
        ['Line Input', 'Option Explicit\nSub T()\n    Dim s As String\n    Line Input #1, s\nEnd Sub\n'],
        ['FileAttr', 'Option Explicit\nSub T()\n    Dim n As Long\n    n = FileAttr(1, 1)\nEnd Sub\n'],
        ['IMEStatus', 'Option Explicit\nSub T()\n    Dim n As Long\n    n = IMEStatus\nEnd Sub\n'],
        ['MacID', 'Option Explicit\nSub T()\n    Dim s As String\n    s = Dir("x", MacID("TEXT"))\nEnd Sub\n'],
        ['fmMatchEntry', 'Option Explicit\nSub T()\n    Dim x As Long\n    x = fmMatchEntryComplete\nEnd Sub\n'],
    ])('%s analyzes clean', (_label, source) => {
        expect(analyze(source)).toEqual([]);
    });

    it('still reports a name nobody declares', () => {
        const findings = analyze('Option Explicit\nSub T()\n    Dim n As Long\n    n = NotARealName\nEnd Sub\n');
        expect(findings.some((finding) => finding.includes('NotARealName'))).toBe(true);
    });
});
